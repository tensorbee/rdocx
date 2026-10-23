//! Pure evaluation of Word fields against an explicit document context.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::path::Path;
use std::sync::Arc;

use oxml_core::Length;
use oxml_core::custom_properties::CustomPropertyValue;
use oxml_core::xml::{XmlLexicalError, validate_strict_xml_1_0};
use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::comments::CT_Comments;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Body, CT_Document, CT_SectPr};
use rdocx_oxml::drawing::{CT_Drawing, CT_Inline};
use rdocx_oxml::footnotes::CT_Footnotes;
use rdocx_oxml::header_footer::CT_HdrFtr;
use rdocx_oxml::namespace::{R_NS, W_NS, matches_local_name};
use rdocx_oxml::numbering::ST_LvlSuffix;
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::revision::{CT_Revision, RevisionContent, RevisionKind};
use rdocx_oxml::shared::{ST_SectionType, ST_TabJc};
use rdocx_oxml::styles::{CT_Styles, StyleType};
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::CommentRangeMarker;
use rdocx_oxml::text::{
    CT_P, CT_R, CT_Text, Field, FieldArgument, FieldInstruction, RunContent,
    hyperlink_revision_index,
};

pub use rdocx_oxml::text::{LegacyFormFieldKind, LegacyFormFieldValue};

use crate::document::{
    DocumentIdentifiers, FragmentConflictPolicy, StoryKind, prepare_physical_story_projection,
    uniquify_drawing_ids_in_xml,
};
use crate::{Document, Error, Result, style};

/// A nonprinting index marker and its page-number formatting.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct IndexEntry {
    pub levels: Vec<String>,
    pub identifier: Option<String>,
    pub page_range_bookmark: Option<String>,
    pub cross_reference: Option<String>,
    pub bold_page_numbers: bool,
    pub italic_page_numbers: bool,
}

/// A nonprinting authority occurrence. Short citations are category-local keys.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AuthorityEntry {
    pub long_citation: String,
    pub short_citation: String,
    pub category: u8,
    pub page_range_bookmark: Option<String>,
    pub bold_page_numbers: bool,
    pub italic_page_numbers: bool,
}

/// Native INDEX switches and the cached entry leader.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexOptions {
    pub identifier: Option<String>,
    pub heading_separator: Option<String>,
    pub entry_page_separator: String,
    pub page_separator: String,
    pub range_separator: String,
    pub run_in: bool,
    pub leader: crate::TabLeader,
    pub hyperlink: bool,
}

impl Default for IndexOptions {
    fn default() -> Self {
        Self {
            identifier: None,
            heading_separator: None,
            entry_page_separator: ", ".into(),
            page_separator: ", ".into(),
            range_separator: "–".into(),
            run_in: false,
            leader: crate::TabLeader::Dot,
            hyperlink: false,
        }
    }
}

/// A caption-selected TOC, independent of the TOC sequence page-prefix switch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfFiguresOptions {
    pub label: String,
    pub include_label_and_number: bool,
    pub hyperlink: bool,
    pub leader: crate::TabLeader,
    pub entry_page_separator: String,
}

impl Default for TableOfFiguresOptions {
    fn default() -> Self {
        Self {
            label: "Figure".into(),
            include_label_and_number: true,
            hyperlink: true,
            leader: crate::TabLeader::Dot,
            entry_page_separator: "\t".into(),
        }
    }
}

/// Native category, heading and passim selection for authorities.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesOptions {
    pub category: Option<u8>,
    pub include_category_headings: bool,
    pub use_passim: bool,
    pub entry_page_separator: String,
    pub page_separator: String,
    pub range_separator: String,
    pub leader: crate::TabLeader,
}

impl Default for TableOfAuthoritiesOptions {
    fn default() -> Self {
        Self {
            category: None,
            include_category_headings: true,
            use_passim: true,
            entry_page_separator: "\t".into(),
            page_separator: ", ".into(),
            range_separator: "–".into(),
            leader: crate::TabLeader::Dot,
        }
    }
}

/// Counts of materialized entries and stable retained-cache diagnostics.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct GeneratedTablesReport {
    pub index_entries: usize,
    pub figure_entries: usize,
    pub authority_entries: usize,
    pub bookmark_count: usize,
    pub diagnostics: Vec<String>,
}

impl Document {
    /// Insert a checked complex XE marker without a printing cached result.
    pub fn insert_index_entry(
        &mut self,
        position: &crate::StoryRunPosition,
        entry: &IndexEntry,
    ) -> Result<()> {
        if entry.levels.is_empty()
            || entry
                .levels
                .iter()
                .any(|level| level.is_empty() || level.contains(':'))
        {
            return Err(Error::Other(
                "index hierarchy requires nonempty components without colons".into(),
            ));
        }
        if entry.page_range_bookmark.is_some() && entry.cross_reference.is_some() {
            return Err(Error::Other(
                "index marker cannot combine a range and cross-reference".into(),
            ));
        }
        let mut switches = Vec::new();
        for (name, value) in [
            ("f", &entry.identifier),
            ("r", &entry.page_range_bookmark),
            ("t", &entry.cross_reference),
        ] {
            if let Some(value) = value {
                generated_nonempty_operand(name, value)?;
                if name == "r" {
                    validate_reference_name(value)?;
                }
                switches.push(field_option_switch(name, Some(value.clone())));
            }
        }
        for (enabled, name) in [
            (entry.bold_page_numbers, "b"),
            (entry.italic_page_numbers, "i"),
        ] {
            if enabled {
                switches.push(field_option_switch(name, None));
            }
        }
        let instruction = FieldInstruction::new(
            "XE",
            vec![FieldArgument::Text(entry.levels.join(":"))],
            switches,
        )?;
        self.insert_checked_story_field(position, instruction)
    }

    /// Insert a checked complex TA occurrence with a category-local short key.
    pub fn insert_authority_entry(
        &mut self,
        position: &crate::StoryRunPosition,
        entry: &AuthorityEntry,
    ) -> Result<()> {
        generated_category(entry.category)?;
        generated_nonempty_operand("long citation", &entry.long_citation)?;
        generated_nonempty_operand("short citation", &entry.short_citation)?;
        let mut switches = vec![
            field_option_switch("l", Some(entry.long_citation.clone())),
            field_option_switch("s", Some(entry.short_citation.clone())),
            field_option_switch("c", Some(entry.category.to_string())),
        ];
        if let Some(bookmark) = &entry.page_range_bookmark {
            validate_reference_name(bookmark)?;
            switches.push(field_option_switch("r", Some(bookmark.clone())));
        }
        for (enabled, name) in [
            (entry.bold_page_numbers, "b"),
            (entry.italic_page_numbers, "i"),
        ] {
            if enabled {
                switches.push(field_option_switch(name, None));
            }
        }
        self.insert_checked_story_field(
            position,
            FieldInstruction::new("TA", Vec::new(), switches)?,
        )
    }

    /// Insert a dynamic INDEX at a checked story boundary.
    pub fn insert_index(
        &mut self,
        before: &crate::ContentLocation,
        options: &IndexOptions,
    ) -> Result<crate::ContentLocation> {
        let mut switches = vec![field_option_switch("z", Some("1033".into()))];
        for (name, value) in [
            ("f", &options.identifier),
            ("h", &options.heading_separator),
        ] {
            if let Some(value) = value {
                generated_nonempty_operand(name, value)?;
                switches.push(field_option_switch(name, Some(value.clone())));
            }
        }
        for (name, value) in [
            ("e", &options.entry_page_separator),
            ("l", &options.page_separator),
            ("g", &options.range_separator),
        ] {
            switches.push(field_option_switch(name, Some(value.clone())));
        }
        if options.run_in {
            switches.push(field_option_switch("r", None));
        }
        self.insert_generated_table_fields(
            before,
            &[FieldInstruction::new("INDEX", Vec::new(), switches)?],
            options.leader,
            options.hyperlink,
        )
    }

    /// Insert Word's caption-selected TOC instruction.
    pub fn insert_table_of_figures(
        &mut self,
        before: &crate::ContentLocation,
        options: &TableOfFiguresOptions,
    ) -> Result<crate::ContentLocation> {
        generated_nonempty_operand("caption label", &options.label)?;
        let selector = if options.include_label_and_number {
            "c"
        } else {
            "a"
        };
        let mut switches = vec![
            field_option_switch(selector, Some(options.label.clone())),
            field_option_switch("p", Some(options.entry_page_separator.clone())),
        ];
        if options.hyperlink {
            switches.push(field_option_switch("h", None));
        }
        self.insert_generated_table_fields(
            before,
            &[FieldInstruction::new("TOC", Vec::new(), switches)?],
            options.leader,
            options.hyperlink,
        )
    }

    /// Insert every populated numbered category atomically when category is None.
    pub fn insert_table_of_authorities(
        &mut self,
        before: &crate::ContentLocation,
        options: &TableOfAuthoritiesOptions,
    ) -> Result<crate::ContentLocation> {
        let categories = if let Some(category) = options.category {
            generated_category(category)?;
            vec![category]
        } else {
            generated_authority_categories(self)?
        };
        if categories.is_empty() {
            return Err(Error::Other("no populated authority category".into()));
        }
        let mut instructions = Vec::new();
        for category in categories {
            let mut switches = vec![field_option_switch("c", Some(category.to_string()))];
            for (enabled, name) in [
                (options.include_category_headings, "h"),
                (options.use_passim, "p"),
            ] {
                if enabled {
                    switches.push(field_option_switch(name, None));
                }
            }
            for (name, value) in [
                ("e", &options.entry_page_separator),
                ("l", &options.page_separator),
                ("g", &options.range_separator),
            ] {
                switches.push(field_option_switch(name, Some(value.clone())));
            }
            instructions.push(FieldInstruction::new("TOA", Vec::new(), switches)?);
        }
        self.insert_generated_table_fields(before, &instructions, options.leader, false)
    }

    fn insert_generated_table_fields(
        &mut self,
        before: &crate::ContentLocation,
        instructions: &[FieldInstruction],
        leader: crate::TabLeader,
        hyperlink: bool,
    ) -> Result<crate::ContentLocation> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let index = before
            .index_path()
            .first()
            .copied()
            .unwrap_or(candidate.story_items(before.story())?.len());
        // Keep the original checked sibling boundary while inserting in native category order.
        let mut original_index = index;
        let mut boundary = before.clone();
        for instruction in instructions {
            let previous_count = candidate.story_items(boundary.story())?.len();
            let mut paragraph = CT_P::new();
            let mut properties = CT_PPr {
                style_id: Some(
                    match instruction.name.as_str() {
                        "INDEX" => "Index1",
                        "TOA" => "TableofAuthorities",
                        _ => "TableofFigures",
                    }
                    .into(),
                ),
                tabs: Some(rdocx_oxml::borders::CT_Tabs {
                    tabs: vec![rdocx_oxml::borders::CT_TabStop {
                        val: ST_TabJc::Right,
                        pos: rdocx_oxml::units::Twips(9360),
                        leader: Some(generated_tab_leader(leader)),
                        source_occurrence: None,
                    }],
                }),
                ..Default::default()
            };
            if hyperlink && instruction.name == "INDEX" {
                properties.rpr = Some(CT_RPr {
                    style_id: Some("Hyperlink".into()),
                    ..Default::default()
                });
            }
            paragraph.properties = Some(properties);
            let mut run = CT_R::new("");
            if hyperlink {
                run.properties = Some(CT_RPr {
                    style_id: Some("Hyperlink".into()),
                    ..Default::default()
                });
            }
            run.content = vec![RunContent::Field(Field::from_instruction(
                instruction.clone(),
                rdocx_oxml::text::FieldForm::Complex,
                Vec::new(),
            )?)];
            paragraph.runs.push(run);
            candidate.insert_content(&boundary, crate::ContentFragment::paragraph(paragraph)?)?;
            let story = candidate
                .stories()?
                .into_iter()
                .find(|story| {
                    story.kind() == before.story().kind()
                        && story.part_name() == before.story().part_name()
                        && story.owner_index() == before.story().owner_index()
                })
                .ok_or_else(|| Error::Other("generated table story disappeared".into()))?;
            let new_count = candidate.story_items(&story)?.len();
            original_index += new_count
                .checked_sub(previous_count)
                .ok_or_else(|| Error::Other("generated insertion removed story items".into()))?;
            boundary = if before.index_path().is_empty() {
                crate::ContentLocation::end(story)
            } else {
                crate::ContentLocation::new(story, before.item_kind(), vec![original_index])
            };
        }
        let reopened = candidate.prepare_and_reopen_staged()?;
        let story = reopened
            .stories()?
            .into_iter()
            .find(|story| {
                story.kind() == before.story().kind()
                    && story.part_name() == before.story().part_name()
                    && story.owner_index() == before.story().owner_index()
            })
            .ok_or_else(|| Error::Other("inserted generated story disappeared".into()))?;
        let first = reopened
            .story_items(&story)?
            .get(index)
            .ok_or_else(|| Error::Other("inserted generated table has no checked location".into()))?
            .location()
            .clone();
        self.commit_staged_mutation(reopened);
        Ok(first)
    }
}

fn generated_nonempty_operand(name: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::Other(format!("{name} must not be empty")));
    }
    oxml_core::xml::reject_non_xml_characters(name, value)?;
    Ok(())
}

fn generated_category(category: u8) -> Result<()> {
    if !(1..=16).contains(&category) {
        return Err(Error::Other(
            "authority category must be 1 through 16".into(),
        ));
    }
    Ok(())
}

fn generated_tab_leader(leader: crate::TabLeader) -> rdocx_oxml::shared::ST_TabLeader {
    match leader {
        crate::TabLeader::None => rdocx_oxml::shared::ST_TabLeader::None,
        crate::TabLeader::Dot => rdocx_oxml::shared::ST_TabLeader::Dot,
        crate::TabLeader::Hyphen => rdocx_oxml::shared::ST_TabLeader::Hyphen,
        crate::TabLeader::Underscore => rdocx_oxml::shared::ST_TabLeader::Underscore,
        crate::TabLeader::Heavy => rdocx_oxml::shared::ST_TabLeader::Heavy,
        crate::TabLeader::MiddleDot => rdocx_oxml::shared::ST_TabLeader::MiddleDot,
    }
}

/// Switches governing one authored sequence field.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SequenceOptions {
    pub restart: Option<i64>,
    pub repeat: bool,
    pub hidden: bool,
    pub restart_heading: Option<u8>,
    pub format: Option<String>,
}

/// Caption content and its three bookmark targets.
#[derive(Debug, Clone, PartialEq)]
pub struct CaptionOptions {
    pub label: String,
    pub text: String,
    pub bookmark: String,
    pub sequence: SequenceOptions,
    pub separator: String,
    pub properties: Option<CT_PPr>,
}

impl Default for CaptionOptions {
    fn default() -> Self {
        Self {
            label: "Figure".into(),
            text: String::new(),
            bookmark: "Caption".into(),
            sequence: SequenceOptions::default(),
            separator: ": ".into(),
            properties: None,
        }
    }
}

/// Checked caption location and the actual allocated bookmark names.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptionTarget {
    pub paragraph: crate::ContentLocation,
    pub entire_caption: String,
    pub label_and_number: String,
    pub number: String,
}

/// Numbering context requested by a cross-reference.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum CrossReferenceNumber {
    #[default]
    Text,
    Level,
    Relative,
    FullContext,
}

/// Checked REF switches, including typed note copying.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CrossReferenceOptions {
    pub number: CrossReferenceNumber,
    pub position: bool,
    pub hyperlink: bool,
    pub omit_non_numeric_text: bool,
    pub delimiter: Option<String>,
    pub copy_referenced_notes: bool,
}

/// One legacy form field and its stable story-part identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyFormFieldInfo {
    pub source_part: String,
    pub ordinal: usize,
    pub name: Option<String>,
    pub enabled: bool,
    pub calculate_on_exit: bool,
    pub kind: LegacyFormFieldKind,
    pub value: LegacyFormFieldValue,
    pub choices: Vec<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LegacyStoryKind {
    Header,
    Footer,
    Footnotes,
    Endnotes,
}

impl LegacyStoryKind {
    fn package_kind(self) -> PackageStoryKind {
        match self {
            Self::Header => PackageStoryKind::Header,
            Self::Footer => PackageStoryKind::Footer,
            Self::Footnotes => PackageStoryKind::Footnotes,
            Self::Endnotes => PackageStoryKind::Endnotes,
        }
    }

    fn content_type(self) -> &'static str {
        match self {
            Self::Header => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
            }
            Self::Footer => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
            }
            Self::Footnotes => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
            }
            Self::Endnotes => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
            }
        }
    }
}

impl Document {
    /// Insert a dynamic sequence field at a checked accepted run boundary.
    pub fn insert_sequence(
        &mut self,
        position: &crate::StoryRunPosition,
        identifier: &str,
        options: &SequenceOptions,
    ) -> Result<()> {
        let instruction = sequence_instruction(identifier, options)?;
        self.insert_checked_story_field(position, instruction)
    }

    /// Insert a dynamic cross-reference without flattening its target.
    pub fn insert_cross_reference(
        &mut self,
        position: &crate::StoryRunPosition,
        bookmark: &str,
        options: &CrossReferenceOptions,
    ) -> Result<()> {
        validate_reference_name(bookmark)?;
        let mut switches = Vec::new();
        let number = match options.number {
            CrossReferenceNumber::Text => None,
            CrossReferenceNumber::Level => Some("n"),
            CrossReferenceNumber::Relative => Some("r"),
            CrossReferenceNumber::FullContext => Some("w"),
        };
        switches.extend(number.map(|name| field_option_switch(name, None)));
        for (enabled, name) in [
            (options.position, "p"),
            (options.hyperlink, "h"),
            (options.omit_non_numeric_text, "t"),
            (options.copy_referenced_notes, "f"),
        ] {
            if enabled {
                switches.push(field_option_switch(name, None));
            }
        }
        if let Some(delimiter) = &options.delimiter {
            switches.push(field_option_switch("d", Some(delimiter.clone())));
        }
        let instruction =
            FieldInstruction::new("REF", vec![FieldArgument::Text(bookmark.into())], switches)?;
        self.insert_checked_story_field(position, instruction)
    }

    pub(crate) fn insert_checked_story_field(
        &mut self,
        position: &crate::StoryRunPosition,
        instruction: FieldInstruction,
    ) -> Result<()> {
        let field = Field::from_instruction(
            instruction,
            rdocx_oxml::text::FieldForm::Complex,
            Vec::new(),
        )?;
        self.insert_field(position, field)
    }

    /// Insert a field at one checked physical run boundary without replacing its owner.
    pub fn insert_field(&mut self, position: &crate::StoryRunPosition, field: Field) -> Result<()> {
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Field(field)];
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        candidate.insert_story_run_staged(position, run)?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(())
    }

    /// Insert a caption and allocate separate whole, label and number targets.
    pub fn insert_caption(
        &mut self,
        destination: &crate::ContentLocation,
        options: &CaptionOptions,
    ) -> Result<CaptionTarget> {
        let instruction = sequence_instruction(&options.label, &options.sequence)?;
        validate_reference_name(&options.bookmark)?;
        oxml_core::xml::reject_non_xml_characters("caption text", &options.text)?;
        oxml_core::xml::reject_non_xml_characters("caption separator", &options.separator)?;
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let mut names = caption_bookmark_names(&candidate)?;
        let mut allocated = Vec::new();
        for suffix in ["", "_Label", "_Number"] {
            let base = format!("{}{suffix}", options.bookmark);
            validate_reference_name(&base)?;
            let mut name = base.clone();
            let mut counter = 1u32;
            while names.contains(&name) {
                name = format!("{base}_{counter}");
                validate_reference_name(&name)?;
                counter = counter
                    .checked_add(1)
                    .ok_or_else(|| Error::Other("caption bookmark names exhausted".into()))?;
            }
            names.insert(name.clone());
            allocated.push(name);
        }
        candidate
            .identifiers
            .observe_package_graph(&candidate.package)?;
        let mut paragraph = CT_P::new();
        paragraph.properties = options.properties.clone();
        paragraph.add_run(&format!("{} ", options.label));
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Field(Field::from_instruction(
            instruction,
            rdocx_oxml::text::FieldForm::Complex,
            Vec::new(),
        )?)];
        paragraph.runs.push(run);
        paragraph.add_run(&format!("{}{}", options.separator, options.text));
        for (name, start, end) in [
            (&allocated[0], 0, 3),
            (&allocated[1], 0, 2),
            (&allocated[2], 1, 2),
        ] {
            let id = candidate.identifiers.reserve_bookmark_id()?;
            paragraph
                .anchor_accepted_range(
                    Some(start),
                    Some(end),
                    rdocx_oxml::text::RangeAnchor::Bookmark { id, name },
                )
                .map_err(|error| Error::Other(format!("caption target: {error}")))?;
        }
        let index = destination
            .index_path()
            .first()
            .copied()
            .unwrap_or(candidate.story_items(destination.story())?.len());
        candidate.insert_content(destination, crate::ContentFragment::paragraph(paragraph)?)?;
        let story = candidate
            .stories()?
            .into_iter()
            .find(|story| {
                story.kind() == destination.story().kind()
                    && story.part_name() == destination.story().part_name()
                    && story.owner_index() == destination.story().owner_index()
            })
            .ok_or_else(|| Error::Other("caption story disappeared after insertion".into()))?;
        let target = CaptionTarget {
            paragraph: crate::ContentLocation::new(
                story,
                crate::StoryItemKind::Paragraph,
                vec![index],
            ),
            entire_caption: allocated[0].clone(),
            label_and_number: allocated[1].clone(),
            number: allocated[2].clone(),
        };
        candidate.story_ranges()?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(target)
    }
}

fn field_option_switch(name: &str, argument: Option<String>) -> rdocx_oxml::text::FieldSwitch {
    rdocx_oxml::text::FieldSwitch {
        name: name.into(),
        argument: argument.map(FieldArgument::Text),
    }
}

fn validate_reference_name(name: &str) -> Result<()> {
    let mut characters = name.chars();
    if name.len() > 40
        || !characters
            .next()
            .is_some_and(|value| value.is_ascii_alphabetic() || value == '_')
        || !characters.all(|value| value.is_ascii_alphanumeric() || value == '_')
    {
        return Err(Error::Other(format!("invalid Word reference name {name}")));
    }
    Ok(())
}

fn sequence_instruction(identifier: &str, options: &SequenceOptions) -> Result<FieldInstruction> {
    validate_reference_name(identifier)?;
    if !identifier.as_bytes()[0].is_ascii_alphabetic() {
        return Err(Error::Other(
            "SEQ identifier must begin with a letter".into(),
        ));
    }
    if options.repeat && options.restart.is_some() {
        return Err(Error::Other(
            "SEQ repeat and restart are mutually exclusive".into(),
        ));
    }
    if options
        .restart_heading
        .is_some_and(|level| !(1..=9).contains(&level))
    {
        return Err(Error::Other("SEQ heading level must be 1 through 9".into()));
    }
    let mut switches = Vec::new();
    for (enabled, name) in [(options.repeat, "c"), (options.hidden, "h")] {
        if enabled {
            switches.push(field_option_switch(name, None));
        }
    }
    if let Some(restart) = options.restart {
        switches.push(field_option_switch("r", Some(restart.to_string())));
    }
    if let Some(level) = options.restart_heading {
        switches.push(field_option_switch("s", Some(level.to_string())));
    }
    if let Some(format) = &options.format {
        switches.push(field_option_switch("*", Some(format.clone())));
    }
    let instruction = FieldInstruction::new(
        "SEQ",
        vec![FieldArgument::Text(identifier.into())],
        switches,
    )?;
    rdocx_layout::engine::format_numeric_field_general(&instruction, "1").map_err(Error::Other)?;
    Ok(instruction)
}

fn caption_bookmark_names(document: &Document) -> Result<HashSet<String>> {
    let mut names = HashSet::new();
    for (part_name, xml) in &document.package.parts {
        if !document
            .package
            .content_types
            .content_type_for(part_name)
            .is_some_and(|kind| kind.ends_with("xml"))
        {
            continue;
        }
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::Other(format!("caption bookmark inventory: {error}")))?;
            let word = namespace_is_word(&namespace);
            match event {
                Event::Start(element) | Event::Empty(element)
                    if word && element.local_name().as_ref() == b"bookmarkStart" =>
                {
                    if let Some((_, name)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"name",
                        AttributeNamespace::Word,
                    )? {
                        names.insert(name);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    Ok(names)
}

impl Document {
    /// Inventory typed legacy form fields in deterministic story-part order.
    pub fn legacy_form_fields(&self) -> Result<Vec<LegacyFormFieldInfo>> {
        let mut fields = Vec::new();
        let mut paragraphs = Vec::new();
        collect_body_paragraphs(&self.document.body, &mut paragraphs);
        append_legacy_form_infos(&mut fields, &self.doc_part_name, &paragraphs)?;

        for (part_name, kind, xml) in legacy_story_parts(self)? {
            let story = legacy_story_paragraphs(&xml, kind.package_kind())?;
            let paragraphs = story
                .iter()
                .map(|paragraph| &paragraph.paragraph)
                .collect::<Vec<_>>();
            append_legacy_form_infos(&mut fields, &part_name, &paragraphs)?;
        }
        Ok(fields)
    }

    /// Set the typed value of one existing legacy form field atomically.
    pub fn set_legacy_form_field_value(
        &mut self,
        source_part: &str,
        ordinal: usize,
        value: LegacyFormFieldValue,
    ) -> Result<LegacyFormFieldInfo> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        if source_part == candidate.doc_part_name {
            let mut remaining = ordinal;
            if !set_nth_legacy_form_in_body(
                &mut candidate.document.body,
                &mut remaining,
                value.clone(),
            )? {
                return Err(Error::Other("stale legacy form identity".to_owned()));
            }
        } else {
            let (_, kind, xml) = legacy_story_parts(&candidate)?
                .into_iter()
                .find(|(part_name, _, _)| part_name == source_part)
                .ok_or_else(|| Error::Other("stale legacy form story part".to_owned()))?;
            let mut story = legacy_story_paragraphs(&xml, kind.package_kind())?;
            let mut remaining = ordinal;
            let mut changed = false;
            for paragraph in &mut story {
                if set_nth_legacy_form_in_paragraph(
                    &mut paragraph.paragraph,
                    &mut remaining,
                    value.clone(),
                )? {
                    changed = true;
                    break;
                }
            }
            if !changed {
                return Err(Error::Other("stale legacy form identity".to_owned()));
            }
            let updated = patch_legacy_story_field_sources(&xml, &story, false)?;
            candidate.package.set_part(source_part, updated);
        }

        let reopened = candidate.prepare_and_reopen_staged()?;
        let result = reopened
            .legacy_form_fields()?
            .into_iter()
            .find(|field| field.source_part == source_part && field.ordinal == ordinal)
            .ok_or_else(|| {
                Error::Other("legacy form identity did not survive reopen".to_owned())
            })?;
        if result.value != value {
            return Err(Error::Other(
                "legacy form value did not survive reopen".to_owned(),
            ));
        }
        self.commit_staged_mutation(reopened);
        Ok(result)
    }
}

fn field_story_parts(
    document: &Document,
    include_comments: bool,
) -> Vec<(String, PackageStoryKind, Vec<u8>)> {
    let mut parts = Vec::new();
    for (is_header, kind) in [
        (true, PackageStoryKind::Header),
        (false, PackageStoryKind::Footer),
    ] {
        parts.extend(
            referenced_header_footer_parts(document, is_header, true)
                .into_iter()
                .map(|(part, xml)| (part, kind, xml)),
        );
    }
    for (relationship, kind) in [
        (rel_types::FOOTNOTES, PackageStoryKind::Footnotes),
        (rel_types::ENDNOTES, PackageStoryKind::Endnotes),
    ] {
        parts.extend(
            relationship_parts(document, relationship)
                .into_iter()
                .map(|(part, xml)| (part, kind, xml)),
        );
    }
    if include_comments {
        parts.extend(
            relationship_parts(document, rel_types::COMMENTS)
                .into_iter()
                .map(|(part, xml)| (part, PackageStoryKind::Comments, xml)),
        );
    }
    parts
}

fn legacy_story_parts(document: &Document) -> Result<Vec<(String, LegacyStoryKind, Vec<u8>)>> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Ok(Vec::new());
    };
    let story_relationship_types = [
        rel_types::HEADER,
        rel_types::FOOTER,
        rel_types::FOOTNOTES,
        rel_types::ENDNOTES,
    ];
    let mut relationship_ids = HashSet::new();
    if relationships
        .items
        .iter()
        .any(|relationship| story_relationship_types.contains(&relationship.rel_type.as_str()))
        && relationships
            .items
            .iter()
            .any(|relationship| !relationship_ids.insert(&relationship.id))
    {
        return Err(Error::Other(
            "duplicate legacy form story relationship id".to_owned(),
        ));
    }
    let mut parts = Vec::new();
    for (relationship_type, kind) in [
        (rel_types::HEADER, LegacyStoryKind::Header),
        (rel_types::FOOTER, LegacyStoryKind::Footer),
        (rel_types::FOOTNOTES, LegacyStoryKind::Footnotes),
        (rel_types::ENDNOTES, LegacyStoryKind::Endnotes),
    ] {
        let matching_relationships = relationships
            .items
            .iter()
            .filter(|relationship| {
                relationship.rel_type == relationship_type
                    && crate::document::relationship_is_internal(relationship)
            })
            .collect::<Vec<_>>();
        if matches!(kind, LegacyStoryKind::Footnotes | LegacyStoryKind::Endnotes)
            && matching_relationships.len() > 1
        {
            return Err(Error::Other(format!(
                "legacy form story has {} {relationship_type} relationships, expected at most one",
                matching_relationships.len()
            )));
        }
        for relationship in matching_relationships {
            crate::building_block::validate_internal_target(
                &document.doc_part_name,
                &relationship.target,
            )?;
            let part_name =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            if document.package.content_types.override_for(&part_name) != Some(kind.content_type())
            {
                return Err(Error::Other(
                    "legacy form story requires its exact content type override".to_owned(),
                ));
            }
            if parts.iter().any(|(existing_part, existing_kind, _)| {
                existing_part == &part_name && existing_kind != &kind
            }) {
                return Err(Error::Other(
                    "legacy form story part has conflicting relationship roles".to_owned(),
                ));
            }
            let xml = document.package.get_part(&part_name).ok_or_else(|| {
                Error::Other("legacy form story relationship target is missing".to_owned())
            })?;
            parts.push((part_name, kind, xml.to_vec()));
        }
    }
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    parts.dedup_by(|left, right| left.0 == right.0);
    Ok(parts)
}

fn append_legacy_form_infos(
    output: &mut Vec<LegacyFormFieldInfo>,
    source_part: &str,
    paragraphs: &[&CT_P],
) -> Result<()> {
    let mut ordinal = 0usize;
    for paragraph in paragraphs {
        append_legacy_form_infos_in_paragraph(output, source_part, &mut ordinal, paragraph)?;
    }
    Ok(())
}

fn append_legacy_form_infos_in_paragraph(
    output: &mut Vec<LegacyFormFieldInfo>,
    source_part: &str,
    ordinal: &mut usize,
    paragraph: &CT_P,
) -> Result<()> {
    for boundary in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter()
            .filter(|(position, _, _, _)| *position == boundary)
        {
            append_legacy_form_infos_in_control(output, source_part, ordinal, control)?;
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            append_legacy_form_infos_in_run(output, source_part, ordinal, run)?;
        }
    }
    Ok(())
}

fn append_legacy_form_infos_in_run(
    output: &mut Vec<LegacyFormFieldInfo>,
    source_part: &str,
    ordinal: &mut usize,
    run: &CT_R,
) -> Result<()> {
    for content in &run.content {
        if let RunContent::Field(field) = content {
            append_legacy_form_field_info(output, source_part, ordinal, field)?;
        }
    }
    Ok(())
}

fn append_legacy_form_infos_in_control(
    output: &mut Vec<LegacyFormFieldInfo>,
    source_part: &str,
    ordinal: &mut usize,
    control: &CT_Sdt,
) -> Result<()> {
    let inline_fields = control.inline_legacy_form_fields_with_content_indices()?;
    let mut inline_index = 0usize;
    for (content_index, content) in control.content.iter().enumerate() {
        while inline_fields
            .get(inline_index)
            .is_some_and(|(position, _)| *position == content_index)
        {
            append_one_legacy_form_field_info(
                output,
                source_part,
                ordinal,
                &inline_fields[inline_index].1,
            );
            inline_index += 1;
        }
        match content {
            SdtContent::Run(_) => {}
            SdtContent::ContentControl(control) => {
                append_legacy_form_infos_in_control(output, source_part, ordinal, control)?
            }
            SdtContent::Paragraph(paragraph) => {
                append_legacy_form_infos_in_paragraph(output, source_part, ordinal, paragraph)?
            }
            SdtContent::Table(_)
            | SdtContent::Row(_)
            | SdtContent::Cell(_)
            | SdtContent::RawXml(_) => {}
        }
    }
    if inline_index != inline_fields.len() {
        return Err(Error::Other(
            "inline legacy form source boundary is invalid".to_owned(),
        ));
    }
    Ok(())
}

fn append_legacy_form_field_info(
    output: &mut Vec<LegacyFormFieldInfo>,
    source_part: &str,
    ordinal: &mut usize,
    field: &Field,
) -> Result<()> {
    field.validate_legacy_form_owner()?;
    append_one_legacy_form_field_info(output, source_part, ordinal, field);
    for nested in field.nested_fields_in_source_order() {
        append_legacy_form_field_info(output, source_part, ordinal, nested)?;
    }
    Ok(())
}

fn append_one_legacy_form_field_info(
    output: &mut Vec<LegacyFormFieldInfo>,
    source_part: &str,
    ordinal: &mut usize,
    field: &Field,
) {
    if let Some(form) = &field.legacy_form {
        output.push(LegacyFormFieldInfo {
            source_part: source_part.to_owned(),
            ordinal: *ordinal,
            name: form.name.clone(),
            enabled: form.enabled,
            calculate_on_exit: form.calculate_on_exit,
            kind: form.kind,
            value: form.value.clone(),
            choices: form.choices.clone(),
        });
        *ordinal += 1;
    }
}

fn set_nth_legacy_form_in_body(
    body: &mut CT_Body,
    remaining: &mut usize,
    value: LegacyFormFieldValue,
) -> Result<bool> {
    for content in &mut body.content {
        let changed = match content {
            BodyContent::Paragraph(paragraph) => {
                set_nth_legacy_form_in_paragraph(paragraph, remaining, value.clone())?
            }
            BodyContent::Table(table) => {
                set_nth_legacy_form_in_table(table, remaining, value.clone())?
            }
            BodyContent::ContentControl(control) => {
                set_nth_legacy_form_in_control(control, remaining, value.clone())?
            }
            BodyContent::RawXml(_) => false,
        };
        if changed {
            return Ok(true);
        }
    }
    Ok(false)
}

fn set_nth_legacy_form_in_paragraph(
    paragraph: &mut CT_P,
    remaining: &mut usize,
    value: LegacyFormFieldValue,
) -> Result<bool> {
    for boundary in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter_mut()
            .filter(|(position, _, _, _)| *position == boundary)
        {
            if set_nth_legacy_form_in_control(control, remaining, value.clone())? {
                return Ok(true);
            }
        }
        if let Some(run) = paragraph.runs.get_mut(boundary) {
            for content in &mut run.content {
                if let RunContent::Field(field) = content
                    && set_nth_legacy_form_in_field(field, remaining, value.clone())?
                {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

fn set_nth_legacy_form_in_field(
    field: &mut Field,
    remaining: &mut usize,
    value: LegacyFormFieldValue,
) -> Result<bool> {
    field
        .set_nth_legacy_form_value_in_source_order(remaining, &value)
        .map_err(Into::into)
}

fn set_nth_legacy_form_in_table(
    table: &mut CT_Tbl,
    remaining: &mut usize,
    value: LegacyFormFieldValue,
) -> Result<bool> {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter_mut()
            .filter(|(position, _, _)| *position == boundary)
        {
            if set_nth_legacy_form_in_control(control, remaining, value.clone())? {
                return Ok(true);
            }
        }
        if let Some(row) = table.rows.get_mut(boundary)
            && set_nth_legacy_form_in_row(row, remaining, value.clone())?
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn set_nth_legacy_form_in_row(
    row: &mut CT_Row,
    remaining: &mut usize,
    value: LegacyFormFieldValue,
) -> Result<bool> {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter_mut()
            .filter(|(position, _, _)| *position == boundary)
        {
            if set_nth_legacy_form_in_control(control, remaining, value.clone())? {
                return Ok(true);
            }
        }
        if let Some(cell) = row.cells.get_mut(boundary) {
            for content in &mut cell.content {
                let changed = match content {
                    CellContent::Paragraph(paragraph) => {
                        set_nth_legacy_form_in_paragraph(paragraph, remaining, value.clone())?
                    }
                    CellContent::Table(table) => {
                        set_nth_legacy_form_in_table(table, remaining, value.clone())?
                    }
                    CellContent::ContentControl(control) => {
                        set_nth_legacy_form_in_control(control, remaining, value.clone())?
                    }
                };
                if changed {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

fn set_nth_legacy_form_in_control(
    control: &mut CT_Sdt,
    remaining: &mut usize,
    value: LegacyFormFieldValue,
) -> Result<bool> {
    let inline_positions = control
        .inline_legacy_form_fields_with_content_indices()?
        .into_iter()
        .map(|(content_index, _)| content_index)
        .collect::<Vec<_>>();
    let mut inline_index = 0usize;
    for content_index in 0..control.content.len() {
        while inline_positions
            .get(inline_index)
            .is_some_and(|position| *position == content_index)
        {
            if *remaining == 0 {
                return control
                    .set_inline_legacy_form_value(inline_index, value)
                    .map_err(Into::into);
            }
            *remaining -= 1;
            inline_index += 1;
        }
        let content = &mut control.content[content_index];
        let changed = match content {
            SdtContent::Paragraph(paragraph) => {
                set_nth_legacy_form_in_paragraph(paragraph, remaining, value.clone())?
            }
            SdtContent::Table(table) => {
                set_nth_legacy_form_in_table(table, remaining, value.clone())?
            }
            SdtContent::Row(row) => {
                let mut table = CT_Tbl::new();
                table.rows.push(row.clone());
                let changed = set_nth_legacy_form_in_table(&mut table, remaining, value.clone())?;
                if changed {
                    *row = table.rows.remove(0);
                }
                changed
            }
            SdtContent::Cell(cell) => {
                let mut row = CT_Row::new();
                row.cells.push(cell.clone());
                let mut table = CT_Tbl::new();
                table.rows.push(row);
                let changed = set_nth_legacy_form_in_table(&mut table, remaining, value.clone())?;
                if changed {
                    *cell = table.rows.remove(0).cells.remove(0);
                }
                changed
            }
            SdtContent::Run(_) => false,
            SdtContent::ContentControl(control) => {
                set_nth_legacy_form_in_control(control, remaining, value.clone())?
            }
            SdtContent::RawXml(_) => false,
        };
        if changed {
            return Ok(true);
        }
    }
    if inline_index != inline_positions.len() {
        return Err(Error::Other(
            "inline legacy form source boundary is invalid".to_owned(),
        ));
    }
    Ok(false)
}

/// A deterministic civil date and time supplied to field evaluation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FieldDateTime {
    pub year: i32,
    pub month: u8,
    pub day: u8,
    pub hour: u8,
    pub minute: u8,
    pub second: u8,
}

/// Explicit values that are not stored in the document package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FieldEvaluationContext {
    pub now: Option<FieldDateTime>,
    pub file_name: Option<String>,
    pub file_path: Option<String>,
    pub merge_fields: BTreeMap<String, String>,
    pub included_text: BTreeMap<String, String>,
    /// One-based source record number for mail-merge control fields.
    pub merge_record_number: Option<u32>,
    /// One-based output sequence number for mail-merge control fields.
    pub merge_sequence_number: Option<u32>,
}

/// An image value embedded by a rich mail merge.
#[derive(Debug, Clone, PartialEq)]
pub struct MailMergeImage {
    pub data: Vec<u8>,
    pub filename: String,
    pub width: Length,
    pub height: Length,
}

/// A scalar value accepted by a rich mail merge field.
#[derive(Debug, Clone, PartialEq)]
pub enum MailMergeValue {
    Text(String),
    Image(MailMergeImage),
    Fragment(Vec<u8>),
}

/// One rich mail merge record with nested lexical regions.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct MailMergeRecord {
    pub values: BTreeMap<String, MailMergeValue>,
    pub regions: BTreeMap<String, Vec<MailMergeRecord>>,
}

/// Root records and reusable named sources for a rich mail merge.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct MailMergeData {
    pub records: Vec<MailMergeRecord>,
    pub sources: BTreeMap<String, Vec<MailMergeRecord>>,
}

/// Stable location metadata supplied to a rich merge formatter.
pub struct MailMergeFormatContext<'a> {
    pub source_name: &'a str,
    pub region_path: &'a [String],
    pub field_name: &'a str,
    pub record_number: u32,
    pub output_sequence_number: u32,
}

/// Text and optional run properties exposed to a rich merge formatter.
pub struct MailMergeFormattedText {
    pub text: String,
    pub run_properties: Option<CT_RPr>,
}

/// The result of evaluating one field in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldEvaluation {
    pub field_index: usize,
    pub instruction: String,
    pub cached_result: String,
    pub outcome: FieldOutcome,
}

impl FieldEvaluation {
    /// Complete ordered evaluation facts shared by CLI and Python reports.
    #[must_use]
    pub fn to_json_value(&self) -> serde_json::Value {
        use serde_json::json;
        let outcome = match &self.outcome {
            FieldOutcome::Resolved(text) => json!({"kind":"resolved","text":text}),
            FieldOutcome::DeferredPagination => json!({"kind":"deferred_pagination"}),
            FieldOutcome::KeepStored { diagnostic } => {
                json!({"kind":"keep_stored","diagnostic":diagnostic})
            }
            FieldOutcome::TableOfContents(request) => {
                let entries = match &request.entries {
                    TocEntrySelection::None => json!({"kind":"none"}),
                    TocEntrySelection::All => json!({"kind":"all"}),
                    TocEntrySelection::Identifier(value) => {
                        json!({"kind":"identifier","value":value})
                    }
                };
                json!({"kind":"table_of_contents","request":{
                    "heading_levels":request.heading_levels,"custom_styles":request.custom_styles,
                    "entries":entries,"sequence_identifier":request.sequence_identifier,
                    "bookmark":request.bookmark,"hyperlink":request.hyperlink,
                    "use_outline_levels":request.use_outline_levels,
                    "omit_page_number_levels":request.omit_page_number_levels,
                    "page_number_separator":request.page_number_separator,"entry_page_separator":request.entry_page_separator}})
            }
            FieldOutcome::TableOfContentsEntry(request) => {
                json!({"kind":"table_of_contents_entry","request":{
                "entry":request.entry,"level":request.level,"table_identifier":request.table_identifier,
                "omit_page_number":request.omit_page_number}})
            }
            FieldOutcome::MailMergeControl(control) => {
                let control = match control {
                    MailMergeControl::NextRecord { record_number } => {
                        json!({"kind":"next_record","record_number":record_number})
                    }
                    MailMergeControl::NextRecordIf {
                        condition,
                        record_number,
                    } => {
                        json!({"kind":"next_record_if","condition":condition,"record_number":record_number})
                    }
                    MailMergeControl::SkipRecordIf {
                        condition,
                        record_number,
                    } => {
                        json!({"kind":"skip_record_if","condition":condition,"record_number":record_number})
                    }
                    MailMergeControl::RecordNumber(value) => {
                        json!({"kind":"record_number","value":value})
                    }
                    MailMergeControl::SequenceNumber(value) => {
                        json!({"kind":"sequence_number","value":value})
                    }
                };
                json!({"kind":"mail_merge_control","control":control})
            }
            FieldOutcome::Barcode(request) => json!({"kind":"barcode","request":{
                "value":request.value,"kind":format!("{:?}",request.kind),"height":request.height,
                "scale":request.scale,"error_correction":request.error_correction,
                "point_of_sale_style":request.point_of_sale_style.map(|value| format!("{value:?}")),
                "case_style":request.case_style.map(|value| format!("{value:?}")),
                "fix_check_digit":request.fix_check_digit,"rotation":request.rotation,
                "foreground_color":request.foreground_color,"background_color":request.background_color,
                "display_text":request.display_text,"add_start_stop":request.add_start_stop}}),
        };
        json!({"field_index":self.field_index,"instruction":self.instruction,
            "cached_result":self.cached_result,"outcome":outcome})
    }
}

/// Result of one atomic ordinary and layout-backed stored-cache update.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldsUpdateReport {
    pub ordinary_updated: usize,
    pub layout: LayoutBackedFieldUpdateReport,
    pub evaluations: Vec<FieldEvaluation>,
    pub diagnostics: Vec<String>,
    pub total_updated: usize,
    /// Supported typed and placed caches only. Header/footer PAGE and NUMPAGES stay stored.
    pub saved_cache_scope: &'static str,
}

/// A pure field-evaluation decision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FieldOutcome {
    Resolved(String),
    DeferredPagination,
    TableOfContents(TocField),
    TableOfContentsEntry(TcField),
    MailMergeControl(MailMergeControl),
    Barcode(BarcodeField),
    KeepStored { diagnostic: String },
}

/// A validated table-of-contents rebuild request.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TocField {
    pub heading_levels: Option<(u8, u8)>,
    pub custom_styles: Vec<(String, u8)>,
    pub entries: TocEntrySelection,
    pub sequence_identifier: Option<String>,
    pub bookmark: Option<String>,
    pub hyperlink: bool,
    pub use_outline_levels: bool,
    pub omit_page_number_levels: Option<(u8, u8)>,
    pub page_number_separator: Option<String>,
    pub entry_page_separator: Option<String>,
}

/// Outcome produced by one atomic table-of-contents rebuild.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TocRebuildReport {
    pub entry_count: usize,
    pub bookmark_count: usize,
    pub diagnostics: Vec<String>,
}

/// Results of refreshing caches whose values require deterministic layout.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct LayoutBackedFieldUpdateReport {
    pub page_fields: usize,
    pub num_pages_fields: usize,
    pub page_reference_fields: usize,
    pub section_fields: usize,
    pub section_pages_fields: usize,
    pub diagnostics: Vec<String>,
}

impl LayoutBackedFieldUpdateReport {
    /// Number of field caches written by the operation.
    #[must_use]
    pub fn updated_count(&self) -> usize {
        self.page_fields
            + self.num_pages_fields
            + self.page_reference_fields
            + self.section_fields
            + self.section_pages_fields
    }

    /// Number of layout diagnostics returned by the operation.
    #[must_use]
    pub fn diagnostic_count(&self) -> usize {
        self.diagnostics.len()
    }
}

impl TocRebuildReport {
    /// Number of retained fields that produced diagnostics.
    #[must_use]
    pub fn diagnostic_count(&self) -> usize {
        self.diagnostics.len()
    }
}

/// Which TC entries contribute to a table of contents.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TocEntrySelection {
    None,
    All,
    Identifier(String),
}

/// A validated table-of-contents entry request.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TcField {
    pub entry: String,
    pub level: u8,
    pub table_identifier: Option<String>,
    pub omit_page_number: bool,
}

/// A validated mail-merge control decision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MailMergeControl {
    NextRecord { record_number: u32 },
    NextRecordIf { condition: bool, record_number: u32 },
    SkipRecordIf { condition: bool, record_number: u32 },
    RecordNumber(u32),
    SequenceNumber(u32),
}

/// A validated generated-barcode request.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarcodeField {
    pub value: String,
    pub kind: BarcodeKind,
    pub height: Option<u32>,
    pub scale: Option<u16>,
    pub error_correction: Option<u8>,
    pub point_of_sale_style: Option<BarcodePointOfSaleStyle>,
    pub case_style: Option<BarcodeCaseStyle>,
    pub fix_check_digit: bool,
    pub rotation: Option<u8>,
    pub foreground_color: Option<u32>,
    pub background_color: Option<u32>,
    pub display_text: bool,
    pub add_start_stop: bool,
}

/// A barcode type accepted by Word's barcode field grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodeKind {
    Upca,
    Upce,
    Jan13,
    Jan8,
    Ean13,
    Ean8,
    Case,
    Itf14,
    Nw7,
    Code39,
    Code128,
    JpPost,
    Qr,
}

/// A point-of-sale style accepted by `\p`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodePointOfSaleStyle {
    Standard,
    SupplementalTwoDigit,
    SupplementalFiveDigit,
    Case,
}

/// An ITF14 case style accepted by `\c`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodeCaseStyle {
    Standard,
    Extended,
    Add,
}

#[cfg(test)]
thread_local! {
    static NAVIGATION_OWNER_VISITS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

impl Document {
    pub(crate) fn navigation_exclusions_for_input(
        &self,
        input: &rdocx_layout::LayoutInput,
    ) -> Result<Option<Arc<rdocx_layout::WordSequenceSnapshot>>> {
        let xml = input.document.to_xml()?;
        let mut spans = scan_dynamic_table_spans(&xml, DynamicOwnerPolicy::Bibliography)?;
        spans.retain(|span| {
            Field::new(&span.instruction, "")
                .effective_instruction()
                .name
                != "CITATION"
        });
        if spans.is_empty() {
            return Ok(None);
        }
        let mut snapshot = rdocx_layout::engine::evaluate_sequence_fields(input)?;
        let mut paths = Vec::new();
        for index in 1.. {
            let Some(node) = rdocx_layout::SourceNodeId::new(index) else {
                break;
            };
            let Some(path) = snapshot.source_node(node) else {
                break;
            };
            if path.story == rdocx_layout::WordStory::Document {
                paths.push(path.clone());
            }
        }
        let mut paragraphs = Vec::new();
        collect_body_paragraphs(&input.document.body, &mut paragraphs);
        if paragraphs.len() != paths.len() {
            return Err(Error::Other(
                "navigation generated-range physical source inventory disagrees".into(),
            ));
        }
        let mut exclusions = Vec::new();
        // Scanner spans are disjoint and physically ordered. Advance the paragraph
        // window once, then search end positions only inside that window. A single
        // paragraph may contain many owners, so scanning its whole window per run
        // would still amplify legal input quadratically.
        let mut span_start = 0;
        let mut span_end = 0;
        for (paragraph_index, (paragraph, path)) in paragraphs.iter().zip(paths).enumerate() {
            while span_start < spans.len() {
                #[cfg(test)]
                NAVIGATION_OWNER_VISITS.set(NAVIGATION_OWNER_VISITS.get() + 1);
                if spans[span_start].end_paragraph >= paragraph_index {
                    break;
                }
                span_start += 1;
            }
            span_end = span_end.max(span_start);
            while span_end < spans.len() {
                #[cfg(test)]
                NAVIGATION_OWNER_VISITS.set(NAVIGATION_OWNER_VISITS.get() + 1);
                if spans[span_end].begin_paragraph > paragraph_index {
                    break;
                }
                span_end += 1;
            }
            let relevant = &spans[span_start..span_end];
            if relevant.is_empty() {
                continue;
            }
            let excluded = accepted_toc_runs(paragraph)
                .iter()
                .enumerate()
                .filter_map(|(index, run)| {
                    let position = TocOwnedPosition {
                        paragraph: paragraph_index,
                        run: TocRunPosition {
                            run_boundary: run.run_boundary,
                            raw_order: run.raw_order,
                            nested_order: run.nested_order,
                        },
                    };
                    let candidate = relevant.partition_point(|span| {
                        #[cfg(test)]
                        NAVIGATION_OWNER_VISITS.set(NAVIGATION_OWNER_VISITS.get() + 1);
                        TocOwnedPosition {
                            paragraph: span.end_paragraph,
                            run: span.result_end_position,
                        } <= position
                    });
                    relevant.get(candidate).and_then(|span| {
                        toc_source_position_is_owned(std::slice::from_ref(span), position)
                            .then_some(index)
                    })
                })
                .collect();
            exclusions.push((path, excluded));
        }
        snapshot.set_navigation_exclusions(input, exclusions)?;
        Ok(Some(Arc::new(snapshot)))
    }
}

impl Document {
    /// Evaluate every typed field without mutating stored results.
    pub fn evaluate_fields(
        &self,
        context: &FieldEvaluationContext,
    ) -> Result<Vec<FieldEvaluation>> {
        let mut source = self.clone_for_staging();
        source.flush_dirty_related_story_models()?;
        prepare_physical_story_projection(&mut source.document.body, &mut [])?;
        let (evaluations, visible) = source.evaluate_fields_with_policy(context, false)?;
        Ok(evaluations
            .into_iter()
            .zip(visible)
            .filter_map(|(evaluation, visible)| visible.then_some(evaluation))
            .enumerate()
            .map(|(index, mut evaluation)| {
                evaluation.field_index = index;
                evaluation
            })
            .collect())
    }

    fn evaluate_fields_with_policy(
        &self,
        context: &FieldEvaluationContext,
        missing_merge_fields_as_empty: bool,
    ) -> Result<(Vec<FieldEvaluation>, Vec<bool>)> {
        let source = self;
        let mut evaluator = if missing_merge_fields_as_empty {
            Evaluator::for_mail_merge(source, context)
        } else {
            Evaluator::new(source, context)
        };
        let input = source.build_layout_input()?;
        let snapshot = rdocx_layout::engine::evaluate_sequence_fields(&input)
            .map_err(|error| Error::Other(format!("sequence source evaluation failed: {error}")))?;
        let mut nodes_by_story =
            HashMap::<rdocx_layout::WordStory, Vec<rdocx_layout::SourceNodeId>>::new();
        for index in 1u32.. {
            let Some(node) = rdocx_layout::SourceNodeId::new(index) else {
                break;
            };
            let Some(path) = snapshot.source_node(node) else {
                break;
            };
            nodes_by_story
                .entry(path.story.clone())
                .or_default()
                .push(node);
        }
        evaluator.sequence_snapshot = Some(Arc::new(snapshot));
        let mut main = Vec::new();
        collect_body_paragraphs(&source.document.body, &mut main);
        evaluator.sequence_nodes = nodes_by_story
            .get(&rdocx_layout::WordStory::Document)
            .map(|nodes| nodes.iter().copied().map(Some).collect())
            .unwrap_or_default();
        evaluator.refresh_main_bookmark_sequence_text(&main)?;
        evaluator.evaluate_story("main", &main)?;

        let general_parts = [true, false]
            .into_iter()
            .flat_map(|header| referenced_header_footer_parts(source, header, false))
            .map(|(part, _)| part)
            .collect::<HashSet<_>>();
        for (part_name, kind, xml) in field_story_parts(source, true) {
            let records = legacy_story_paragraphs(&xml, kind)?;
            evaluator.general_field_contexts = records
                .iter()
                .map(|record| {
                    record.general_context
                        && (!matches!(kind, PackageStoryKind::Header | PackageStoryKind::Footer)
                            || general_parts.contains(&part_name))
                })
                .collect();
            let paragraphs = records
                .iter()
                .map(|entry| &entry.paragraph)
                .collect::<Vec<_>>();
            let mut owners = nodes_by_story
                .iter()
                .filter(|(story, _)| {
                    input.story_part_names.get(*story) == Some(&part_name)
                        && matches!(
                            (story, kind),
                            (
                                rdocx_layout::WordStory::Header { .. },
                                PackageStoryKind::Header
                            ) | (
                                rdocx_layout::WordStory::Footer { .. },
                                PackageStoryKind::Footer
                            ) | (
                                rdocx_layout::WordStory::Footnote { .. },
                                PackageStoryKind::Footnotes
                            ) | (
                                rdocx_layout::WordStory::Endnote { .. },
                                PackageStoryKind::Endnotes
                            )
                        )
                })
                .collect::<Vec<_>>();
            owners.sort_by_key(|(_, nodes)| nodes.first().map(|id| id.get()));
            evaluator.sequence_nodes =
                if matches!(kind, PackageStoryKind::Header | PackageStoryKind::Footer) {
                    owners
                        .first()
                        .map(|(_, nodes)| nodes.iter().copied().map(Some).collect())
                        .unwrap_or_default()
                } else {
                    // Each note owner is physically distinct, unlike relationship aliases to furniture.
                    owners
                        .into_iter()
                        .flat_map(|(_, nodes)| nodes.iter().copied().map(Some))
                        .collect()
                };
            let story = match kind {
                PackageStoryKind::Header => format!("header:{part_name}"),
                PackageStoryKind::Footer => format!("footer:{part_name}"),
                PackageStoryKind::Footnotes => "footnotes".into(),
                PackageStoryKind::Endnotes => "endnotes".into(),
                PackageStoryKind::Comments => "comments".into(),
            };
            if matches!(kind, PackageStoryKind::Comments) {
                evaluator.sequence_nodes = vec![None; paragraphs.len()];
            }
            evaluator.evaluate_story(&story, &paragraphs)?;
        }

        for source in source.text_box_cache_paragraphs()? {
            let paragraph = parsed_physical_field_paragraph(&source.xml)?;
            let mut fields = Vec::new();
            let mut root_indices = HashSet::new();
            for run in paragraph.source_runs() {
                for content in &run.content {
                    if let RunContent::Field(field) = content {
                        root_indices.insert(fields.len());
                        let mut pending = vec![(field, false, false)];
                        while let Some((field, inherited_lock, cached)) = pending.pop() {
                            let locked = inherited_lock || field.locked() == Some(true);
                            fields.push((field, locked, cached));
                            pending.extend(
                                field
                                    .cached_fields_in_source_order()
                                    .into_iter()
                                    .rev()
                                    .map(|child| (child, locked, true)),
                            );
                            pending.extend(
                                field
                                    .nested_fields_in_source_order()
                                    .into_iter()
                                    .rev()
                                    .map(|child| (child, locked, cached)),
                            );
                        }
                    }
                }
            }
            if fields.is_empty() {
                continue;
            }
            let snapshot = evaluator
                .sequence_snapshot
                .as_ref()
                .expect("sequence inventory");
            let mut matching = Vec::new();
            for (story, nodes) in &nodes_by_story {
                if !matches!(story, rdocx_layout::WordStory::TextBox { part_name, .. } if part_name == source.location.story().part_name())
                    || snapshot.text_box_owner_index(story) != Some(source.owner_index)
                {
                    continue;
                }
                let Some(node) = nodes.get(source.paragraph_index).copied() else {
                    continue;
                };
                let mut enclosing_qualified = false;
                let qualified =
                    fields
                        .iter()
                        .enumerate()
                        .all(|(index, (field, locked, cached))| {
                            let root = root_indices.contains(&index);
                            if root {
                                enclosing_qualified = false;
                            }
                            let source = oxml_layout::FieldSource {
                                node,
                                index: index as u32,
                            };
                            if snapshot.source_field_context(source)
                                != Some((
                                    field.effective_instruction_text().as_str(),
                                    *locked,
                                    *cached,
                                ))
                            {
                                return false;
                            }
                            let model_raw = field
                                .source_replacement()
                                .ok()
                                .flatten()
                                .map(|(raw, _)| raw);
                            let source_raw = snapshot.source_field_xml(source);
                            match (model_raw, source_raw) {
                                (Some(model), Some(source)) if model == source => {
                                    if root {
                                        enclosing_qualified = true;
                                    }
                                    true
                                }
                                (None, None) if !root && enclosing_qualified => true,
                                _ => false,
                            }
                        });
                if qualified
                    && snapshot
                        .source_field_context(oxml_layout::FieldSource {
                            node,
                            index: fields.len() as u32,
                        })
                        .is_none()
                {
                    matching.push(node);
                }
            }
            evaluator.sequence_nodes = vec![(matching.len() == 1).then(|| matching[0])];
            evaluator.general_field_contexts = vec![false];
            evaluator.evaluate_story("textbox", &[&paragraph])?;
        }

        Ok((evaluator.results, evaluator.visible_results))
    }

    /// Evaluate and materialize every typed field cache in document order.
    pub fn update_fields(&mut self, context: &FieldEvaluationContext) -> Result<usize> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_dirty_related_story_models()?;
        let updated = candidate.update_fields_with_policy(context, false)?;
        if updated != 0 {
            self.commit_staged_mutation(candidate.prepare_and_reopen_staged()?);
        }
        Ok(updated)
    }

    /// Refresh ordinary and deterministic layout caches in one outer transaction.
    /// Omitted time remains unresolved. Ordinary retention diagnostics precede layout diagnostics.
    pub fn update_all_fields(
        &mut self,
        context: &FieldEvaluationContext,
    ) -> Result<FieldsUpdateReport> {
        let mut candidate = self.clone_for_staging();
        let evaluations = candidate.evaluate_fields(context)?;
        let ordinary = candidate.update_fields(context)?;
        let layout = candidate.update_layout_backed_fields()?;
        let mut diagnostics = evaluations
            .iter()
            .filter_map(|field| match &field.outcome {
                FieldOutcome::KeepStored { diagnostic } => Some(diagnostic.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        diagnostics.extend(layout.diagnostics.iter().cloned());
        let total_updated = ordinary
            .checked_add(layout.updated_count())
            .ok_or_else(|| Error::Other("field update count overflow".to_owned()))?;
        if total_updated != 0 {
            self.commit_staged_mutation(candidate.prepare_and_reopen_staged()?);
        }
        Ok(FieldsUpdateReport {
            ordinary_updated: ordinary,
            layout,
            evaluations,
            diagnostics,
            total_updated,
            saved_cache_scope: "supported-typed-and-placed-caches-with-dynamic-header-footer-page-caches-retained",
        })
    }

    /// Write layout-backed field caches and return the number changed.
    ///
    /// PAGE takes its placed story owner's displayed number. Word preserves
    /// dynamic PAGE and NUMPAGES caches in headers and footers on save, so
    /// those caches remain unchanged. PAGEREF takes its resolved
    /// bookmark. Written fields are marked clean. Every other field keeps its
    /// cache, as does a layout-backed field that layout does not place, that
    /// uses an unsupported switch, or whose section page number format layout
    /// does not reproduce. `update_fields` still defers these kinds. Use
    /// [`Self::update_layout_backed_fields`] for per-kind counts and layout
    /// diagnostics.
    pub fn update_page_fields(&mut self) -> Result<usize> {
        Ok(self.update_layout_backed_fields()?.updated_count())
    }

    /// Write page, section, total and resolved bookmark page caches from one layout.
    ///
    /// Each field uses its immutable placement and owning section. Header and
    /// footer PAGE and NUMPAGES caches retain Word's saved-cache behavior. Fields
    /// with an unsupported switch or unresolved target retain their caches.
    /// Successful writes are marked clean and committed atomically.
    pub fn update_layout_backed_fields(&mut self) -> Result<LayoutBackedFieldUpdateReport> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_dirty_related_story_models()?;
        prepare_physical_story_projection(&mut candidate.document.body, &mut [])?;
        let layout = candidate.layout_deterministic()?;
        let (updates, mut report) = candidate.page_field_updates(&layout)?;
        let textbox_updates = candidate.text_box_page_field_updates(&layout, &mut report)?;
        report.diagnostics.extend(
            layout
                .layout
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.clone())
                .collect::<Vec<_>>(),
        );
        let mut updated = candidate.apply_cached_field_updates(&updates, false)?;
        for patch in textbox_updates {
            candidate.patch_story_paragraph_field_sources(&patch.location, &patch.replacements)?;
            updated += patch.updated;
        }
        debug_assert_eq!(updated, report.updated_count());
        if updated != 0 {
            self.commit_staged_mutation(candidate);
        }
        Ok(report)
    }

    /// Stage one optional cache update per field, in update traversal order,
    /// for every PAGE, NUMPAGES, and PAGEREF field that `layout` places.
    fn page_field_updates(
        &self,
        layout: &rdocx_layout::WordLayoutResult,
    ) -> Result<(
        Vec<Option<CachedFieldUpdate>>,
        LayoutBackedFieldUpdateReport,
    )> {
        use rdocx_layout::{SourceNodeId, WordSourcePath, WordStory};

        let placed = placed_page_fields(layout);
        let mut updates = Vec::new();
        let mut report = LayoutBackedFieldUpdateReport::default();

        let mut main = Vec::new();
        collect_body_paragraphs(&self.document.body, &mut main);
        // Layout registers main-story paragraphs in this same flattened order.
        let main_paths = (1u32..)
            .map_while(|index| SourceNodeId::new(index).and_then(|id| layout.source_node(id)))
            .filter(|path| path.story == WordStory::Document)
            .map(|path| vec![path.clone()])
            .collect::<Vec<_>>();
        if main_paths.len() != main.len() {
            return Err(Error::Other(format!(
                "layout identified {} of {} main story paragraphs",
                main_paths.len(),
                main.len()
            )));
        }
        push_page_field_updates(
            &main,
            &main_paths,
            &placed,
            layout,
            &self.document,
            &mut updates,
            &mut report,
        );

        let input = self.build_layout_input()?;
        for (part_name, kind, xml) in field_story_parts(self, false) {
            let story = legacy_story_paragraphs(&xml, kind)?;
            let paragraphs = story
                .iter()
                .map(|entry| &entry.paragraph)
                .collect::<Vec<_>>();
            let mut physical_paths = BTreeMap::<String, Vec<WordSourcePath>>::new();
            for path in (1u32..)
                .map_while(|index| SourceNodeId::new(index).and_then(|id| layout.source_node(id)))
            {
                let matches_kind = matches!(
                    (&path.story, kind),
                    (WordStory::Header { .. }, PackageStoryKind::Header)
                        | (WordStory::Footer { .. }, PackageStoryKind::Footer)
                        | (WordStory::Footnote { .. }, PackageStoryKind::Footnotes)
                        | (WordStory::Endnote { .. }, PackageStoryKind::Endnotes)
                );
                if matches_kind && input.story_part_names.get(&path.story) == Some(&part_name) {
                    // Header relationships sharing one physical part describe
                    // alternative placements of the same paragraph inventory.
                    let key = match &path.story {
                        WordStory::Header { relationship_id }
                        | WordStory::Footer { relationship_id } => relationship_id.clone(),
                        _ => String::new(),
                    };
                    physical_paths.entry(key).or_default().push(path.clone());
                }
            }
            let paths = (0..paragraphs.len())
                .map(|index| {
                    physical_paths
                        .values()
                        .filter(|paths| paths.len() == paragraphs.len())
                        .filter_map(|paths| paths.get(index).cloned())
                        .collect()
                })
                .collect::<Vec<_>>();
            push_page_field_updates(
                &paragraphs,
                &paths,
                &placed,
                layout,
                &self.document,
                &mut updates,
                &mut report,
            );
        }

        Ok((updates, report))
    }

    fn text_box_page_field_updates(
        &self,
        layout: &rdocx_layout::WordLayoutResult,
        report: &mut LayoutBackedFieldUpdateReport,
    ) -> Result<Vec<TextBoxCachePatch>> {
        use rdocx_layout::{SourceNodeId, WordSourcePath, WordStory};
        let mut nodes = Vec::new();
        for index in 1u32.. {
            let Some(node) = SourceNodeId::new(index) else {
                break;
            };
            let Some(path) = layout.source_node(node) else {
                break;
            };
            if matches!(path.story, WordStory::TextBox { .. }) {
                nodes.push((node, path));
            }
        }
        let placed = placed_page_fields(layout);
        let mut result = Vec::new();
        for source in self.text_box_cache_paragraphs()? {
            let crate::document::TextBoxCacheParagraph {
                location,
                xml,
                owner_index,
                paragraph_index,
            } = source;
            let mut paragraph = parsed_physical_field_paragraph(&xml)?;
            let mut fields = Vec::new();
            for field in paragraph
                .source_runs()
                .into_iter()
                .flat_map(|run| &run.content)
                .filter_map(|content| match content {
                    RunContent::Field(field) => Some(field),
                    _ => None,
                })
            {
                collect_preorder_fields(field, false, &mut fields);
            }
            if fields.is_empty() {
                continue;
            }
            let raw = fields
                .iter()
                .map(|(field, _)| {
                    field
                        .source_replacement()
                        .ok()
                        .flatten()
                        .map(|(source, _)| source.to_vec())
                })
                .collect::<Vec<_>>();
            let matching = nodes.iter().filter(|(node, path)| {
                matches!(&path.story, WordStory::TextBox { part_name, .. } if part_name == location.story().part_name())
                    && layout.text_box_owner_index(&path.story) == Some(owner_index)
                    && layout.text_box_paragraph_index(*node) == Some(paragraph_index)
                    && raw.iter().enumerate().all(|(index, raw)| raw.as_deref().is_some_and(|raw| layout.source_field_xml(oxml_layout::FieldSource { node: *node, index: index as u32 }) == Some(raw)))
                    && layout.source_field_xml(oxml_layout::FieldSource { node: *node, index: raw.len() as u32 }).is_none()
            }).collect::<Vec<_>>();
            let paths = if matching.len() == 1 {
                vec![matching[0].1.clone()]
            } else {
                Vec::<WordSourcePath>::new()
            };
            let mut updates = Vec::new();
            push_page_field_updates(
                &[&paragraph],
                &[paths],
                &placed,
                layout,
                &self.document,
                &mut updates,
                report,
            );
            let count = updates.iter().flatten().count();
            if count == 0 {
                continue;
            }
            let original = paragraph.clone();
            let mut consumed = 0;
            apply_updates_to_paragraph(&mut paragraph, &updates, &mut consumed)?;
            if consumed != updates.len() {
                return Err(Error::Other(
                    "text-box field traversal lost source identity".to_owned(),
                ));
            }
            let replacements = paragraph_field_source_replacements(&original, &paragraph)?;
            if replacements.is_empty() {
                return Err(Error::Other(
                    "text-box cache edit has no producer source span".to_owned(),
                ));
            }
            result.push(TextBoxCachePatch {
                location,
                replacements,
                updated: count,
            });
        }
        Ok(result)
    }

    /// Rebuild every supported table of contents already present in the document.
    ///
    /// The operation stages bookmarks, cached entries, and deterministic page
    /// targets on an independent document. Any malformed or ambiguous source
    /// leaves the receiver unchanged.
    ///
    /// The cached entry paragraphs are replaced, so comment and bookmark
    /// markers placed on them are dropped with them. A comment anchored only
    /// there stays in the comments part without an anchor.
    pub fn rebuild_toc(&mut self) -> Result<TocRebuildReport> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        let document_xml = candidate
            .package
            .get_part(&candidate.doc_part_name)
            .ok_or(Error::NoDocumentPart)?
            .to_vec();
        let toc_spans = scan_dynamic_table_spans(&document_xml, DynamicOwnerPolicy::Toc)?;
        let mut diagnostics = collect_simple_toc_diagnostics(&document_xml)?;
        if toc_spans.is_empty() {
            return Ok(TocRebuildReport {
                diagnostics: diagnostics
                    .into_iter()
                    .map(|(_, diagnostic)| diagnostic)
                    .collect(),
                ..Default::default()
            });
        }

        let (toc_fields, mut dynamic_diagnostics) =
            parse_dynamic_toc_fields(&candidate, &document_xml, &toc_spans)?;
        diagnostics.append(&mut dynamic_diagnostics);
        diagnostics.sort_by_key(|(offset, _)| *offset);
        let mut diagnostics = diagnostics
            .into_iter()
            .map(|(_, diagnostic)| diagnostic)
            .collect::<Vec<_>>();
        if toc_fields.iter().all(Option::is_none) {
            return Ok(TocRebuildReport {
                diagnostics,
                ..Default::default()
            });
        }
        let bookmark_state = inspect_toc_bookmarks(&candidate.document.body, &document_xml)?;
        let numbering_layout = candidate.layout_deterministic()?;
        let sources = discover_toc_sources(
            &candidate,
            &toc_spans,
            &toc_fields,
            &bookmark_state,
            &numbering_layout,
        )?;
        diagnostics.extend(toc_duplicate_style_diagnostics(&candidate.styles));
        diagnostics.extend(toc_default_style_diagnostics(&candidate.styles));
        let (toc_entry_styles, retained_style_defects) =
            ensure_toc_entry_styles(&mut candidate, &sources)?;
        let entry_style_ids = toc_entry_style_ids(&candidate.styles, &toc_entry_styles);
        diagnostics.extend(retained_style_defects);
        candidate.flush_to_package()?;
        let rebuilt_toc_spans = toc_spans
            .iter()
            .zip(&toc_fields)
            .filter_map(|(span, field)| field.is_some().then_some(span.clone()))
            .collect::<Vec<_>>();
        let bookmark_repairs =
            toc_crossing_bookmark_repairs(&bookmark_state, &toc_spans, &toc_fields);
        let authored_bookmark_names = candidate.identifiers.authored_bookmark_names().clone();
        let mut allocator = TocBookmarkAllocator::new(bookmark_state, &authored_bookmark_names);
        let mut bookmark_by_paragraph = BTreeMap::<usize, TocBookmark>::new();
        let mut bookmark_name_remap = BTreeMap::<String, String>::new();
        for source in sources.iter().flatten() {
            if !source.needs_bookmark {
                continue;
            }
            if bookmark_by_paragraph.contains_key(&source.paragraph_index) {
                continue;
            }
            let is_partial_boundary_source = rebuilt_toc_spans.iter().any(|span| {
                source.paragraph_index == span.begin_paragraph
                    || source.paragraph_index == span.end_paragraph
            });
            if !is_partial_boundary_source
                && let Some(existing) = allocator.whole_paragraph_name(source.paragraph_index)
            {
                let authored = candidate.identifiers.is_authored_bookmark(existing.0);
                let name = if authored {
                    let name = allocator.allocate_name()?;
                    if name != existing.1 {
                        bookmark_name_remap.insert(existing.1, name.clone());
                    }
                    name
                } else {
                    existing.1
                };
                bookmark_by_paragraph.insert(
                    source.paragraph_index,
                    TocBookmark {
                        id: existing.0,
                        name,
                        insert: false,
                        authored,
                    },
                );
            } else {
                let allocated = allocator.allocate(&mut candidate.identifiers)?;
                bookmark_by_paragraph.insert(
                    source.paragraph_index,
                    TocBookmark {
                        id: allocated.0,
                        name: allocated.1,
                        insert: true,
                        authored: true,
                    },
                );
            }
        }
        if !bookmark_name_remap.is_empty() {
            let remap = BodyIdentityRemap {
                bookmark_names: bookmark_name_remap,
                ..Default::default()
            };
            let renamed_xml = patch_body_identity_attributes(&document_xml, &remap)?;
            CT_Document::from_xml(&renamed_xml)?;
            candidate
                .package
                .set_part(&candidate.doc_part_name, renamed_xml);
            let mut renamed = reopen_staged_document(candidate)?;
            restore_toc_bookmark_provenance(&mut renamed, &bookmark_by_paragraph);
            let report = renamed.rebuild_toc()?;
            self.commit_staged_mutation(renamed);
            return Ok(report);
        }
        let bookmark_count = bookmark_by_paragraph
            .values()
            .filter(|bookmark| bookmark.insert)
            .count();
        let provisional_xml = insert_toc_bookmarks_xml(
            &document_xml,
            &bookmark_by_paragraph,
            &bookmark_repairs,
            &toc_spans,
            &rebuilt_toc_spans,
        )?;
        CT_Document::from_xml(&provisional_xml)?;
        candidate
            .package
            .set_part(&candidate.doc_part_name, provisional_xml.clone());
        candidate = reopen_staged_document(candidate)?;
        let provisional_spans =
            scan_dynamic_table_spans(&provisional_xml, DynamicOwnerPolicy::Toc)?;
        if provisional_spans.len() != toc_fields.len() {
            return Err(Error::Other(
                "table of contents ownership changed while staging bookmarks".to_owned(),
            ));
        }
        let end_bookmark_starts = provisional_spans
            .iter()
            .enumerate()
            .map(|(toc_index, span)| {
                end_boundary_bookmark_starts(
                    toc_index,
                    span,
                    &provisional_spans,
                    &toc_fields,
                    &bookmark_by_paragraph,
                    &bookmark_repairs,
                )
            })
            .collect::<Vec<_>>();
        let provisional_paragraphs = toc_paragraph_insertions(&provisional_xml)?;
        let mut placeholders = Vec::new();
        let mut edits = Vec::new();
        let mut entry_count = 0usize;
        for (toc_index, ((span, toc), toc_sources)) in provisional_spans
            .iter()
            .zip(&toc_fields)
            .zip(&sources)
            .enumerate()
        {
            let Some(toc) = toc else {
                continue;
            };
            let generated = render_toc_entries(
                toc_index,
                toc,
                toc_sources,
                &toc_entry_styles,
                toc_section_text_width(&candidate.document.body, span.begin_paragraph),
                &bookmark_by_paragraph,
                &provisional_xml,
                &mut placeholders,
            )?;
            entry_count += toc_sources.len();
            if let Some(mut merged) = entry_boundary_toc_edits(
                &provisional_xml,
                span,
                &provisional_paragraphs,
                &generated,
                &entry_style_ids,
                &end_bookmark_starts[toc_index],
            )? {
                edits.append(&mut merged);
                continue;
            }
            if span.begin_paragraph == span.end_paragraph {
                return Err(Error::Other(
                    "table of contents result must span paragraph boundaries".to_owned(),
                ));
            }
            edits.push(FieldSourceEdit {
                start: span.result_start,
                end: span.result_end,
                replacement: dynamic_toc_replacement(
                    &provisional_xml,
                    span,
                    &generated,
                    &end_bookmark_starts[toc_index],
                )?,
            });
        }
        let mut staged_xml = provisional_xml;
        for edit in edits.into_iter().rev() {
            staged_xml.splice(edit.start..edit.end, edit.replacement);
        }
        CT_Document::from_xml(&staged_xml)?;
        candidate
            .package
            .set_part(&candidate.doc_part_name, staged_xml);
        let mut provisional = reopen_staged_document(candidate)?;

        let page_values = deterministic_toc_page_values(&provisional)?;
        let current_xml = provisional
            .package
            .get_part(&provisional.doc_part_name)
            .ok_or(Error::NoDocumentPart)?;
        let mut final_xml = current_xml.to_vec();
        let final_spans = scan_dynamic_table_spans(&final_xml, DynamicOwnerPolicy::Toc)?;
        if final_spans.len() != toc_fields.len() {
            return Err(Error::Other(
                "table of contents ownership changed during page substitution".to_owned(),
            ));
        }
        let mut page_edits = Vec::with_capacity(placeholders.len());
        for placeholder in &placeholders {
            let target = page_values.get(&placeholder.bookmark).ok_or_else(|| {
                Error::Other(format!(
                    "table of contents page target {} was not resolved",
                    placeholder.bookmark
                ))
            })?;
            let span = &final_spans[placeholder.toc_index];
            let owned = &final_xml[span.result_start..span.result_end];
            let matches = byte_match_offsets(owned, placeholder.token.as_bytes());
            if matches.len() != 1 {
                return Err(Error::Other(format!(
                    "table of contents page placeholder {} was not uniquely owned",
                    placeholder.token
                )));
            }
            let start = span.result_start + matches[0];
            page_edits.push(FieldSourceEdit {
                start,
                end: start + placeholder.token.len(),
                replacement: target.as_bytes().to_vec(),
            });
        }
        page_edits.sort_by_key(|edit| edit.start);
        for edit in page_edits.into_iter().rev() {
            final_xml.splice(edit.start..edit.end, edit.replacement);
        }
        let final_spans = scan_dynamic_table_spans(&final_xml, DynamicOwnerPolicy::Toc)?;
        final_xml =
            relocate_end_boundary_bookmark_starts(final_xml, &final_spans, &end_bookmark_starts)?;
        CT_Document::from_xml(&final_xml)?;
        provisional
            .package
            .set_part(&provisional.doc_part_name, final_xml);
        let mut completed = reopen_staged_document(provisional)?;
        restore_toc_bookmark_provenance(&mut completed, &bookmark_by_paragraph);
        completed.invalidate_layout();
        self.commit_staged_mutation(completed);

        Ok(TocRebuildReport {
            entry_count,
            bookmark_count,
            diagnostics,
        })
    }

    fn update_fields_with_policy(
        &mut self,
        context: &FieldEvaluationContext,
        missing_merge_fields_as_empty: bool,
    ) -> Result<usize> {
        prepare_physical_story_projection(&mut self.document.body, &mut [])?;
        let (evaluations, visible) =
            self.evaluate_fields_with_policy(context, missing_merge_fields_as_empty)?;
        let original = self.clone_for_staging();
        let mut fields = Vec::new();
        let mut body = Vec::new();
        collect_body_paragraphs(&original.document.body, &mut body);
        for paragraph in body {
            append_ref_copy_paragraph(paragraph, false, &mut fields);
        }
        for (_, kind, xml) in field_story_parts(&original, true) {
            for record in legacy_story_paragraphs(&xml, kind)? {
                append_ref_copy_paragraph(&record.paragraph, true, &mut fields);
            }
        }
        for source in original.text_box_cache_paragraphs()? {
            append_ref_copy_paragraph(
                &parsed_physical_field_paragraph(&source.xml)?,
                true,
                &mut fields,
            );
        }
        if fields.len() != evaluations.len() {
            return Err(Error::Other(
                "REF copy inventory disagrees with physical field evaluations".into(),
            ));
        }
        let mut updates = evaluations
            .iter()
            .map(|evaluation| match &evaluation.outcome {
                FieldOutcome::Resolved(value) => CachedFieldUpdate {
                    typed_runs: None,
                    comment_ranges: Vec::new(),
                    cached_result: value.clone(),
                    dirty: false,
                },
                FieldOutcome::DeferredPagination
                | FieldOutcome::TableOfContents(_)
                | FieldOutcome::TableOfContentsEntry(_)
                | FieldOutcome::MailMergeControl(_)
                | FieldOutcome::Barcode(_)
                | FieldOutcome::KeepStored { .. } => CachedFieldUpdate {
                    typed_runs: None,
                    comment_ranges: Vec::new(),
                    cached_result: evaluation.cached_result.clone(),
                    dirty: true,
                },
            })
            .map(Some)
            .collect::<Vec<_>>();
        for ((source, update), visible) in fields.iter().zip(&mut updates).zip(visible) {
            if source.locked || !visible {
                *update = None;
            }
        }
        let source_projection = if fields.iter().any(|source| {
            let instruction = source.field.effective_instruction();
            instruction.name == "REF"
                && has_switch(&instruction, "f")
                && !source.locked
                && !source.generated
        }) {
            let sequence_updates = fields
                .iter()
                .zip(&evaluations)
                .map(|(source, evaluation)| {
                    if source.locked
                        || source.generated
                        || source.field.effective_instruction().name != "SEQ"
                    {
                        return None;
                    }
                    if let FieldOutcome::Resolved(value) = &evaluation.outcome {
                        Some(CachedFieldUpdate {
                            cached_result: value.clone(),
                            dirty: false,
                            typed_runs: None,
                            comment_ranges: Vec::new(),
                        })
                    } else {
                        None
                    }
                })
                .collect::<Vec<_>>();
            if sequence_updates.iter().any(Option::is_some) {
                let mut projection = original.clone_for_staging();
                projection.apply_cached_physical_field_updates(&sequence_updates)?;
                projection.document = CT_Document::from_xml(&projection.document.to_xml()?)?;
                Some(projection)
            } else {
                None
            }
        } else {
            None
        };
        let ref_source = source_projection.as_ref().unwrap_or(&original);
        let mut note_copies = Vec::new();
        let mut comment_copies = Vec::new();
        let mut occupied_comments = original
            .comments
            .as_ref()
            .map(|comments| {
                comments
                    .comments
                    .iter()
                    .map(|comment| comment.id)
                    .collect::<HashSet<_>>()
            })
            .unwrap_or_default();
        let mut occupied_comment_identities = None;
        let mut occupied_notes = HashMap::new();
        let note_reference_counts = if fields.iter().any(|source| {
            let instruction = source.field.effective_instruction();
            !source.suppress_copied_references
                && !source.locked
                && !source.generated
                && instruction.name == "REF"
                && has_switch(&instruction, "f")
        }) {
            ref_note_reference_counts(&original.package)?
        } else {
            HashMap::new()
        };
        for (source, update) in fields.iter().zip(&mut updates) {
            let instruction = source.field.effective_instruction();
            if instruction.name != "REF" || !has_switch(&instruction, "f") {
                continue;
            }
            if source.locked || source.generated {
                *update = None;
                continue;
            }
            if update.as_ref().is_some_and(|update| update.dirty) {
                *update = None;
                continue;
            }
            let Some(update) = update else {
                continue;
            };
            let target = text_argument(&instruction, 0)
                .ok_or_else(|| Error::Other("REF f target is absent".into()))?;
            let Some(cache) = ref_bookmark_runs(ref_source, target)? else {
                continue;
            };
            let comment_replacements = if source.suppress_copied_references {
                Vec::new()
            } else {
                ref_comment_replacement_ids(&original, &source.field, &cache)?
            };
            let mut comment_cursor = 0;
            let mut runs = cache.runs;
            let mut comment_ranges = if source.suppress_copied_references {
                Vec::new()
            } else {
                cache.comment_ranges
            };
            let mut old_references = Vec::new();
            collect_ref_cached_note_references(&source.field, &mut old_references);
            for run in &mut runs {
                if source.suppress_copied_references {
                    run.content.retain(|content| {
                        !matches!(
                            content,
                            RunContent::FootnoteRef { .. }
                                | RunContent::EndnoteRef { .. }
                                | RunContent::CommentReference { .. }
                        )
                    });
                } else {
                    for content in &mut run.content {
                        if let RunContent::CommentReference { id, .. } = content {
                            if occupied_comment_identities.is_none() {
                                occupied_comment_identities =
                                    Some(ref_comment_occupied_identity_values(&original.package)?);
                            }
                            let (new_id, copy) = prepare_ref_comment_copy(
                                self,
                                ref_source,
                                *id,
                                comment_replacements[comment_cursor],
                                &mut occupied_comments,
                                occupied_comment_identities.as_mut().unwrap(),
                            )?;
                            comment_cursor += 1;
                            for marker in &mut comment_ranges {
                                let marker_id = match marker {
                                    CommentRangeMarker::Start { id, .. }
                                    | CommentRangeMarker::End { id, .. } => id,
                                };
                                if *marker_id == *id {
                                    *marker_id = new_id;
                                }
                            }
                            *id = new_id;
                            comment_copies.push(copy);
                            continue;
                        }
                        let kind = match content {
                            RunContent::FootnoteRef { .. } => StoryKind::Footnote,
                            RunContent::EndnoteRef { .. } => StoryKind::Endnote,
                            _ => continue,
                        };
                        let id = match content {
                            RunContent::FootnoteRef { id, .. }
                            | RunContent::EndnoteRef { id, .. } => id,
                            _ => unreachable!("note kind"),
                        };
                        let replacement = old_references
                            .iter()
                            .position(|(family, _)| *family == kind)
                            .map(|index| old_references.remove(index).1)
                            .filter(|old| {
                                *old != *id && note_reference_counts.get(&(kind, *old)) == Some(&1)
                            });
                        let (new_id, copy) = prepare_ref_note_copy(
                            self,
                            ref_source,
                            kind,
                            *id,
                            replacement,
                            &mut occupied_notes,
                        )?;
                        *id = new_id;
                        note_copies.push(copy);
                    }
                }
            }
            let mut candidate = source.field.clone();
            candidate.set_cached_runs_with_comment_ranges(runs.clone(), comment_ranges.clone())?;
            update.cached_result = candidate.cached_result;
            update.typed_runs = Some(runs);
            update.comment_ranges = comment_ranges;
        }
        let updated = self.apply_cached_physical_field_updates(&updates)?;
        for copy in note_copies {
            self.publish_fragment_note_staged(
                copy.kind,
                &copy.destination_part,
                copy.xml,
                copy.replace_id,
            )?;
        }
        for copy in comment_copies {
            publish_ref_comment_copy(self, copy)?;
        }
        Ok(updated)
    }

    fn apply_cached_physical_field_updates(
        &mut self,
        updates: &[Option<CachedFieldUpdate>],
    ) -> Result<usize> {
        let boxes = self.text_box_cache_paragraphs()?;
        let mut paragraphs = Vec::new();
        let mut box_count = 0;
        for source in boxes {
            let paragraph = parsed_physical_field_paragraph(&source.xml)?;
            let mut fields = Vec::new();
            for run in accepted_toc_runs(&paragraph) {
                for content in &run.run.content {
                    if let RunContent::Field(field) = content {
                        collect_preorder_fields(field, false, &mut fields);
                    }
                }
            }
            if fields.is_empty() {
                continue;
            }
            let count = fields.len();
            box_count += count;
            paragraphs.push((source.location, paragraph, count));
        }
        let split = updates.len().checked_sub(box_count).ok_or_else(|| {
            Error::Other("text-box update inventory exceeds field evaluations".into())
        })?;
        let mut cursor = split;
        let mut patches = Vec::new();
        for (location, mut paragraph, count) in paragraphs {
            let end = cursor + count;
            if !updates[cursor..end].iter().any(Option::is_some) {
                cursor = end;
                continue;
            }
            let original = paragraph.clone();
            let mut consumed = 0;
            apply_updates_to_paragraph(&mut paragraph, &updates[cursor..end], &mut consumed)?;
            if consumed != count {
                return Err(Error::Other(
                    "text-box update field inventory is inconsistent".into(),
                ));
            }
            let replacements = paragraph_field_source_replacements(&original, &paragraph)?;
            if replacements.is_empty() {
                return Err(Error::Other(
                    "text-box update has no preserved field source".into(),
                ));
            }
            patches.push(TextBoxCachePatch {
                location,
                replacements,
                updated: updates[cursor..end].iter().flatten().count(),
            });
            cursor = end;
        }
        if cursor != updates.len() {
            return Err(Error::Other(
                "text-box update left unconsumed evaluations".into(),
            ));
        }
        let mut updated = self.apply_cached_field_updates(&updates[..split], true)?;
        for patch in patches {
            self.patch_story_paragraph_field_sources(&patch.location, &patch.replacements)?;
            updated += patch.updated;
        }
        Ok(updated)
    }

    fn apply_cached_field_updates(
        &mut self,
        updates: &[Option<CachedFieldUpdate>],
        include_comments: bool,
    ) -> Result<usize> {
        if updates
            .iter()
            .flatten()
            .any(|update| !update.cached_result.chars().all(valid_xml_character))
        {
            return Err(Error::Other(
                "field result contains a character forbidden by XML 1.0".to_owned(),
            ));
        }
        let updated = updates.iter().flatten().count();
        if updated == 0 {
            return Ok(0);
        }
        let package_before = self.package.clone();
        let mut document = self.document.clone();
        let mut update_index = 0;
        apply_updates_to_body(&mut document.body, updates, &mut update_index)?;
        let mut staged = Vec::new();
        for (part_name, kind, xml) in field_story_parts(self, include_comments) {
            let mut paragraphs = legacy_story_paragraphs(&xml, kind)?;
            let first = update_index;
            for entry in &mut paragraphs {
                apply_updates_to_paragraph(&mut entry.paragraph, updates, &mut update_index)?;
            }
            if updates[first..update_index].iter().any(Option::is_some) {
                let xml = patch_legacy_story_field_sources(&xml, &paragraphs, true)?;
                match kind {
                    PackageStoryKind::Header | PackageStoryKind::Footer => {
                        CT_HdrFtr::from_xml(&xml)?;
                    }
                    PackageStoryKind::Footnotes | PackageStoryKind::Endnotes => {
                        CT_Footnotes::from_xml(&xml)?;
                    }
                    PackageStoryKind::Comments => {
                        CT_Comments::from_xml(&xml)?;
                    }
                }
                staged.push((part_name, kind, xml));
            }
        }
        if update_index != updates.len() {
            return Err(Error::Other(format!(
                "page field traversal consumed {update_index} of {} staged evaluations",
                updates.len()
            )));
        }
        CT_Document::from_xml(&document.to_xml()?)?;
        self.document = document;
        for (part_name, kind, xml) in staged {
            if matches!(kind, PackageStoryKind::Footnotes) {
                self.footnotes = CT_Footnotes::from_xml(&xml)?;
                self.footnotes_dirty = false;
            }
            if matches!(kind, PackageStoryKind::Comments) {
                self.comments = Some(CT_Comments::from_xml(&xml)?);
                self.comments_dirty = false;
            }
            self.package.set_part(&part_name, xml);
        }
        self.package_signatures_invalidated |= self
            .retained_package_signature_would_be_invalidated()?
            || crate::embedded::synchronized_package_mutation_invalidates_signature(
                &package_before,
                &self.package,
            );
        self.invalidate_layout();
        Ok(updated)
    }

    /// Materialize one independent document for each flat mail-merge record.
    pub fn mail_merge(&self, records: &[BTreeMap<String, String>]) -> Result<Vec<Document>> {
        if records.is_empty() {
            return Err(Error::Other(
                "mail merge requires at least one record".to_owned(),
            ));
        }

        let mut outputs = Vec::with_capacity(records.len());
        for (record_index, record) in records.iter().enumerate() {
            let record_number = u32::try_from(record_index)
                .ok()
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| Error::Other("mail merge record count exceeds u32".to_owned()))?;
            let context = FieldEvaluationContext {
                merge_fields: record.clone(),
                merge_record_number: Some(record_number),
                merge_sequence_number: Some(record_number),
                ..Default::default()
            };
            let mut candidate = self.clone_for_staging();
            candidate.update_fields_with_policy(&context, true)?;
            outputs.push(candidate.prepare_and_reopen_staged()?);
        }
        Ok(outputs)
    }

    /// Materialize one document whose record bodies form next-page sections.
    pub fn mail_merge_sections(&self, records: &[BTreeMap<String, String>]) -> Result<Document> {
        reject_varying_non_body_merge_fields(self, records)?;
        let mut candidates = self.mail_merge(records)?;

        let mut identity_state = BodyIdentityState::from_documents(&candidates)?;
        let mut identifiers = candidates
            .first()
            .ok_or_else(|| Error::Other("mail merge requires at least one record".to_owned()))?
            .identifiers
            .clone();
        for candidate in candidates.iter_mut().skip(1) {
            remap_body_identities(candidate, &mut identifiers, &mut identity_state)?;
        }
        combine_mail_merge_sections(candidates)
    }

    /// Materialize independent documents from nested, typed mail merge data.
    #[allow(clippy::type_complexity)]
    pub fn mail_merge_rich(
        &self,
        data: &MailMergeData,
        mut formatter: Option<
            &mut dyn FnMut(&MailMergeFormatContext<'_>, &mut MailMergeFormattedText) -> Result<()>,
        >,
    ) -> Result<Vec<Document>> {
        if data.records.is_empty() {
            return Err(Error::Other(
                "rich mail merge requires at least one record".to_owned(),
            ));
        }

        let mut outputs = Vec::with_capacity(data.records.len());
        for (record_index, record) in data.records.iter().enumerate() {
            let record_number = one_based_record_number(record_index)?;
            let mut candidate = self.clone_for_staging();
            let mut identity_state =
                BodyIdentityState::from_documents(std::slice::from_ref(&candidate))?;
            let template = std::mem::take(&mut candidate.document.body.content);
            let mut region_path = Vec::new();
            let scopes = vec![record];
            let mut output_sequence_number = 0;
            candidate.document.body.content = expand_rich_body(
                &mut candidate,
                &template,
                data,
                &scopes,
                "records",
                &mut region_path,
                record_number,
                &mut output_sequence_number,
                &mut identity_state,
                &mut formatter,
            )?;
            candidate.invalidate_layout();
            outputs.push(candidate.prepare_and_reopen_staged()?);
        }
        Ok(outputs)
    }

    /// Materialize rich merge records as next-page sections in one document.
    #[allow(clippy::type_complexity)]
    pub fn mail_merge_sections_rich(
        &self,
        data: &MailMergeData,
        formatter: Option<
            &mut dyn FnMut(&MailMergeFormatContext<'_>, &mut MailMergeFormattedText) -> Result<()>,
        >,
    ) -> Result<Document> {
        let mut candidates = self.mail_merge_rich(data, formatter)?;
        let mut identity_state = BodyIdentityState::from_documents(&candidates)?;
        let mut identifiers = candidates
            .first()
            .ok_or_else(|| Error::Other("rich mail merge requires at least one record".to_owned()))?
            .identifiers
            .clone();
        for candidate in candidates.iter_mut().skip(1) {
            remap_body_identities(candidate, &mut identifiers, &mut identity_state)?;
        }
        combine_mail_merge_sections(candidates)
    }

    /// Update typed field caches, then save the package to a file path.
    pub fn save_with_field_updates<P: AsRef<Path>>(
        &mut self,
        path: P,
        context: &FieldEvaluationContext,
    ) -> Result<()> {
        self.update_fields(context)?;
        self.save(path)
    }

    /// Update typed field caches, then save the package to bytes.
    pub fn to_bytes_with_field_updates(
        &mut self,
        context: &FieldEvaluationContext,
    ) -> Result<Vec<u8>> {
        self.update_fields(context)?;
        self.to_bytes()
    }
}

fn restore_toc_bookmark_provenance(
    document: &mut Document,
    bookmarks: &BTreeMap<usize, TocBookmark>,
) {
    for bookmark in bookmarks.values().filter(|bookmark| bookmark.authored) {
        document
            .identifiers
            .restore_authored_bookmark(bookmark.id, &bookmark.name);
    }
}

fn one_based_record_number(index: usize) -> Result<u32> {
    u32::try_from(index)
        .ok()
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| Error::Other("rich mail merge record count exceeds u32".to_owned()))
}

fn combine_mail_merge_sections(mut candidates: Vec<Document>) -> Result<Document> {
    let bodies = candidates
        .iter()
        .map(|candidate| candidate.document.body.clone())
        .collect::<Vec<_>>();
    let mut combined = candidates.remove(0);
    combined.document.body.content.clear();
    combined.document.body.sect_pr = None;
    let final_index = bodies.len() - 1;
    for (index, mut body) in bodies.into_iter().enumerate() {
        combined.document.body.content.append(&mut body.content);
        if index == final_index {
            combined.document.body.sect_pr = body.sect_pr;
        } else {
            let mut section = body.sect_pr.unwrap_or_else(empty_section_properties);
            section.section_type = Some(ST_SectionType::NextPage);
            let mut paragraph = CT_P::new();
            paragraph.properties = Some(CT_PPr {
                sect_pr: Some(section),
                ..Default::default()
            });
            combined
                .document
                .body
                .content
                .push(BodyContent::Paragraph(Box::new(paragraph)));
        }
    }
    combined.invalidate_layout();
    combined.prepare_and_reopen_staged()
}

type RichFormatter<'a> = Option<
    &'a mut dyn FnMut(&MailMergeFormatContext<'_>, &mut MailMergeFormattedText) -> Result<()>,
>;

#[derive(Debug, Clone, PartialEq, Eq)]
enum RichRegionMarker {
    Start(String),
    End(String),
}

struct RichExpandedRow {
    source_index: usize,
    row: CT_Row,
}

const RICH_MERGE_REGION_DEPTH_LIMIT: usize = 32;

fn merge_field_name(field: &Field) -> Option<String> {
    let instruction = field.effective_instruction();
    if instruction.name != "MERGEFIELD" {
        return None;
    }
    instruction
        .arguments
        .first()
        .and_then(argument_text)
        .map(str::to_owned)
}

fn rich_region_marker(paragraph: &CT_P) -> Option<RichRegionMarker> {
    if paragraph_has_raw_content(paragraph)
        || paragraph.properties.is_some()
        || !paragraph.content_controls.is_empty()
        || !paragraph.revisions.is_empty()
        || !paragraph.hyperlinks.is_empty()
        || !paragraph.comment_ranges.is_empty()
        || !paragraph.bookmark_markers.is_empty()
        || !paragraph.equations.is_empty()
        || paragraph.runs.len() != 1
    {
        return None;
    }
    let run = &paragraph.runs[0];
    if run
        .extra_xml
        .iter()
        .any(|raw| !raw.iter().all(u8::is_ascii_whitespace))
        || !run.alt_drawings.is_empty()
    {
        return None;
    }
    let [RunContent::Field(field)] = run.content.as_slice() else {
        return None;
    };
    let instruction = field.effective_instruction();
    if instruction.name != "MERGEFIELD"
        || instruction.arguments.len() != 1
        || !instruction.switches.is_empty()
    {
        return None;
    }
    let name = merge_field_name(field)?;
    name.strip_prefix("TableStart:")
        .map(|name| RichRegionMarker::Start(name.to_owned()))
        .or_else(|| {
            name.strip_prefix("TableEnd:")
                .map(|name| RichRegionMarker::End(name.to_owned()))
        })
}

/// Whether a paragraph holds raw XML other than whitespace. The attributes
/// of its start tag, such as `w:rsidR` or `w14:paraId`, are not content.
fn paragraph_has_raw_content(paragraph: &CT_P) -> bool {
    paragraph.extra_xml.iter().any(|(position, raw)| {
        !CT_P::raw_is_root_attributes(*position, raw) && !raw.iter().all(u8::is_ascii_whitespace)
    })
}

fn find_body_region_end(items: &[BodyContent], start: usize, name: &str) -> Result<usize> {
    let mut stack = vec![name.to_owned()];
    for (index, item) in items.iter().enumerate().skip(start) {
        let BodyContent::Paragraph(paragraph) = item else {
            continue;
        };
        match rich_region_marker(paragraph) {
            Some(RichRegionMarker::Start(nested)) => stack.push(nested),
            Some(RichRegionMarker::End(end)) => {
                if stack.last() != Some(&end) {
                    return Err(Error::Other(format!(
                        "crossed rich mail merge region marker TableEnd:{end}"
                    )));
                }
                stack.pop();
                if stack.is_empty() {
                    return Ok(index);
                }
            }
            None => {}
        }
    }
    Err(Error::Other(format!(
        "rich mail merge region TableStart:{name} has no matching end"
    )))
}

#[allow(clippy::too_many_arguments)]
fn expand_rich_body(
    document: &mut Document,
    items: &[BodyContent],
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &mut Vec<String>,
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<Vec<BodyContent>> {
    let mut output = Vec::new();
    let mut index = 0;
    while index < items.len() {
        if let BodyContent::Paragraph(paragraph) = &items[index]
            && let Some(marker) = rich_region_marker(paragraph)
        {
            match marker {
                RichRegionMarker::Start(name) => {
                    if region_path.len() >= RICH_MERGE_REGION_DEPTH_LIMIT {
                        return Err(Error::Other(format!(
                            "rich mail merge region depth exceeds {RICH_MERGE_REGION_DEPTH_LIMIT}"
                        )));
                    }
                    let end = find_body_region_end(items, index + 1, &name)?;
                    let records = resolve_region_records(scopes, data, &name)?;
                    region_path.push(name.clone());
                    for (child_index, child) in records.iter().enumerate() {
                        let mut child_scopes = scopes.to_vec();
                        child_scopes.push(child);
                        let mut child_content = expand_rich_body(
                            document,
                            &items[index + 1..end],
                            data,
                            &child_scopes,
                            &name,
                            region_path,
                            one_based_record_number(child_index)?,
                            output_sequence_number,
                            identity_state,
                            formatter,
                        )?;
                        remap_rich_body_content(document, &mut child_content, identity_state)?;
                        output.extend(child_content);
                    }
                    region_path.pop();
                    index = end + 1;
                    continue;
                }
                RichRegionMarker::End(name) => {
                    return Err(Error::Other(format!(
                        "unexpected rich mail merge region marker TableEnd:{name}"
                    )));
                }
            }
        }
        if let BodyContent::Paragraph(paragraph) = &items[index]
            && let Some(fragment) = whole_paragraph_fragment(paragraph, scopes)?
        {
            output.extend(import_rich_fragment(document, fragment, identity_state)?);
            index += 1;
            continue;
        }
        let mut item = items[index].clone();
        replace_rich_body_item(
            document,
            &mut item,
            data,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        )?;
        output.push(item);
        index += 1;
    }
    Ok(output)
}

fn whole_paragraph_fragment<'a>(
    paragraph: &CT_P,
    scopes: &[&'a MailMergeRecord],
) -> Result<Option<&'a [u8]>> {
    if paragraph_has_raw_content(paragraph)
        || !paragraph.content_controls.is_empty()
        || !paragraph.revisions.is_empty()
        || !paragraph.hyperlinks.is_empty()
        || !paragraph.comment_ranges.is_empty()
        || !paragraph.bookmark_markers.is_empty()
        || !paragraph.equations.is_empty()
        || paragraph.runs.len() != 1
    {
        return Ok(None);
    }
    let [RunContent::Field(field)] = paragraph.runs[0].content.as_slice() else {
        return Ok(None);
    };
    let Some(name) = merge_field_name(field) else {
        return Ok(None);
    };
    match resolve_rich_value(scopes, &name) {
        Some(MailMergeValue::Fragment(bytes)) => Ok(Some(bytes)),
        _ => Ok(None),
    }
}

fn import_rich_fragment(
    document: &mut Document,
    bytes: &[u8],
    identity_state: &mut BodyIdentityState,
) -> Result<Vec<BodyContent>> {
    import_fragment_content_with_state(
        document,
        bytes,
        false,
        false,
        FragmentConflictPolicy::rename_all().with_style_reuse(true),
        &document.doc_part_name.clone(),
        identity_state,
        false,
        &mut Vec::new(),
    )
    .map(|imported| imported.typed)
}

pub(crate) fn import_document_fragment_content(
    document: &mut Document,
    bytes: &[u8],
    include_final_section_properties: bool,
    policy: FragmentConflictPolicy,
    destination_part: &str,
    source_owned_numbering: bool,
    note_copy_sources: &mut Vec<(StoryKind, String, i32, String, i32)>,
) -> Result<Vec<u8>> {
    let mut identity_state = BodyIdentityState::from_documents(std::slice::from_ref(document))?;
    import_fragment_content_with_state(
        document,
        bytes,
        include_final_section_properties,
        true,
        policy,
        destination_part,
        &mut identity_state,
        source_owned_numbering,
        note_copy_sources,
    )
    .map(|imported| imported.xml)
}

struct ImportedFragmentContent {
    typed: Vec<BodyContent>,
    xml: Vec<u8>,
}

struct FragmentNoteCopy {
    kind: StoryKind,
    destination_part: String,
    xml: Vec<u8>,
    replace_id: Option<i32>,
}

fn import_fragment_content_with_state(
    document: &mut Document,
    bytes: &[u8],
    include_final_section_properties: bool,
    allow_external_relationships: bool,
    policy: FragmentConflictPolicy,
    destination_part: &str,
    identity_state: &mut BodyIdentityState,
    source_owned_numbering: bool,
    note_copy_sources: &mut Vec<(StoryKind, String, i32, String, i32)>,
) -> Result<ImportedFragmentContent> {
    let mut fragment = Document::from_bytes(bytes)
        .map_err(|error| Error::Other(format!("invalid document fragment: {error}")))?;
    let mut fragment_xml = fragment
        .package
        .get_part(&fragment.doc_part_name)
        .ok_or_else(|| Error::Other("document fragment main part is missing".to_owned()))?
        .to_vec();
    fragment.prepare_staged_package()?;
    if document.glossary_part_name.as_deref() == Some(destination_part) {
        // The prepared physical body selects current note ids, excluding root payload.
        fragment_xml =
            fragment.omit_glossary_selected_body_comments(include_final_section_properties)?;
    }
    let final_section_properties = fragment.document.body.sect_pr.take();
    if include_final_section_properties && let Some(section_properties) = final_section_properties {
        let mut paragraph = CT_P::new();
        paragraph.properties = Some(CT_PPr {
            sect_pr: Some(section_properties),
            ..Default::default()
        });
        fragment
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(paragraph)));
    }
    let dependency_source = if document.glossary_part_name.as_deref() == Some(destination_part) {
        wrap_fragment_companion(&crate::document::package_authoritative_body_fragment(
            &fragment.document.to_xml()?,
            false,
            &BTreeMap::new(),
        )?)
    } else {
        fragment.document.to_xml()?
    };
    prune_document_fragment_dependencies(&mut fragment, &dependency_source)?;
    let mut fragment_identity_values = body_identity_values(&fragment_xml)?;
    for companion in fragment_dependency_companions(&fragment, &dependency_source)? {
        for id in body_identity_values(&wrap_fragment_companion(&companion))?.comment_ids {
            if !fragment_identity_values.comment_ids.contains(&id) {
                fragment_identity_values.comment_ids.push(id);
            }
        }
    }
    validate_fragment_identity_ownership(&fragment_identity_values)?;
    if source_owned_numbering {
        document.source_owned_numbering_remap(&fragment)?;
    }
    remap_fragment_style_collisions(
        document,
        &mut fragment,
        &mut fragment_xml,
        policy.reuse_styles(),
        policy.reuse_numbering(),
        source_owned_numbering,
    )?;
    let dependency_xml =
        wrap_fragment_companion(&crate::document::package_authoritative_body_fragment(
            &fragment_xml,
            include_final_section_properties,
            &BTreeMap::new(),
        )?);
    let used_relationships = relationship_ids_in_xml(&dependency_xml)?;
    let source_rels = fragment
        .package
        .get_part_rels(&fragment.doc_part_name)
        .cloned()
        .unwrap_or_default();
    let mut relationship_map = BTreeMap::new();
    let mut comment_relationship_map = BTreeMap::new();
    let mut part_map = BTreeMap::new();
    let mut imports = Vec::new();
    let mut external_imports = Vec::new();
    for relationship_id in used_relationships {
        let relationship = source_rels.get_by_id(&relationship_id).ok_or_else(|| {
            Error::Other(format!(
                "rich mail merge fragment relationship {relationship_id} is missing"
            ))
        })?;
        if !crate::document::relationship_is_internal(relationship) {
            if !allow_external_relationships {
                return Err(Error::Other(
                    "rich mail merge fragment has a non-internal relationship".to_owned(),
                ));
            }
            external_imports.push((
                true,
                destination_part.to_owned(),
                relationship_id,
                relationship.clone(),
            ));
            continue;
        }
        let source_part =
            OpcPackage::resolve_rel_target(&fragment.doc_part_name, &relationship.target);
        imports.push((
            true,
            destination_part.to_owned(),
            relationship_id,
            relationship.clone(),
            source_part,
        ));
    }
    let mut binding_stores = Vec::new();
    for store_item_id in fragment_store_item_ids(&dependency_xml)? {
        let source_part = crate::content_control::resolve_custom_xml_part(
            &fragment.package,
            &fragment.doc_part_name,
            &store_item_id,
        )?;
        let relationship = source_rels.items.iter().find(|relationship| {
            relationship.rel_type.ends_with("/customXml")
                && crate::document::relationship_is_internal(relationship)
                && OpcPackage::resolve_rel_target(&fragment.doc_part_name, &relationship.target)
                    == source_part
        }).ok_or_else(|| Error::Other(format!(
            "document fragment custom XML store {store_item_id} has no main-part relationship"
        )))?;
        if !imports
            .iter()
            .any(|(_, _, id, _, _)| id == &relationship.id)
        {
            imports.push((
                true,
                document.doc_part_name.clone(),
                relationship.id.clone(),
                relationship.clone(),
                source_part.clone(),
            ));
        }
        binding_stores.push((store_item_id, source_part));
    }
    if let Some(comment_xml) = Document::fragment_comment_dependency_xml(
        fragment.comments.as_ref(),
        fragment.comments_extended.as_ref(),
        &fragment_identity_values.comment_ids,
    )? {
        for store_item_id in fragment_store_item_ids(&comment_xml)? {
            let source_item_part = crate::content_control::resolve_custom_xml_part(
                &fragment.package,
                &fragment.doc_part_name,
                &store_item_id,
            )?;
            if !binding_stores.iter().any(|(id, _)| id == &store_item_id) {
                binding_stores.push((store_item_id, source_item_part));
            }
        }
        let destination_owner = document.ensure_fragment_comment_models_staged()?;
        let source_owner = fragment.comments_part_name.as_ref().ok_or_else(|| {
            Error::Other("document fragment comments part name is missing".to_owned())
        })?;
        let source_rels = fragment
            .package
            .get_part_rels(source_owner)
            .cloned()
            .unwrap_or_default();
        for relationship_id in relationship_ids_in_xml(&comment_xml)? {
            let relationship = source_rels.get_by_id(&relationship_id).ok_or_else(|| {
                Error::Other(format!(
                    "document fragment comment relationship {relationship_id} is missing"
                ))
            })?;
            if !crate::document::relationship_is_internal(relationship) {
                if !allow_external_relationships {
                    return Err(Error::Other(
                        "rich mail merge fragment has a non-internal relationship".to_owned(),
                    ));
                }
                external_imports.push((
                    false,
                    destination_owner.clone(),
                    relationship_id,
                    relationship.clone(),
                ));
                continue;
            }
            let source_part = OpcPackage::resolve_rel_target(source_owner, &relationship.target);
            imports.push((
                false,
                destination_owner.clone(),
                relationship_id,
                relationship.clone(),
                source_part,
            ));
        }
    }
    let mut note_copies: Vec<FragmentNoteCopy> = Vec::new();
    let mut note_relationship_imports = Vec::new();
    let mut note_external_imports = Vec::new();
    let mut pending_notes = fragment_note_references(&dependency_xml)?;
    if let Some(comments) = fragment.comments.as_ref()
        && let Some(xml) = Document::fragment_comment_dependency_xml(
            Some(comments),
            fragment.comments_extended.as_ref(),
            &fragment_identity_values.comment_ids,
        )?
    {
        pending_notes.extend(fragment_note_references(&xml)?);
    }
    let mut seen_notes = HashSet::new();
    let mut reserved_note_ids: HashMap<StoryKind, HashSet<i32>> = HashMap::new();
    let mut note_id_maps: HashMap<StoryKind, BTreeMap<String, String>> = HashMap::new();
    let mut note_index = 0;
    while note_index < pending_notes.len() {
        let (kind, old_id) = pending_notes[note_index];
        note_index += 1;
        if !seen_notes.insert((kind, old_id)) {
            continue;
        }
        let (source_part, note_xml) = fragment.fragment_note_dependency(kind, old_id)?;
        for store_item_id in fragment_store_item_ids(&note_xml)? {
            let source_item_part = crate::content_control::resolve_custom_xml_part(
                &fragment.package,
                &fragment.doc_part_name,
                &store_item_id,
            )?;
            if !binding_stores.iter().any(|(id, _)| id == &store_item_id) {
                binding_stores.push((store_item_id, source_item_part));
            }
        }
        let destination_note_part = document.ensure_fragment_note_part_staged(kind)?;
        let current_notes = CT_Footnotes::from_xml(
            document
                .package
                .get_part(&destination_note_part)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "document fragment destination note part {destination_note_part} is missing"
                    ))
                })?,
        )?;
        let source_notes =
            CT_Footnotes::from_xml(fragment.package.get_part(&source_part).ok_or_else(|| {
                Error::Other(format!(
                    "document fragment source note part {source_part} is missing"
                ))
            })?)?;
        let used = reserved_note_ids.entry(kind).or_insert_with(|| {
            current_notes
                .footnotes
                .iter()
                .chain(&source_notes.footnotes)
                .map(|note| note.id)
                .collect()
        });
        let mut new_id = 2_i32;
        while used.contains(&new_id) {
            new_id = new_id.checked_add(1).ok_or_else(|| {
                Error::Other(format!("document fragment {kind:?} ID range is exhausted"))
            })?;
        }
        used.insert(new_id);
        note_copy_sources.push((
            kind,
            source_part.clone(),
            old_id,
            destination_note_part.clone(),
            new_id,
        ));
        note_id_maps
            .entry(kind)
            .or_default()
            .insert(old_id.to_string(), new_id.to_string());
        for related in fragment_note_references(&note_xml)? {
            if !seen_notes.contains(&related) && !pending_notes.contains(&related) {
                pending_notes.push(related);
            }
        }
        let source_note_rels = fragment
            .package
            .get_part_rels(&source_part)
            .cloned()
            .unwrap_or_default();
        for relationship_id in relationship_ids_in_xml(&note_xml)? {
            if note_relationship_imports.iter().any(
                |(owner, id, _, _): &(
                    String,
                    String,
                    oxml_opc::relationship::Relationship,
                    String,
                )| owner == &destination_note_part && id == &relationship_id,
            ) || note_external_imports.iter().any(
                |(owner, id, _): &(String, String, oxml_opc::relationship::Relationship)| {
                    owner == &destination_note_part && id == &relationship_id
                },
            ) {
                continue;
            }
            let relationship = source_note_rels
                .get_by_id(&relationship_id)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "document fragment {kind:?} relationship {relationship_id} is missing"
                    ))
                })?
                .clone();
            if crate::document::relationship_is_internal(&relationship) {
                let child = OpcPackage::resolve_rel_target(&source_part, &relationship.target);
                note_relationship_imports.push((
                    destination_note_part.clone(),
                    relationship_id,
                    relationship,
                    child,
                ));
            } else {
                if !allow_external_relationships {
                    return Err(Error::Other(
                        "rich mail merge fragment has a non-internal relationship".to_owned(),
                    ));
                }
                note_external_imports.push((
                    destination_note_part.clone(),
                    relationship_id,
                    relationship,
                ));
            }
        }
        note_copies.push(FragmentNoteCopy {
            kind,
            destination_part: destination_note_part,
            xml: note_xml,
            replace_id: None,
        });
    }
    fragment_xml = patch_fragment_note_ids(&fragment_xml, &note_id_maps)?;
    for note in &mut note_copies {
        note.xml = patch_fragment_note_ids(&note.xml, &note_id_maps)?;
    }
    if let Some(comments) = fragment.comments.as_mut() {
        *comments = CT_Comments::from_xml(&patch_fragment_note_ids(
            &comments.to_xml()?,
            &note_id_maps,
        )?)?;
    }
    for (_, source_part) in &binding_stores {
        if imports.iter().any(|(_, _, _, _, part)| part == source_part) {
            continue;
        }
        let relationship = source_rels
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type.ends_with("/customXml")
                    && crate::document::relationship_is_internal(relationship)
                    && OpcPackage::resolve_rel_target(&fragment.doc_part_name, &relationship.target)
                        == *source_part
            })
            .ok_or_else(|| {
                Error::Other(format!(
                    "document fragment custom XML store {source_part} has no main-part relationship"
                ))
            })?;
        // Binding stores are document-wide, even when selected content lives in another story.
        imports.push((
            true,
            document.doc_part_name.clone(),
            relationship.id.clone(),
            relationship.clone(),
            source_part.clone(),
        ));
    }
    let mut occupied_store_ids = binding_stores
        .iter()
        .map(|(id, _)| id.clone())
        .collect::<BTreeSet<_>>();
    for package in [&fragment.package, &document.package] {
        for (part, xml) in &package.parts {
            if package.content_types.content_type_for(part)
                == Some("application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                && let Some(id) = crate::content_control::parse_store_item_id(xml)?
            {
                occupied_store_ids.insert(crate::content_control::normalize_store_item_id(&id));
            }
        }
    }
    let mut binding_remaps = Vec::new();
    for (store_item_id, source_item_part) in &binding_stores {
        match crate::content_control::custom_xml_store_item_count(
            &document.package,
            &document.doc_part_name,
            store_item_id,
        )? {
            0 => {}
            1 => {
                let new_id = (1_u64..)
                    .map(|index| format!("{{F2760000-0000-4000-8000-{index:012X}}}"))
                    .find(|candidate| {
                        !occupied_store_ids
                            .contains(&crate::content_control::normalize_store_item_id(candidate))
                    })
                    .ok_or_else(|| {
                        Error::Other(
                            "document fragment custom XML store ID range is exhausted".to_owned(),
                        )
                    })?;
                occupied_store_ids.insert(crate::content_control::normalize_store_item_id(&new_id));
                binding_remaps.push((store_item_id.clone(), source_item_part.clone(), new_id));
            }
            _ => {
                return Err(Error::Other(format!(
                    "document fragment custom XML store {store_item_id} is ambiguous in the destination"
                )));
            }
        }
    }
    let mut remapped_store_properties = HashSet::new();
    for (_, item_part, _) in &binding_remaps {
        if let Some(rels) = fragment.package.get_part_rels(item_part) {
            for relationship in &rels.items {
                if relationship.rel_type.ends_with("/customXmlProps")
                    && crate::document::relationship_is_internal(relationship)
                {
                    remapped_store_properties.insert(OpcPackage::resolve_rel_target(
                        item_part,
                        &relationship.target,
                    ));
                }
            }
        }
    }
    let mut closure = HashSet::new();
    for (_, _, _, _, source_part) in &imports {
        discover_fragment_part_closure(&fragment.package, source_part, &mut closure)?;
    }
    for (_, _, _, source_part) in &note_relationship_imports {
        discover_fragment_part_closure(&fragment.package, source_part, &mut closure)?;
    }
    let mut closure = closure.into_iter().collect::<Vec<_>>();
    closure.sort();
    if !allow_external_relationships
        && closure.iter().any(|part| {
            fragment
                .package
                .get_part_rels(part)
                .is_some_and(|relationships| {
                    relationships.items.iter().any(|relationship| {
                        !crate::document::relationship_is_internal(relationship)
                    })
                })
        })
    {
        return Err(Error::Other(
            "rich mail merge fragment has a non-internal relationship".to_owned(),
        ));
    }

    for source_part in closure {
        let destination_part = if policy.reuse_related_parts()
            && !remapped_store_properties.contains(&source_part)
        {
            equivalent_fragment_leaf_part(&fragment.package, document, &source_part).map_or_else(
                || {
                    document
                        .identifiers
                        .reserve_fragment_part_name(&source_part)
                },
                Ok,
            )?
        } else {
            document
                .identifiers
                .reserve_fragment_part_name(&source_part)?
        };
        part_map.insert(source_part, destination_part);
    }
    for (source_part, destination_part) in &part_map {
        let payload = fragment.package.get_part(source_part).ok_or_else(|| {
            Error::Other(format!(
                "document fragment reachable part {source_part} is missing"
            ))
        })?;
        if let Some(relationships) = fragment.package.get_part_rels(source_part) {
            for relationship in &relationships.items {
                if !crate::document::relationship_is_internal(relationship) {
                    continue;
                }
                let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
                let copied_target = part_map.get(&target).ok_or_else(|| {
                    Error::Other(format!("fragment target {target} was not reserved"))
                })?;
                let new_target = relative_fragment_target(destination_part, copied_target);
                if new_target != relationship.target
                    && !relationship.target.is_empty()
                    && payload
                        .windows(relationship.target.len())
                        .any(|window| window == relationship.target.as_bytes())
                {
                    return Err(Error::Other(format!(
                        "document fragment part {source_part} embeds renamed target {}; opaque payload cannot be rewritten safely",
                        relationship.target
                    )));
                }
            }
        }
        for (old, new) in &part_map {
            if old != new
                && payload
                    .windows(old.len())
                    .any(|window| window == old.as_bytes())
            {
                return Err(Error::Other(format!(
                    "document fragment part {source_part} embeds renamed part name {old}; opaque payload cannot be rewritten safely for {destination_part}"
                )));
            }
        }
    }
    let mut copied = HashSet::new();
    for (selected_owner, destination_owner, relationship_id, relationship, source_part) in imports {
        let destination_part = copy_fragment_part(
            &fragment.package,
            document,
            &source_part,
            &part_map,
            &mut copied,
        )?;
        let target = relative_fragment_target(&destination_owner, &destination_part);
        let destination_id = document
            .add_internal_relationship_checked(
                &destination_owner,
                &relationship.rel_type,
                &target,
            )
            .map_err(|error| {
                Error::Other(format!(
                    "rich mail merge relationship allocation failed for {destination_owner}: {error}"
                ))
            })?;
        if selected_owner {
            relationship_map.insert(relationship_id, destination_id);
        } else {
            comment_relationship_map.insert(relationship_id, destination_id);
        }
    }
    for (selected_owner, destination_owner, relationship_id, relationship) in external_imports {
        let destination_id = document.add_external_relationship_checked(
            &destination_owner,
            &relationship.rel_type,
            &relationship.target,
        )?;
        if selected_owner {
            relationship_map.insert(relationship_id, destination_id);
        } else {
            comment_relationship_map.insert(relationship_id, destination_id);
        }
    }
    let mut note_relationship_maps: HashMap<String, BTreeMap<String, String>> = HashMap::new();
    for (destination_owner, old_id, relationship, source_part) in note_relationship_imports {
        let copied_part = copy_fragment_part(
            &fragment.package,
            document,
            &source_part,
            &part_map,
            &mut copied,
        )?;
        let new_id = document.add_relative_internal_relationship_checked(
            &destination_owner,
            &relationship.rel_type,
            &copied_part,
        )?;
        note_relationship_maps
            .entry(destination_owner)
            .or_default()
            .insert(old_id, new_id);
    }
    for (destination_owner, old_id, relationship) in note_external_imports {
        let new_id = document.add_external_relationship_checked(
            &destination_owner,
            &relationship.rel_type,
            &relationship.target,
        )?;
        note_relationship_maps
            .entry(destination_owner)
            .or_default()
            .insert(old_id, new_id);
    }
    for (store_item_id, source_item_part, new_id) in binding_remaps {
        let item_relationships = fragment.package.get_part_rels(&source_item_part).ok_or_else(|| Error::Other(format!(
                    "document fragment custom XML store {store_item_id} has no item properties relationship"
                )))?;
        let props_relationship = item_relationships.items.iter().find(|relationship| relationship.rel_type.ends_with("/customXmlProps") && crate::document::relationship_is_internal(relationship)).ok_or_else(|| Error::Other(format!(
                    "document fragment custom XML store {store_item_id} has no item properties relationship"
                )))?;
        let source_props =
            OpcPackage::resolve_rel_target(&source_item_part, &props_relationship.target);
        let destination_props = part_map.get(&source_props).ok_or_else(|| {
            Error::Other(format!(
                "document fragment custom XML properties {source_props} were not copied"
            ))
        })?;
        if fragment_store_item_ids(&fragment_xml)?.contains(&store_item_id) {
            fragment_xml = replace_fragment_store_id(
                &fragment_xml,
                &store_item_id,
                &new_id,
                b"dataBinding",
                b"storeItemID",
            )?;
        }
        for note in &mut note_copies {
            if fragment_store_item_ids(&note.xml)?.contains(&store_item_id) {
                note.xml = replace_fragment_store_id(
                    &note.xml,
                    &store_item_id,
                    &new_id,
                    b"dataBinding",
                    b"storeItemID",
                )?;
            }
        }
        if let Some(comments) = fragment.comments.as_mut() {
            let comments_xml = comments.to_xml()?;
            if fragment_store_item_ids(&comments_xml)?.contains(&store_item_id) {
                *comments = CT_Comments::from_xml(&replace_fragment_store_id(
                    &comments_xml,
                    &store_item_id,
                    &new_id,
                    b"dataBinding",
                    b"storeItemID",
                )?)?;
            }
        }
        let props_xml = document
            .package
            .get_part(destination_props)
            .ok_or_else(|| {
                Error::Other(format!(
                    "document fragment custom XML properties {destination_props} are missing"
                ))
            })?;
        let patched_props = replace_fragment_store_id(
            props_xml,
            &store_item_id,
            &new_id,
            b"datastoreItem",
            b"itemID",
        )?;
        document.package.set_part(destination_props, patched_props);
    }
    if !relationship_map.is_empty() {
        fragment_xml = patch_relationship_ids(&fragment_xml, &relationship_map)?;
    }
    fragment.document = CT_Document::from_xml(&fragment_xml)?;
    let mut identity_xml = fragment_xml.clone();
    for note in &note_copies {
        validate_fragment_identity_ownership(&body_identity_values(&wrap_fragment_companion(
            &note.xml,
        ))?)?;
        identity_xml.extend_from_slice(&note.xml);
    }
    if let Some(comments) = Document::fragment_comment_dependency_xml(
        fragment.comments.as_ref(),
        fragment.comments_extended.as_ref(),
        &fragment_identity_values.comment_ids,
    )? {
        identity_xml.extend(comments);
    }
    let (_, mut identity_remap) = remap_fragment_xml_identities(
        &wrap_fragment_companion(&identity_xml),
        &mut document.identifiers,
        identity_state,
        false,
    )?;
    let patched_comments = Document::fragment_comment_dependency_xml(
        fragment.comments.as_ref(),
        fragment.comments_extended.as_ref(),
        &fragment_identity_values.comment_ids,
    )?
    .map(|xml| -> Result<CT_Comments> {
        let xml = patch_relationship_ids(&xml, &comment_relationship_map)?;
        let wrapped =
            patch_body_identity_attributes(&wrap_fragment_companion(&xml), &identity_remap)?;
        let xml = unwrap_fragment_companion(&wrapped);
        Ok(CT_Comments::from_xml(&freshen_fragment_marker_ids(
            document, &xml,
        )?)?)
    })
    .transpose()?;
    let comment_ids = document.import_fragment_comments_staged(
        patched_comments.as_ref().or(fragment.comments.as_ref()),
        fragment.comments_extended.as_ref(),
        &fragment_identity_values.comment_ids,
    )?;
    if !comment_ids.is_empty() {
        let comment_remap = BodyIdentityRemap {
            comment_ids: comment_ids.clone(),
            ..Default::default()
        };
        fragment_xml = patch_body_identity_attributes(&fragment_xml, &comment_remap)?;
        for note in &mut note_copies {
            note.xml = unwrap_fragment_companion(&patch_body_identity_attributes(
                &wrap_fragment_companion(&note.xml),
                &comment_remap,
            )?);
        }
        let updated = patch_body_identity_attributes(&fragment.document.to_xml()?, &comment_remap)?;
        fragment.document = CT_Document::from_xml(&updated)?;
    }
    identity_remap.drop_paragraph_identities = true;
    fragment_xml = patch_body_identity_attributes(
        &uniquify_drawing_ids_in_xml(&fragment_xml)?,
        &identity_remap,
    )?;
    fragment.document = CT_Document::from_xml(&fragment_xml)?;
    let insert_at = document.document.body.content.len();
    let numbering_remap = document.insert_document_fragment_content_staged(
        insert_at,
        &fragment,
        policy.reuse_numbering(),
        source_owned_numbering,
    )?;
    let typed = document.document.body.content.drain(insert_at..).collect();
    let numbering_values = numbering_remap
        .iter()
        .map(|(old, new)| (old.to_string(), new.to_string()))
        .collect::<BTreeMap<_, _>>();
    patch_word_values(&mut fragment_xml, &[b"numId"], &numbering_values)?;
    if let Some(comments) = document.comments.as_mut() {
        for comment in &mut comments.comments {
            if !comment_ids.values().any(|id| *id == comment.id.to_string()) {
                continue;
            }
            let mut selected = CT_Comments::new();
            selected.comments.push(comment.clone());
            let mut xml = selected.to_xml()?;
            patch_word_values(&mut xml, &[b"numId"], &numbering_values)?;
            *comment = CT_Comments::from_xml(&xml)?.comments.remove(0);
        }
    }
    for note in note_copies {
        let mut xml = match note_relationship_maps.get(&note.destination_part) {
            Some(map) => patch_relationship_ids(&note.xml, map)?,
            None => note.xml,
        };
        patch_word_values(&mut xml, &[b"numId"], &numbering_values)?;
        let wrapped =
            patch_body_identity_attributes(&wrap_fragment_companion(&xml), &identity_remap)?;
        let xml = freshen_fragment_marker_ids(document, &unwrap_fragment_companion(&wrapped))?;
        document.publish_fragment_note_staged(
            note.kind,
            &note.destination_part,
            xml,
            note.replace_id,
        )?;
    }
    let xml = freshen_fragment_marker_ids(document, &fragment_xml)?;
    Ok(ImportedFragmentContent { typed, xml })
}

fn equivalent_fragment_leaf_part(
    source: &OpcPackage,
    destination: &Document,
    source_part: &str,
) -> Option<String> {
    if source
        .get_part_rels(source_part)
        .is_some_and(|relationships| !relationships.items.is_empty())
    {
        return None;
    }
    let source_bytes = source.get_part(source_part)?;
    let source_content_type = source.content_types.content_type_for(source_part)?;
    destination
        .package
        .parts
        .iter()
        .find_map(|(part_name, bytes)| {
            (bytes.as_slice() == source_bytes
                && destination
                    .package
                    .get_part_rels(part_name)
                    .is_none_or(|relationships| relationships.items.is_empty())
                && destination
                    .package
                    .content_types
                    .content_type_for(part_name)
                    == Some(source_content_type))
            .then(|| part_name.clone())
        })
}

fn prune_document_fragment_dependencies(fragment: &mut Document, selected: &[u8]) -> Result<()> {
    let mut body_xml = selected.to_vec();
    for companion in fragment_dependency_companions(fragment, selected)? {
        body_xml.extend(companion);
    }
    let mut used_numbering = word_value_attributes(&body_xml, &[b"numId"])?
        .into_iter()
        .filter_map(|value| value.parse::<u32>().ok())
        .collect::<HashSet<_>>();
    let mut used_styles = word_value_attributes(&body_xml, &[b"pStyle", b"rStyle", b"tblStyle"])?;
    used_styles.extend(
        fragment
            .styles
            .styles
            .iter()
            .filter(|style| style.is_default)
            .map(|style| style.style_id.clone()),
    );
    loop {
        let before = (used_styles.len(), used_numbering.len());
        for style in &fragment.styles.styles {
            if !used_styles.contains(&style.style_id) {
                continue;
            }
            used_styles.extend(
                [
                    style.based_on.as_ref(),
                    style.next_style.as_ref(),
                    style.linked_style.as_ref(),
                ]
                .into_iter()
                .flatten()
                .cloned(),
            );
        }
        let mut retained_styles = fragment.styles.clone();
        retained_styles
            .styles
            .retain(|style| used_styles.contains(&style.style_id));
        used_numbering.extend(
            word_value_attributes(&retained_styles.to_xml()?, &[b"numId"])?
                .into_iter()
                .filter_map(|value| value.parse::<u32>().ok()),
        );
        if let Some(numbering) = &fragment.numbering {
            let mut retained_numbering = numbering.clone();
            retained_numbering
                .nums
                .retain(|instance| used_numbering.contains(&instance.num_id));
            let used_abstract = retained_numbering
                .nums
                .iter()
                .map(|instance| instance.abstract_num_id)
                .collect::<HashSet<_>>();
            retained_numbering
                .abstract_nums
                .retain(|definition| used_abstract.contains(&definition.abstract_num_id));
            used_styles.extend(word_value_attributes(
                &retained_numbering.to_xml()?,
                &[b"pStyle", b"rStyle", b"styleLink", b"numStyleLink"],
            )?);
        }
        if (used_styles.len(), used_numbering.len()) == before {
            break;
        }
    }
    fragment
        .styles
        .styles
        .retain(|style| used_styles.contains(&style.style_id));
    if let Some(numbering) = fragment.numbering.as_mut() {
        numbering
            .nums
            .retain(|instance| used_numbering.contains(&instance.num_id));
        let used_abstract = numbering
            .nums
            .iter()
            .map(|instance| instance.abstract_num_id)
            .collect::<HashSet<_>>();
        numbering
            .abstract_nums
            .retain(|definition| used_abstract.contains(&definition.abstract_num_id));
        if numbering.nums.is_empty() {
            fragment.numbering = None;
        }
    }
    Ok(())
}

fn word_value_attributes(xml: &[u8], elements: &[&[u8]]) -> Result<HashSet<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut values = HashSet::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid document fragment dependency XML: {error}"))
        })? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace)
                    && elements.contains(&local.as_ref())
                    && let Some((_, value)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"val",
                        AttributeNamespace::Word,
                    )?
                {
                    values.insert(value);
                }
            }
            Event::Eof => return Ok(values),
            _ => {}
        }
        buffer.clear();
    }
}

// The final direct owned numId is modeled. Earlier duplicates and Word-looking
// descendants below opaque owners remain exact. Only the writer's redundant
// canonical Word closure on the selected pPr, numPr and numId is comparison
// bookkeeping.
fn fragment_style_numbering_projection(
    style: &rdocx_oxml::styles::CT_Style,
    mapped_num_id: Option<u32>,
) -> Result<Vec<u8>> {
    let mut writer = quick_xml::Writer::new(Vec::new());
    style.to_xml(&mut writer)?;
    let xml = writer.into_inner();
    let prefix =
        b"<w:fragment xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">";
    let mut wrapped = prefix.to_vec();
    wrapped.extend_from_slice(&xml);
    wrapped.extend_from_slice(b"</w:fragment>");
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut buffer = Vec::new();
    struct ProjectionFrame {
        owned: bool,
        word_prefix: bool,
        closure: Option<(usize, usize)>,
    }
    let mut path: Vec<ProjectionFrame> = Vec::new();
    let expected: &[&[u8]] = &[b"fragment", b"style", b"pPr", b"numPr", b"numId"];
    let mut selected = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid fragment style numbering XML: {error}"))
        })?;
        let is_start = matches!(&event, Event::Start(_));
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(node) | Event::Empty(node) => {
                let (namespace, local) = reader.resolver().resolve_element(node.name());
                let owned = path.iter().all(|entry| entry.owned)
                    && expected
                        .get(path.len())
                        .is_some_and(|expected| *expected == local.as_ref())
                    && namespace_is_word(&namespace);
                let (w_namespace, _) = reader
                    .resolver()
                    .resolve_element(quick_xml::name::QName(b"w:reference"));
                let w_is_word = namespace_is_word(&w_namespace);
                let mut closure = None;
                if owned && matches!(path.len(), 2..=4)
                    && path.last().is_some_and(|parent| parent.word_prefix)
                    && let Some((name_start, _value_start, value_end)) = attribute_source_span(&wrapped[before..after], b"xmlns:w")
                    && &wrapped[before + name_start..before + value_end + 1]
                        == b"xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                {
                    let mut start = before + name_start;
                    while start > before && wrapped[start - 1].is_ascii_whitespace() { start -= 1; }
                    closure = Some((start - prefix.len(), before + value_end + 1 - prefix.len()));
                }
                if owned && path.len() == 4 {
                    let (key, value) = resolved_element_attribute(
                        &node,
                        reader.resolver(),
                        b"val",
                        AttributeNamespace::Word,
                    )?
                    .ok_or_else(|| {
                        Error::Other("fragment style numbering value is missing".into())
                    })?;
                    let (start, end) = attribute_value_span(&wrapped[before..after], &key)
                        .ok_or_else(|| {
                            Error::Other(
                                "fragment style numbering attribute source was not found".into(),
                            )
                        })?;
                    let mut closures = path
                        .iter()
                        .filter_map(|frame| frame.closure)
                        .collect::<Vec<_>>();
                    if let Some(closure) = closure {
                        closures.push(closure);
                    }
                    selected = Some((
                        value,
                        before + start - prefix.len(),
                        before + end - prefix.len(),
                        closures,
                    ));
                }
                if is_start {
                    path.push(ProjectionFrame {
                        owned,
                        word_prefix: w_is_word,
                        closure,
                    });
                }
            }
            Event::End(_) => {
                path.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut edits = Vec::new();
    if let Some((value, start, end, closures)) = selected {
        if Some(
            value
                .parse::<u32>()
                .map_err(|_| Error::Other("invalid fragment style numbering ID".into()))?,
        ) != style.ppr.as_ref().and_then(|properties| properties.num_id)
        {
            return Err(Error::Other(
                "fragment style numbering projection disagrees with its owned reference".into(),
            ));
        }
        if let Some(id) = mapped_num_id {
            edits.push((start, end, id.to_string().into_bytes()));
        }
        for (start, end) in closures {
            edits.push((start, end, Vec::new()));
        }
    } else if mapped_num_id.is_some() {
        return Err(Error::Other(
            "fragment style has no owned numbering reference".into(),
        ));
    }
    let mut xml = xml;
    edits.sort_by_key(|edit| edit.0);
    for (start, end, replacement) in edits.into_iter().rev() {
        xml.splice(start..end, replacement);
    }
    Ok(xml)
}

fn remap_fragment_style_collisions(
    destination: &Document,
    fragment: &mut Document,
    fragment_xml: &mut Vec<u8>,
    reuse_equivalent: bool,
    reuse_equivalent_numbering: bool,
    source_owned_numbering: bool,
) -> Result<bool> {
    let mut replacements = BTreeMap::new();
    let mut used = destination
        .styles
        .styles
        .iter()
        .chain(fragment.styles.styles.iter())
        .map(|style| style.style_id.clone())
        .collect::<HashSet<_>>();
    for style in &fragment.styles.styles {
        let Some(existing) = destination.styles.get_by_id(&style.style_id) else {
            continue;
        };
        let source_num_id = style.ppr.as_ref().and_then(|properties| properties.num_id);
        let reusable_num_id = source_num_id
            .filter(|_| reuse_equivalent_numbering)
            .and_then(|num_id| {
                if source_owned_numbering {
                    Some(num_id)
                } else {
                    destination.equivalent_numbering_dependency_id(fragment, num_id)
                }
            });
        let numbering_is_reusable = source_num_id.is_none() || reusable_num_id.is_some();
        // Compare the source-owned projection with only the selected numbering
        // reference remapped. Preservation snapshots and generated namespace
        // closures are not new style facts, while opaque XML remains exact.
        let styles_match = if existing == style {
            true
        } else if let Some(num_id) = reusable_num_id {
            fragment_style_numbering_projection(existing, None)?
                == fragment_style_numbering_projection(style, Some(num_id))?
        } else {
            false
        };
        if styles_match && (reuse_equivalent || style.is_default) && numbering_is_reusable {
            continue;
        }
        let mut ordinal = 1u64;
        loop {
            let candidate = format!("{}Merge{ordinal}", style.style_id);
            if used.insert(candidate.clone()) {
                replacements.insert(style.style_id.clone(), candidate);
                break;
            }
            ordinal = ordinal.checked_add(1).ok_or_else(|| {
                Error::Other("rich mail merge fragment style name space is exhausted".to_owned())
            })?;
        }
    }
    if replacements.is_empty() {
        return Ok(false);
    }
    for style in &mut fragment.styles.styles {
        if let Some(replacement) = replacements.get(&style.style_id) {
            style.style_id = replacement.clone();
        }
        if let Some(replacement) = style
            .based_on
            .as_ref()
            .and_then(|style_id| replacements.get(style_id))
        {
            style.based_on = Some(replacement.clone());
        }
        if let Some(replacement) = style
            .next_style
            .as_ref()
            .and_then(|style_id| replacements.get(style_id))
        {
            style.next_style = Some(replacement.clone());
        }
        if let Some(replacement) = style
            .linked_style
            .as_ref()
            .and_then(|style_id| replacements.get(style_id))
        {
            style.linked_style = Some(replacement.clone());
        }
        for (_, raw) in &mut style.extra_xml {
            for (old, new) in &replacements {
                patch_word_value(raw, b"link", old, new)?;
            }
        }
    }
    if let Some(numbering) = &mut fragment.numbering {
        let mut xml = numbering.to_xml()?;
        patch_word_values(
            &mut xml,
            &[b"styleLink", b"numStyleLink", b"pStyle"],
            &replacements,
        )?;
        *numbering = rdocx_oxml::numbering::CT_Numbering::from_xml(&xml)?;
    }
    for (old, new) in &replacements {
        for element in [b"pStyle".as_slice(), b"rStyle", b"tblStyle"] {
            patch_word_value(fragment_xml, element, old, new)?;
        }
    }
    for kind in [StoryKind::Footnote, StoryKind::Endnote] {
        let stories = fragment.stories()?;
        if let Some(story) = stories.iter().find(|story| story.kind() == kind) {
            let part = story.part_name().to_owned();
            let mut xml = fragment
                .package
                .get_part(&part)
                .ok_or_else(|| Error::Other(format!("fragment note part {part} is missing")))?
                .to_vec();
            for (old, new) in &replacements {
                for element in [b"pStyle".as_slice(), b"rStyle", b"tblStyle"] {
                    patch_word_value(&mut xml, element, old, new)?;
                }
            }
            fragment.package.set_part(&part, xml);
        }
    }
    if let Some(comments) = fragment.comments.as_mut() {
        let mut xml = comments.to_xml()?;
        for (old, new) in &replacements {
            for element in [b"pStyle".as_slice(), b"rStyle", b"tblStyle"] {
                patch_word_value(&mut xml, element, old, new)?;
            }
        }
        *comments = CT_Comments::from_xml(&xml)?;
    }
    let styles_part = fragment
        .package
        .get_part_rels(&fragment.doc_part_name)
        .and_then(|relationships| {
            relationships.items.iter().find(|relationship| {
                relationship.rel_type == rel_types::STYLES
                    && crate::document::relationship_is_internal(relationship)
            })
        })
        .map(|relationship| {
            OpcPackage::resolve_rel_target(&fragment.doc_part_name, &relationship.target)
        })
        .ok_or_else(|| Error::Other("rich mail merge fragment has no styles part".to_owned()))?;
    fragment
        .package
        .set_part(&styles_part, fragment.styles.to_xml()?);
    if let Some(numbering) = &fragment.numbering {
        let numbering_part = fragment
            .package
            .get_part_rels(&fragment.doc_part_name)
            .and_then(|relationships| {
                relationships.items.iter().find(|relationship| {
                    relationship.rel_type == rel_types::NUMBERING
                        && crate::document::relationship_is_internal(relationship)
                })
            })
            .map(|relationship| {
                OpcPackage::resolve_rel_target(&fragment.doc_part_name, &relationship.target)
            })
            .ok_or_else(|| {
                Error::Other("rich mail merge fragment has no numbering part".to_owned())
            })?;
        fragment
            .package
            .set_part(&numbering_part, numbering.to_xml()?);
    }
    Ok(true)
}

fn remap_rich_body_content(
    document: &mut Document,
    content: &mut Vec<BodyContent>,
    identity_state: &mut BodyIdentityState,
) -> Result<()> {
    let mut occurrence = document.clone_for_staging();
    occurrence.document.body.content = std::mem::take(content);
    occurrence.document.body.sect_pr = None;
    remap_body_identities(&mut occurrence, &mut document.identifiers, identity_state)?;
    *content = occurrence.document.body.content;
    Ok(())
}

fn remap_rich_rows(
    document: &mut Document,
    rows: &mut [RichExpandedRow],
    identity_state: &mut BodyIdentityState,
) -> Result<()> {
    let mut table = CT_Tbl::new();
    table.rows = rows
        .iter_mut()
        .map(|item| std::mem::take(&mut item.row))
        .collect();
    let mut content = vec![BodyContent::Table(Box::new(table))];
    remap_rich_body_content(document, &mut content, identity_state)?;
    let Some(BodyContent::Table(table)) = content.pop() else {
        return Err(Error::Other(
            "rich mail merge row identity staging lost its table".to_owned(),
        ));
    };
    for (item, row) in rows.iter_mut().zip(table.rows) {
        item.row = row;
    }
    Ok(())
}

fn patch_word_value(xml: &mut Vec<u8>, element_local: &[u8], old: &str, new: &str) -> Result<()> {
    patch_word_values(
        xml,
        &[element_local],
        &BTreeMap::from([(old.to_owned(), new.to_owned())]),
    )
}

fn patch_word_values(
    xml: &mut Vec<u8>,
    elements: &[&[u8]],
    replacements: &BTreeMap<String, String>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml.as_slice());
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "invalid rich mail merge style reference XML: {error}"
            ))
        })?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(node) | Event::Empty(node) => {
                let (namespace, local) = reader.resolver().resolve_element(node.name());
                if namespace_is_word(&namespace)
                    && elements.contains(&local.as_ref())
                    && let Some((key, value)) = resolved_element_attribute(
                        &node,
                        reader.resolver(),
                        b"val",
                        AttributeNamespace::Word,
                    )?
                    && let Some(new) = replacements.get(&value)
                {
                    let Some((relative_start, relative_end)) =
                        attribute_value_span(&xml[before..after], &key)
                    else {
                        return Err(Error::Other(
                            "rich mail merge style reference attribute source was not found"
                                .to_owned(),
                        ));
                    };
                    edits.push(FieldSourceEdit {
                        start: before + relative_start,
                        end: before + relative_end,
                        replacement: new.as_bytes().to_vec(),
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    edits.sort_by_key(|edit| edit.start);
    for edit in edits.into_iter().rev() {
        xml.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(())
}

// Companion roots are namespace-complete, so the existing body identity scanner can
// inspect them without changing their own namespace bindings.
fn wrap_fragment_companion(xml: &[u8]) -> Vec<u8> {
    let mut wrapped = b"<f276:document xmlns:f276=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><f276:body>".to_vec();
    wrapped.extend_from_slice(xml);
    wrapped.extend_from_slice(b"</f276:body></f276:document>");
    wrapped
}

fn unwrap_fragment_companion(xml: &[u8]) -> Vec<u8> {
    let prefix_len = wrap_fragment_companion(&[]).len() - b"</f276:body></f276:document>".len();
    xml[prefix_len..xml.len() - b"</f276:body></f276:document>".len()].to_vec()
}

fn fragment_dependency_companions(fragment: &Document, selected: &[u8]) -> Result<Vec<Vec<u8>>> {
    let mut pending = vec![selected.to_vec()];
    let mut companions = Vec::new();
    let mut seen_notes = HashSet::new();
    let mut seen_comments = HashSet::new();
    let mut index = 0;
    while index < pending.len() {
        let xml = pending[index].clone();
        index += 1;
        for (kind, id) in fragment_note_references(&xml)? {
            if seen_notes.insert((kind, id)) {
                let (_, note) = fragment.fragment_note_dependency(kind, id)?;
                pending.push(note.clone());
                companions.push(note);
            }
        }
        let ids = body_identity_values(&wrap_fragment_companion(&xml))?
            .comment_ids
            .into_iter()
            .filter(|id| seen_comments.insert(id.clone()))
            .collect::<Vec<_>>();
        if let Some(comments) = Document::fragment_comment_dependency_xml(
            fragment.comments.as_ref(),
            fragment.comments_extended.as_ref(),
            &ids,
        )? {
            pending.push(comments.clone());
            companions.push(comments);
        }
    }
    Ok(companions)
}

fn patch_fragment_note_ids(
    xml: &[u8],
    maps: &HashMap<StoryKind, BTreeMap<String, String>>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid fragment note XML: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace) {
                    let kind = match local.as_ref() {
                        b"footnote" | b"footnoteReference" => Some(StoryKind::Footnote),
                        b"endnote" | b"endnoteReference" => Some(StoryKind::Endnote),
                        _ => None,
                    };
                    if let Some(map) = kind.and_then(|kind| maps.get(&kind)) {
                        add_identity_attribute_edit(
                            xml,
                            before,
                            after,
                            &element,
                            reader.resolver(),
                            b"id",
                            AttributeNamespace::Word,
                            map,
                            &mut edits,
                        )?;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut updated = xml.to_vec();
    for edit in edits.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

fn fragment_marker_family(local: &[u8]) -> Option<(&'static str, u8)> {
    match local {
        b"permStart" => Some(("permission", 0)),
        b"permEnd" => Some(("permission", 1)),
        b"moveFromRangeStart" => Some(("move-from", 0)),
        b"moveFromRangeEnd" => Some(("move-from", 1)),
        b"moveToRangeStart" => Some(("move-to", 0)),
        b"moveToRangeEnd" => Some(("move-to", 1)),
        b"customXmlInsRangeStart" => Some(("custom-ins", 0)),
        b"customXmlInsRangeEnd" => Some(("custom-ins", 1)),
        b"customXmlDelRangeStart" => Some(("custom-del", 0)),
        b"customXmlDelRangeEnd" => Some(("custom-del", 1)),
        b"customXmlMoveFromRangeStart" => Some(("custom-move-from", 0)),
        b"customXmlMoveFromRangeEnd" => Some(("custom-move-from", 1)),
        b"customXmlMoveToRangeStart" => Some(("custom-move-to", 0)),
        b"customXmlMoveToRangeEnd" => Some(("custom-move-to", 1)),
        b"ins" | b"del" | b"moveFrom" | b"moveTo" | b"pPrChange" | b"rPrChange"
        | b"tblPrChange" | b"trPrChange" | b"tcPrChange" | b"sectPrChange" | b"tblGridChange"
        | b"numberingChange" => Some(("revision", 2)),
        _ => None,
    }
}

fn freshen_fragment_marker_ids(document: &Document, xml: &[u8]) -> Result<Vec<u8>> {
    let mut occupied = HashSet::new();
    // All Word IDs form a conservative occupied set. This also avoids aliases
    // between the separately scoped annotation families.
    for bytes in document.package.parts.values() {
        let mut reader = NsReader::from_reader(bytes.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element) | Event::Empty(element)) => {
                    if let Ok(Some((_, value))) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    ) && let Ok(id) = value.parse::<i32>()
                    {
                        occupied.insert(id);
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut maps = BTreeMap::<String, BTreeMap<String, String>>::new();
    let mut endpoints = BTreeMap::<(String, String), [usize; 2]>::new();
    let mut pending = Vec::new();
    let mut next = 0_i32;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid fragment annotation XML: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace)
                    && let Some((family, endpoint)) = fragment_marker_family(local.as_ref())
                {
                    let (_, old) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    .ok_or_else(|| {
                        Error::Other(format!("document fragment {family} annotation has no ID"))
                    })?;
                    if endpoint < 2 {
                        endpoints
                            .entry((family.to_owned(), old.clone()))
                            .or_default()[usize::from(endpoint)] += 1;
                    }
                    let map = maps.entry(family.to_owned()).or_default();
                    if let std::collections::btree_map::Entry::Vacant(entry) = map.entry(old) {
                        while occupied.contains(&next) {
                            next = next.checked_add(1).ok_or_else(|| {
                                Error::Other(
                                    "document fragment annotation ID range is exhausted".to_owned(),
                                )
                            })?;
                        }
                        occupied.insert(next);
                        entry.insert(next.to_string());
                    }
                    let mut edits = Vec::new();
                    add_identity_attribute_edit(
                        xml,
                        before,
                        after,
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                        map,
                        &mut edits,
                    )?;
                    pending.extend(edits);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if let Some(((family, id), _)) = endpoints.iter().find(|(_, counts)| **counts != [1, 1]) {
        return Err(Error::Other(format!(
            "document fragment {family} range {id} has incomplete or ambiguous ownership"
        )));
    }
    let mut updated = xml.to_vec();
    for edit in pending.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

fn fragment_store_item_ids(xml: &[u8]) -> Result<BTreeSet<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut ids = BTreeSet::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid document fragment binding XML: {error}"))
        })? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace)
                    && local.as_ref() == b"dataBinding"
                    && let Some((_, value)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"storeItemID",
                        AttributeNamespace::Word,
                    )?
                {
                    ids.insert(crate::content_control::normalize_store_item_id(&value));
                }
            }
            Event::Eof => return Ok(ids),
            _ => {}
        }
        buffer.clear();
    }
}

pub(crate) fn fragment_note_references(xml: &[u8]) -> Result<Vec<(StoryKind, i32)>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut references = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid document fragment note XML: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                let kind = if namespace_is_word(&namespace) {
                    match local.as_ref() {
                        b"footnoteReference" => Some(StoryKind::Footnote),
                        b"endnoteReference" => Some(StoryKind::Endnote),
                        _ => None,
                    }
                } else {
                    None
                };
                if let Some(kind) = kind {
                    let (_, value) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    .ok_or_else(|| {
                        Error::Other(format!("document fragment {kind:?} reference has no ID"))
                    })?;
                    let id = value.parse::<i32>().map_err(|_| {
                        Error::Other(format!(
                            "document fragment {kind:?} reference ID {value} is invalid"
                        ))
                    })?;
                    if !references.contains(&(kind, id)) {
                        references.push((kind, id));
                    }
                }
            }
            Event::Eof => return Ok(references),
            _ => {}
        }
        buffer.clear();
    }
}

fn replace_fragment_store_id(
    xml: &[u8],
    old: &str,
    new: &str,
    element_name: &[u8],
    attribute_name: &[u8],
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "invalid document fragment custom XML store: {error}"
            ))
        })?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                let expected_namespace = if element_name == b"datastoreItem" {
                    namespace_matches(
                        &namespace,
                        "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
                    )
                } else {
                    namespace_is_word(&namespace)
                };
                if !expected_namespace || local.as_ref() != element_name {
                    continue;
                }
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!(
                            "invalid document fragment custom XML store attribute: {error}"
                        ))
                    })?;
                    let (attribute_namespace, local) =
                        reader.resolver().resolve_attribute(attribute.key);
                    let namespace_matches_expected = if element_name == b"datastoreItem" {
                        namespace_matches(
                            &attribute_namespace,
                            "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
                        )
                    } else {
                        namespace_is_word(&attribute_namespace)
                    };
                    if !namespace_matches_expected || local.as_ref() != attribute_name {
                        continue;
                    }
                    let value = quick_xml::escape::unescape(
                        std::str::from_utf8(&attribute.value).map_err(|error| {
                            Error::Other(format!(
                                "invalid document fragment custom XML store value: {error}"
                            ))
                        })?,
                    )
                    .map_err(|error| {
                        Error::Other(format!(
                            "invalid document fragment custom XML store value: {error}"
                        ))
                    })?;
                    if crate::content_control::normalize_store_item_id(&value)
                        != crate::content_control::normalize_store_item_id(old)
                    {
                        continue;
                    }
                    let (start, end) =
                        attribute_value_span(&xml[before..after], attribute.key.as_ref())
                            .ok_or_else(|| {
                                Error::Other(
                                    "document fragment custom XML store attribute span is missing"
                                        .to_owned(),
                                )
                            })?;
                    edits.push((before + start, before + end));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if edits.is_empty() {
        return Err(Error::Other(format!(
            "document fragment custom XML store {old} cannot be rewritten without corrupting its XML"
        )));
    }
    let mut result = xml.to_vec();
    for (start, end) in edits.into_iter().rev() {
        result.splice(start..end, new.as_bytes().iter().copied());
    }
    Ok(result)
}

pub(crate) fn relationship_ids_in_xml(xml: &[u8]) -> Result<Vec<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut ids = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid rich mail merge relationship XML: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!(
                            "invalid rich mail merge relationship attribute: {error}"
                        ))
                    })?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if namespace_matches(&namespace, R_NS)
                        && matches!(local.as_ref(), b"id" | b"embed" | b"link")
                    {
                        let raw = std::str::from_utf8(&attribute.value).map_err(|error| {
                            Error::Other(format!(
                                "invalid rich mail merge relationship id: {error}"
                            ))
                        })?;
                        let id = quick_xml::escape::unescape(raw)
                            .map_err(|error| {
                                Error::Other(format!(
                                    "invalid rich mail merge relationship id: {error}"
                                ))
                            })?
                            .into_owned();
                        if !ids.contains(&id) {
                            ids.push(id);
                        }
                    }
                }
            }
            Event::Eof => return Ok(ids),
            _ => {}
        }
        buffer.clear();
    }
}

pub(crate) fn patch_relationship_ids(
    xml: &[u8],
    replacements: &BTreeMap<String, String>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid rich mail merge relationship XML: {error}"))
        })?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!(
                            "invalid rich mail merge relationship attribute: {error}"
                        ))
                    })?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if !namespace_matches(&namespace, R_NS)
                        || !matches!(local.as_ref(), b"id" | b"embed" | b"link")
                    {
                        continue;
                    }
                    let raw = std::str::from_utf8(&attribute.value).map_err(|error| {
                        Error::Other(format!("invalid rich mail merge relationship id: {error}"))
                    })?;
                    let old = quick_xml::escape::unescape(raw).map_err(|error| {
                        Error::Other(format!("invalid rich mail merge relationship id: {error}"))
                    })?;
                    let Some(replacement) = replacements.get(old.as_ref()) else {
                        continue;
                    };
                    let Some((relative_start, relative_end)) =
                        attribute_value_span(&xml[before..after], attribute.key.as_ref())
                    else {
                        return Err(Error::Other(
                            "rich mail merge relationship attribute source was not found"
                                .to_owned(),
                        ));
                    };
                    edits.push(FieldSourceEdit {
                        start: before + relative_start,
                        end: before + relative_end,
                        replacement: replacement.as_bytes().to_vec(),
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut updated = xml.to_vec();
    edits.sort_by_key(|edit| edit.start);
    for edit in edits.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

pub(crate) fn copy_fragment_part(
    source: &OpcPackage,
    destination: &mut Document,
    source_part: &str,
    part_map: &BTreeMap<String, String>,
    copied: &mut HashSet<String>,
) -> Result<String> {
    let destination_part = part_map
        .get(source_part)
        .or_else(|| {
            part_map.iter().find_map(|(source, destination)| {
                source
                    .eq_ignore_ascii_case(source_part)
                    .then_some(destination)
            })
        })
        .cloned()
        .ok_or_else(|| {
            Error::Other(format!(
                "rich mail merge fragment part {source_part} was not preallocated"
            ))
        })?;
    if !copied.insert(destination_part.to_ascii_lowercase()) {
        return Ok(destination_part);
    }
    let bytes = source.get_part(source_part).ok_or_else(|| {
        Error::Other(format!(
            "rich mail merge fragment part {source_part} is missing"
        ))
    })?;
    destination
        .package
        .set_part(&destination_part, bytes.to_vec());
    destination
        .identifiers
        .preserve_fragment_part(&destination_part);
    if let Some(content_type) = source.content_types.content_type_for(source_part) {
        destination
            .package
            .content_types
            .add_override(&destination_part, content_type);
        destination
            .identifiers
            .register_content_type_override(&destination_part);
    }
    if let Some(relationships) = source.get_part_rels(source_part) {
        let mut relationships = relationships.items.clone();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        for relationship in relationships {
            if !crate::document::relationship_is_internal(&relationship) {
                destination
                    .package
                    .get_or_create_part_rels(&destination_part)
                    .items
                    .push(relationship);
                continue;
            }
            let child_source = OpcPackage::resolve_rel_target(source_part, &relationship.target);
            let child_destination =
                copy_fragment_part(source, destination, &child_source, part_map, copied)?;
            let target = relative_fragment_target(&destination_part, &child_destination);
            let destination_id = destination
                .identifiers
                .reserve_requested_relationship_id_checked(
                    &destination_part,
                    &relationship.id,
                )
                .map_err(|error| {
                    Error::Other(format!(
                        "rich mail merge relationship allocation failed for {destination_part}: {error}"
                    ))
                })?;
            destination
                .package
                .get_or_create_part_rels(&destination_part)
                .add_with_id(&destination_id, &relationship.rel_type, &target);
        }
    }
    Ok(destination_part)
}

fn discover_fragment_part_closure(
    source: &OpcPackage,
    source_part: &str,
    discovered: &mut HashSet<String>,
) -> Result<()> {
    if source
        .content_types
        .content_type_for(source_part)
        .is_some_and(|content_type| {
            content_type.starts_with("application/vnd.openxmlformats-package.digital-signature")
        })
    {
        return Err(Error::Other(format!(
            "document fragment integrity-bound signature part {source_part} cannot remain valid after remapping"
        )));
    }
    if !discovered.insert(source_part.to_owned()) {
        return Ok(());
    }
    source.get_part(source_part).ok_or_else(|| {
        Error::Other(format!(
            "rich mail merge fragment part {source_part} is missing"
        ))
    })?;
    if let Some(relationships) = source.get_part_rels(source_part) {
        for relationship in &relationships.items {
            if !crate::document::relationship_is_internal(relationship) {
                continue;
            }
            let child_source = OpcPackage::resolve_rel_target(source_part, &relationship.target);
            discover_fragment_part_closure(source, &child_source, discovered)?;
        }
    }
    Ok(())
}

fn relative_fragment_target(source_part: &str, target_part: &str) -> String {
    let directory = source_part
        .rfind('/')
        .map_or("/", |position| &source_part[..=position]);
    target_part
        .strip_prefix(directory)
        .filter(|target| !target.contains('/'))
        .unwrap_or(target_part)
        .to_owned()
}

fn resolve_region_records<'a>(
    scopes: &[&'a MailMergeRecord],
    data: &'a MailMergeData,
    name: &str,
) -> Result<&'a [MailMergeRecord]> {
    if let Some(records) = scopes.last().and_then(|record| record.regions.get(name)) {
        return Ok(records);
    }
    data.sources
        .get(name)
        .map(Vec::as_slice)
        .ok_or_else(|| Error::Other(format!("rich mail merge source {name} is missing")))
}

fn resolve_rich_value<'a>(
    scopes: &[&'a MailMergeRecord],
    name: &str,
) -> Option<&'a MailMergeValue> {
    scopes
        .iter()
        .rev()
        .find_map(|record| record.values.get(name))
}

fn resolve_mergefield_text(
    instruction: &FieldInstruction,
    value: Option<&str>,
    missing_as_empty: bool,
) -> std::result::Result<String, String> {
    if let Some(name) = unsupported_switch(instruction) {
        return Err(format!("field MERGEFIELD uses unsupported switch \\{name}"));
    }
    validate_instruction_shape(instruction)?;
    let name = text_argument(instruction, 0)
        .ok_or_else(|| "MERGEFIELD requires a field name".to_owned())?;
    let Some(value) = value else {
        return if missing_as_empty {
            Ok(String::new())
        } else {
            Err(format!("MERGEFIELD input {name} was not supplied"))
        };
    };
    if value.is_empty() {
        return Ok(String::new());
    }
    let mut result = String::new();
    if let Some(prefix) = switch_text(instruction, "b") {
        result.push_str(prefix);
    }
    result.push_str(value);
    if let Some(suffix) = switch_text(instruction, "f") {
        result.push_str(suffix);
    }
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_body_item(
    document: &mut Document,
    item: &mut BodyContent,
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    match item {
        BodyContent::Paragraph(paragraph) => replace_rich_paragraph(
            document,
            paragraph,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        ),
        BodyContent::Table(table) => replace_rich_table(
            document,
            table,
            data,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        ),
        BodyContent::ContentControl(control) => replace_rich_control(
            document,
            control,
            data,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        ),
        BodyContent::RawXml(_) => Ok(()),
    }
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_paragraph(
    document: &mut Document,
    paragraph: &mut CT_P,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    for run in &mut paragraph.runs {
        replace_rich_run(
            document,
            run,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            formatter,
        )?;
    }
    for (_, _, _, control) in &mut paragraph.content_controls {
        replace_rich_control(
            document,
            control,
            &MailMergeData::default(),
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_run(
    document: &mut Document,
    run: &mut CT_R,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    for content in &mut run.content {
        let RunContent::Field(field) = content else {
            continue;
        };
        let Some(field_name) = merge_field_name(field) else {
            continue;
        };
        if field_name.starts_with("TableStart:") || field_name.starts_with("TableEnd:") {
            return Err(Error::Other(format!(
                "rich mail merge region field {field_name} must own its whole block"
            )));
        }
        *output_sequence_number = output_sequence_number.checked_add(1).ok_or_else(|| {
            Error::Other("rich mail merge output sequence exceeds u32".to_owned())
        })?;
        let merge_value = resolve_rich_value(scopes, &field_name);
        if matches!(merge_value, Some(MailMergeValue::Text(_)) | None) {
            let value = match merge_value {
                Some(MailMergeValue::Text(value)) => Some(value.as_str()),
                None => None,
                _ => unreachable!(),
            };
            let instruction = field.effective_instruction();
            let resolved = resolve_mergefield_text(&instruction, value, true)
                .and_then(|value| apply_formats(&instruction, &value, None))
                .map_err(Error::Other)?;
            let mut formatted = MailMergeFormattedText {
                text: resolved,
                run_properties: run.properties.clone(),
            };
            if let Some(callback) = formatter.as_deref_mut() {
                callback(
                    &MailMergeFormatContext {
                        source_name,
                        region_path,
                        field_name: &field_name,
                        record_number,
                        output_sequence_number: *output_sequence_number,
                    },
                    &mut formatted,
                )?;
            }
            if let Some(character) = formatted
                .text
                .chars()
                .find(|character| !valid_xml_character(*character))
            {
                return Err(Error::Other(format!(
                    "rich mail merge field {field_name} contains invalid XML character U+{:04X}",
                    u32::from(character)
                )));
            }
            run.properties = formatted.run_properties;
            *content = RunContent::Text(CT_Text::new(&formatted.text));
            continue;
        }
        match merge_value {
            Some(MailMergeValue::Image(image)) => {
                if image.width.to_emu() <= 0 || image.height.to_emu() <= 0 {
                    return Err(Error::Other(format!(
                        "rich mail merge image {field_name} has non-positive dimensions"
                    )));
                }
                let rel_id = document.embed_image(&image.data, &image.filename);
                let mut inline =
                    CT_Inline::new(&rel_id, image.width.to_emu(), image.height.to_emu());
                inline.doc_pr_id = document.identifiers.reserve_drawing_id()?;
                *content = RunContent::Drawing(CT_Drawing::inline(inline));
            }
            Some(MailMergeValue::Fragment(_)) => {
                return Err(Error::Other(format!(
                    "rich mail merge fragment field {field_name} must own its whole paragraph"
                )));
            }
            Some(MailMergeValue::Text(_)) | None => unreachable!(),
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_table(
    document: &mut Document,
    table: &mut CT_Tbl,
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    let template_rows = std::mem::take(&mut table.rows);
    for (index, row) in template_rows.iter().enumerate() {
        if rich_row_region_marker(row).is_some()
            && (table.extra_xml.iter().any(|(at, _)| *at == index)
                || table.content_controls.iter().any(|(at, _, _)| *at == index))
        {
            return Err(Error::Other(
                "rich mail merge row marker cannot share a table boundary with preserved content"
                    .to_owned(),
            ));
        }
    }
    let expanded = expand_rich_rows(
        document,
        &template_rows,
        0,
        data,
        scopes,
        source_name,
        region_path,
        record_number,
        output_sequence_number,
        identity_state,
        formatter,
    )?;
    let old_extra = std::mem::take(&mut table.extra_xml);
    let old_controls = std::mem::take(&mut table.content_controls);
    for (output_index, expanded_row) in expanded.iter().enumerate() {
        table.extra_xml.extend(
            old_extra
                .iter()
                .filter(|(at, _)| *at == expanded_row.source_index)
                .map(|(_, raw)| (output_index, raw.clone())),
        );
        table.content_controls.extend(
            old_controls
                .iter()
                .filter(|(at, _, _)| *at == expanded_row.source_index)
                .map(|(_, raw_before, control)| (output_index, *raw_before, control.clone())),
        );
    }
    let output_end = expanded.len();
    table.extra_xml.extend(
        old_extra
            .iter()
            .filter(|(at, _)| *at == template_rows.len())
            .map(|(_, raw)| (output_end, raw.clone())),
    );
    table.content_controls.extend(
        old_controls
            .iter()
            .filter(|(at, _, _)| *at == template_rows.len())
            .map(|(_, raw_before, control)| (output_end, *raw_before, control.clone())),
    );
    table.rows = expanded.into_iter().map(|item| item.row).collect();
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn expand_rich_rows(
    document: &mut Document,
    template_rows: &[CT_Row],
    source_offset: usize,
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<Vec<RichExpandedRow>> {
    let mut output = Vec::new();
    let mut index = 0;
    while index < template_rows.len() {
        if let Some(RichRegionMarker::Start(name)) = rich_row_region_marker(&template_rows[index]) {
            if region_path.len() >= RICH_MERGE_REGION_DEPTH_LIMIT {
                return Err(Error::Other(format!(
                    "rich mail merge row region depth exceeds {RICH_MERGE_REGION_DEPTH_LIMIT}"
                )));
            }
            let end = find_row_region_end(template_rows, index + 1, &name)?;
            let records = resolve_region_records(scopes, data, &name)?;
            for (child_index, child) in records.iter().enumerate() {
                let mut child_scopes = scopes.to_vec();
                child_scopes.push(child);
                let mut child_path = region_path.to_vec();
                child_path.push(name.clone());
                let mut child_rows = expand_rich_rows(
                    document,
                    &template_rows[index + 1..end],
                    source_offset + index + 1,
                    data,
                    &child_scopes,
                    &name,
                    &child_path,
                    one_based_record_number(child_index)?,
                    output_sequence_number,
                    identity_state,
                    formatter,
                )?;
                remap_rich_rows(document, &mut child_rows, identity_state)?;
                output.extend(child_rows);
            }
            index = end + 1;
            continue;
        }
        if let Some(RichRegionMarker::End(name)) = rich_row_region_marker(&template_rows[index]) {
            return Err(Error::Other(format!(
                "unexpected rich mail merge row marker TableEnd:{name}"
            )));
        }
        let mut row = template_rows[index].clone();
        replace_rich_row(
            document,
            &mut row,
            data,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        )?;
        output.push(RichExpandedRow {
            source_index: source_offset + index,
            row,
        });
        index += 1;
    }
    Ok(output)
}

fn rich_row_region_marker(row: &CT_Row) -> Option<RichRegionMarker> {
    if row
        .extra_xml
        .iter()
        .any(|(position, raw)| !CT_Row::raw_is_root_attributes(*position, raw))
        || !row.content_controls.is_empty()
        || row.cells.len() != 1
    {
        return None;
    }
    let cell = &row.cells[0];
    if !cell.extra_xml.is_empty() || cell.content.len() != 1 {
        return None;
    }
    let CellContent::Paragraph(paragraph) = &cell.content[0] else {
        return None;
    };
    rich_region_marker(paragraph)
}

fn find_row_region_end(rows: &[CT_Row], start: usize, name: &str) -> Result<usize> {
    let mut stack = vec![name.to_owned()];
    for (index, row) in rows.iter().enumerate().skip(start) {
        match rich_row_region_marker(row) {
            Some(RichRegionMarker::Start(nested)) => stack.push(nested),
            Some(RichRegionMarker::End(end)) => {
                if stack.last() != Some(&end) {
                    return Err(Error::Other(format!(
                        "crossed rich mail merge row marker TableEnd:{end}"
                    )));
                }
                stack.pop();
                if stack.is_empty() {
                    return Ok(index);
                }
            }
            None => {}
        }
    }
    Err(Error::Other(format!(
        "rich mail merge row region TableStart:{name} has no matching end"
    )))
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_row(
    document: &mut Document,
    row: &mut CT_Row,
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    for cell in &mut row.cells {
        replace_rich_cell(
            document,
            cell,
            data,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        )?;
    }
    for (_, _, control) in &mut row.content_controls {
        replace_rich_control(
            document,
            control,
            data,
            scopes,
            source_name,
            region_path,
            record_number,
            output_sequence_number,
            identity_state,
            formatter,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_cell(
    document: &mut Document,
    cell: &mut CT_Tc,
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => replace_rich_paragraph(
                document,
                paragraph,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            CellContent::Table(table) => replace_rich_table(
                document,
                table,
                data,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            CellContent::ContentControl(control) => replace_rich_control(
                document,
                control,
                data,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn replace_rich_control(
    document: &mut Document,
    control: &mut CT_Sdt,
    data: &MailMergeData,
    scopes: &[&MailMergeRecord],
    source_name: &str,
    region_path: &[String],
    record_number: u32,
    output_sequence_number: &mut u32,
    identity_state: &mut BodyIdentityState,
    formatter: &mut RichFormatter<'_>,
) -> Result<()> {
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => replace_rich_paragraph(
                document,
                paragraph,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            SdtContent::Table(table) => replace_rich_table(
                document,
                table,
                data,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            SdtContent::Row(row) => replace_rich_row(
                document,
                row,
                data,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            SdtContent::Cell(cell) => replace_rich_cell(
                document,
                cell,
                data,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            SdtContent::Run(run) => replace_rich_run(
                document,
                run,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                formatter,
            )?,
            SdtContent::ContentControl(nested) => replace_rich_control(
                document,
                nested,
                data,
                scopes,
                source_name,
                region_path,
                record_number,
                output_sequence_number,
                identity_state,
                formatter,
            )?,
            SdtContent::RawXml(_) => {}
        }
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct DynamicTocSpan {
    simple_field: Option<Field>,
    instruction: String,
    field_start: usize,
    field_end: usize,
    begin_paragraph: usize,
    end_paragraph: usize,
    begin_run_start: usize,
    instruction_paragraph_start: usize,
    result_start: usize,
    result_end: usize,
    result_start_position: TocRunPosition,
    result_end_position: TocRunPosition,
    end_run_end: usize,
    start_paragraph_name: String,
    start_paragraph_namespaces: BTreeMap<String, String>,
    separator_wrapper_names: Vec<String>,
    instruction_runs: Vec<DynamicInstructionRun>,
    end_paragraph_start: usize,
    end_paragraph_content_start: usize,
    end_wrapper_prefixes: Vec<(usize, usize)>,
}

#[derive(Debug)]
struct DynamicFieldScan {
    instruction: String,
    field_start: usize,
    begin_paragraph: usize,
    begin_paragraph_start: usize,
    begin_run_start: usize,
    separator_paragraph: Option<usize>,
    separator_run_start: Option<usize>,
    result_start: Option<usize>,
    result_start_position: Option<TocRunPosition>,
    start_paragraph_name: Option<String>,
    start_paragraph_namespaces: BTreeMap<String, String>,
    separator_wrapper_names: Vec<String>,
    instruction_runs: Vec<DynamicInstructionRun>,
}

#[derive(Debug, Clone)]
struct DynamicInstructionRun {
    start: usize,
    end: usize,
    inherited_namespaces: BTreeMap<String, String>,
}

#[derive(Debug)]
struct DynamicXmlElement {
    local_name: Vec<u8>,
    qualified_name: String,
    typed_block_owner: Option<TypedBlockOwner>,
    is_word: bool,
    is_typed_paragraph: bool,
    is_typed_inline_owner: bool,
    revision_depth: usize,
    sdt_content_seen: bool,
    namespace_bindings: BTreeMap<String, String>,
    inherited_namespaces: BTreeMap<String, String>,
    run_position: Option<TocRunPosition>,
    hyperlink_plan: Option<DynamicHyperlinkPlan>,
    start: usize,
    start_tag_end: usize,
    paragraph: Option<usize>,
}

#[derive(Debug)]
struct DynamicHyperlinkPlan {
    revision_orders: Vec<TocRawOrder>,
    next_revision: usize,
    preserved_raw: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TypedBlockOwner {
    Document,
    Body,
    Table,
    Row,
    Cell,
    ContentControl(BlockControlOwner),
    Content(BlockControlOwner),
    Paragraph,
}

const MAX_DYNAMIC_REVISION_NESTING_DEPTH: usize = 32;

fn dynamic_namespace_bindings(
    element: &BytesStart<'_>,
    elements: &[DynamicXmlElement],
) -> Result<(BTreeMap<String, String>, BTreeMap<String, String>)> {
    let inherited = elements
        .last()
        .map_or_else(BTreeMap::new, |parent| parent.namespace_bindings.clone());
    let mut bindings = inherited.clone();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Other(format!("invalid XML namespace attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            Some(String::new())
        } else {
            key.strip_prefix(b"xmlns:")
                .map(|prefix| String::from_utf8_lossy(prefix).into_owned())
        };
        if let Some(prefix) = prefix {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .map_err(|error| Error::Other(format!("invalid XML namespace attribute: {error}")))?
                .into_owned();
            if value.is_empty() {
                bindings.remove(&prefix);
            } else {
                bindings.insert(prefix, value);
            }
        }
    }
    Ok((bindings, inherited))
}

#[derive(Clone, Copy)]
enum DynamicContentControlOwner {
    Block(BlockControlOwner),
    Inline,
}

fn dynamic_element_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut reader = quick_xml::Reader::from_reader(&xml[start..]);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid table of contents XML element: {error}"))
        })? {
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => {
                return Ok(start + reader.buffer_position() as usize);
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Other("table of contents XML element has an unmatched end".to_owned())
                })?;
                if depth == 0 {
                    return Ok(start + reader.buffer_position() as usize);
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "table of contents XML element is unclosed".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

pub(crate) fn xml_fragment_with_namespaces(
    raw: &[u8],
    bindings: &BTreeMap<String, String>,
    description: &str,
) -> Result<Vec<u8>> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (insertion, local_namespaces) = match reader
        .read_event_into(&mut buffer)
        .map_err(|error| Error::Other(format!("invalid {description}: {error}")))?
    {
        Event::Start(start) | Event::Empty(start) => {
            let mut local_namespaces = HashSet::new();
            for attribute in start.attributes() {
                let attribute = attribute
                    .map_err(|error| Error::Other(format!("invalid {description}: {error}")))?;
                let key = attribute.key.as_ref();
                if key == b"xmlns" {
                    local_namespaces.insert(String::new());
                } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                    local_namespaces.insert(String::from_utf8_lossy(prefix).into_owned());
                }
            }
            let tag_end = reader.buffer_position() as usize;
            let insertion = if tag_end >= 2 && raw[tag_end - 2] == b'/' {
                tag_end - 2
            } else {
                tag_end - 1
            };
            (insertion, local_namespaces)
        }
        _ => return Err(Error::Other(format!("{description} has no start tag"))),
    };
    let mut output = Vec::with_capacity(raw.len() + bindings.len() * 32);
    output.extend_from_slice(&raw[..insertion]);
    for (prefix, namespace) in bindings {
        if prefix == "xml" || local_namespaces.contains(prefix) {
            continue;
        }
        if prefix.is_empty() {
            output.extend_from_slice(b" xmlns=\"");
        } else {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
        }
        output.extend_from_slice(xml_escape_attribute(namespace).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(&raw[insertion..]);
    Ok(output)
}

fn dynamic_content_control_is_typed(
    xml: &[u8],
    start: usize,
    bindings: &BTreeMap<String, String>,
    owner: DynamicContentControlOwner,
) -> Result<bool> {
    let end = dynamic_element_end(xml, start)?;
    let control = xml_fragment_with_namespaces(
        &xml[start..end],
        bindings,
        "table of contents content control",
    )?;
    let mut candidate = Vec::with_capacity(control.len() + 256);
    candidate.extend_from_slice(format!(r#"<w:document xmlns:w="{W_NS}"><w:body>"#).as_bytes());
    let (prefix, suffix): (&[u8], &[u8]) = match owner {
        DynamicContentControlOwner::Block(BlockControlOwner::Body) => (b"", b""),
        DynamicContentControlOwner::Block(BlockControlOwner::Table) => {
            (b"<w:tbl>", b"<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>")
        }
        DynamicContentControlOwner::Block(BlockControlOwner::Row) => {
            (b"<w:tbl><w:tr>", b"<w:tc><w:p/></w:tc></w:tr></w:tbl>")
        }
        DynamicContentControlOwner::Block(BlockControlOwner::Cell) => {
            (b"<w:tbl><w:tr><w:tc>", b"<w:p/></w:tc></w:tr></w:tbl>")
        }
        DynamicContentControlOwner::Inline => (b"<w:p>", b"</w:p>"),
    };
    candidate.extend_from_slice(prefix);
    candidate.extend_from_slice(&control);
    candidate.extend_from_slice(suffix);
    candidate.extend_from_slice(b"</w:body></w:document>");
    let Ok(document) = CT_Document::from_xml(&candidate) else {
        return Ok(false);
    };
    let accepted = match owner {
        DynamicContentControlOwner::Block(BlockControlOwner::Body) => matches!(
            document.body.content.first(),
            Some(BodyContent::ContentControl(_))
        ),
        DynamicContentControlOwner::Block(BlockControlOwner::Table) => document
            .body
            .tables()
            .next()
            .is_some_and(|table| table.content_controls.len() == 1),
        DynamicContentControlOwner::Block(BlockControlOwner::Row) => document
            .body
            .tables()
            .next()
            .and_then(|table| table.rows.first())
            .is_some_and(|row| row.content_controls.len() == 1),
        DynamicContentControlOwner::Block(BlockControlOwner::Cell) => document
            .body
            .tables()
            .next()
            .and_then(|table| table.rows.first())
            .and_then(|row| row.cells.first())
            .is_some_and(|cell| {
                matches!(cell.content.first(), Some(CellContent::ContentControl(_)))
            }),
        DynamicContentControlOwner::Inline => document
            .body
            .paragraphs()
            .next()
            .is_some_and(|paragraph| paragraph.content_controls.len() == 1),
    };
    Ok(accepted)
}

fn validate_dynamic_content_control(
    xml: &[u8],
    start: usize,
    word: bool,
    local: &[u8],
    bindings: &BTreeMap<String, String>,
    typed_block_owner: &mut Option<TypedBlockOwner>,
    is_typed_inline_owner: &mut bool,
) -> Result<()> {
    if !word || local != b"sdt" {
        return Ok(());
    }
    let owner = match *typed_block_owner {
        Some(TypedBlockOwner::ContentControl(owner)) => {
            Some(DynamicContentControlOwner::Block(owner))
        }
        _ if *is_typed_inline_owner => Some(DynamicContentControlOwner::Inline),
        _ => None,
    };
    if let Some(owner) = owner
        && !dynamic_content_control_is_typed(xml, start, bindings, owner)?
    {
        *typed_block_owner = None;
        *is_typed_inline_owner = false;
    }
    Ok(())
}

fn is_dynamic_revision_element(local: &[u8]) -> bool {
    matches!(
        local,
        b"ins"
            | b"del"
            | b"moveFrom"
            | b"moveTo"
            | b"rPrChange"
            | b"pPrChange"
            | b"tblPrChange"
            | b"sectPrChange"
    )
}

fn invalidate_overdeep_revision_owner(
    word: bool,
    local: &[u8],
    elements: &mut [DynamicXmlElement],
) -> Option<usize> {
    if !word
        || !is_dynamic_revision_element(local)
        || elements
            .iter()
            .filter(|element| element.is_word && is_dynamic_revision_element(&element.local_name))
            .count()
            < MAX_DYNAMIC_REVISION_NESTING_DEPTH
    {
        return None;
    }
    let owner = elements.iter().position(|element| {
        element.is_typed_inline_owner && matches!(element.local_name.as_slice(), b"ins" | b"moveTo")
    })?;
    let start = elements[owner].start;
    for element in &mut elements[owner..] {
        element.is_typed_inline_owner = false;
    }
    Some(start)
}

fn mark_typed_sdt_content(
    elements: &mut [DynamicXmlElement],
    local: &[u8],
    typed_block_owner: Option<TypedBlockOwner>,
    is_typed_inline_owner: bool,
) {
    if local == b"sdtContent"
        && (matches!(typed_block_owner, Some(TypedBlockOwner::Content(_))) || is_typed_inline_owner)
        && let Some(parent) = elements.last_mut()
    {
        parent.sdt_content_seen = true;
    }
}

fn scan_dynamic_table_spans(
    xml: &[u8],
    generated: DynamicOwnerPolicy,
) -> Result<Vec<DynamicTocSpan>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut fields = Vec::<DynamicFieldScan>::new();
    let mut spans = Vec::<DynamicTocSpan>::new();
    let mut paragraph_count = 0usize;
    let mut paragraph_run_boundaries = Vec::<usize>::new();
    let mut paragraph_nested_run_orders = Vec::<HashMap<(usize, TocRawOrder), usize>>::new();
    let mut paragraph_raw_before = Vec::<usize>::new();
    let mut instruction_depth = None;

    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    fields.retain(|field| field.field_start < start);
                    spans.retain(|span| span.field_start < start);
                    instruction_depth = None;
                }
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    let index = paragraph_count;
                    paragraph_count += 1;
                    paragraph_run_boundaries.push(0);
                    paragraph_nested_run_orders.push(HashMap::new());
                    paragraph_raw_before.push(0);
                    Some(index)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let mut is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let revision_depth = if is_typed_inline_owner {
                    elements.last().map_or(0, |parent| parent.revision_depth)
                        + usize::from(matches!(local.as_slice(), b"ins" | b"moveTo"))
                } else {
                    0
                };
                let modeled_simple_field = if word && local == b"fldSimple" {
                    resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        !Field::new(&instruction, "").instruction.name.is_empty()
                    })
                } else {
                    false
                };
                let run_position = dynamic_toc_run_position(
                    word,
                    &local,
                    is_typed_inline_owner,
                    modeled_simple_field,
                    &mut elements,
                    paragraph,
                    &mut paragraph_run_boundaries,
                    &mut paragraph_nested_run_orders,
                    &mut paragraph_raw_before,
                );
                let direct_paragraph_child = direct_typed_paragraph_parent(&elements, paragraph);
                let hyperlink_plan = if word && local == b"hyperlink" && direct_paragraph_child {
                    Some(dynamic_hyperlink_plan(
                        xml,
                        before,
                        &namespace_bindings,
                        paragraph
                            .and_then(|index| paragraph_raw_before.get(index).copied())
                            .ok_or_else(|| {
                                Error::Other(
                                    "table of contents hyperlink has no paragraph position"
                                        .to_owned(),
                                )
                            })?,
                    )?)
                } else {
                    None
                };
                advance_direct_paragraph_raw_child(
                    word,
                    &local,
                    is_typed_inline_owner,
                    modeled_simple_field,
                    direct_paragraph_child,
                    paragraph,
                    hyperlink_plan.as_ref(),
                    &mut paragraph_raw_before,
                );
                if word && local == b"fldChar" && direct_word_run_parent(&elements, paragraph) {
                    update_dynamic_field_stack(
                        &element,
                        reader.resolver(),
                        before,
                        paragraph,
                        &elements,
                        &mut fields,
                        &mut spans,
                        generated,
                    )?;
                }
                if word
                    && local == b"instrText"
                    && !fields.is_empty()
                    && direct_word_run_parent(&elements, paragraph)
                {
                    instruction_depth = Some(elements.len());
                }
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner,
                    revision_depth,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position,
                    hyperlink_plan,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    fields.retain(|field| field.field_start < start);
                    spans.retain(|span| span.field_start < start);
                    instruction_depth = None;
                }
                let typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let paragraph = if typed_block_owner == Some(TypedBlockOwner::Paragraph) {
                    let index = paragraph_count;
                    paragraph_count += 1;
                    paragraph_run_boundaries.push(0);
                    paragraph_nested_run_orders.push(HashMap::new());
                    paragraph_raw_before.push(0);
                    Some(index)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let modeled_simple_field = word
                    && local == b"fldSimple"
                    && resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        !Field::new(&instruction, "").instruction.name.is_empty()
                    });
                let run_position = dynamic_toc_run_position(
                    word,
                    &local,
                    is_typed_inline_owner,
                    modeled_simple_field,
                    &mut elements,
                    paragraph,
                    &mut paragraph_run_boundaries,
                    &mut paragraph_nested_run_orders,
                    &mut paragraph_raw_before,
                );
                let direct_paragraph_child = direct_typed_paragraph_parent(&elements, paragraph);
                advance_direct_paragraph_raw_child(
                    word,
                    &local,
                    is_typed_inline_owner,
                    modeled_simple_field,
                    direct_paragraph_child,
                    paragraph,
                    None,
                    &mut paragraph_raw_before,
                );
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                if generated != DynamicOwnerPolicy::Toc && modeled_simple_field {
                    let closed = DynamicXmlElement {
                        local_name: local.clone(),
                        qualified_name: String::from_utf8_lossy(element.name().as_ref())
                            .into_owned(),
                        typed_block_owner,
                        is_word: word,
                        is_typed_paragraph: false,
                        is_typed_inline_owner,
                        revision_depth: 0,
                        sdt_content_seen: false,
                        namespace_bindings,
                        inherited_namespaces,
                        run_position,
                        hyperlink_plan: None,
                        start: before,
                        start_tag_end: after,
                        paragraph,
                    };
                    if let Some(span) =
                        generated_simple_span(xml, &closed, &elements, after, after, generated)?
                    {
                        spans.push(span);
                    }
                }
                if word
                    && matches_local_name(element.name().as_ref(), b"fldChar")
                    && direct_word_run_parent(&elements, paragraph)
                {
                    update_dynamic_field_stack(
                        &element,
                        reader.resolver(),
                        before,
                        paragraph,
                        &elements,
                        &mut fields,
                        &mut spans,
                        generated,
                    )?;
                }
            }
            Event::Text(text) if instruction_depth.is_some() => {
                if let Some(field) = fields.last_mut()
                    && field.separator_paragraph.is_none()
                {
                    let decoded = text.decode().map_err(|error| {
                        Error::Other(format!("invalid table of contents instruction: {error}"))
                    })?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::Other(format!("invalid table of contents instruction: {error}"))
                    })?;
                    field.instruction.push_str(&unescaped);
                }
            }
            Event::GeneralRef(_)
                if generated == DynamicOwnerPolicy::Bibliography && instruction_depth.is_some() =>
            {
                if let Some(field) = fields.last_mut()
                    && field.separator_paragraph.is_none()
                {
                    let reference = std::str::from_utf8(&xml[before..after]).map_err(|error| {
                        Error::Other(format!(
                            "invalid bibliography instruction reference: {error}"
                        ))
                    })?;
                    let decoded = quick_xml::escape::unescape(reference).map_err(|error| {
                        Error::Other(format!(
                            "invalid bibliography instruction reference: {error}"
                        ))
                    })?;
                    field.instruction.push_str(&decoded);
                }
            }
            Event::CData(text) if instruction_depth.is_some() => {
                if let Some(field) = fields.last_mut()
                    && field.separator_paragraph.is_none()
                {
                    field.instruction.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid table of contents instruction: {error}"))
                    })?);
                }
            }
            Event::Comment(_) | Event::PI(_) => {
                if let Some(paragraph) = elements.last().and_then(|element| {
                    element
                        .is_typed_paragraph
                        .then_some(element.paragraph)
                        .flatten()
                }) && let Some(raw_before) = paragraph_raw_before.get_mut(paragraph)
                {
                    *raw_before += 1;
                }
            }
            Event::End(element) => {
                let Some(closed) = elements.pop() else {
                    return Err(Error::Other(
                        "table of contents XML has an unmatched end element".to_owned(),
                    ));
                };
                if generated != DynamicOwnerPolicy::Toc
                    && let Some(span) =
                        generated_simple_span(xml, &closed, &elements, before, after, generated)?
                {
                    spans.push(span);
                }
                if closed
                    .hyperlink_plan
                    .as_ref()
                    .is_some_and(|plan| plan.preserved_raw)
                    && let Some(paragraph) = closed.paragraph
                    && let Some(raw_before) = paragraph_raw_before.get_mut(paragraph)
                {
                    *raw_before += 1;
                }
                if word && matches_local_name(element.name().as_ref(), b"instrText") {
                    instruction_depth = None;
                }
                if word && matches_local_name(element.name().as_ref(), b"r") {
                    for field in &mut fields {
                        if closed.is_typed_inline_owner && field.result_start.is_none() {
                            field.instruction_runs.push(DynamicInstructionRun {
                                start: closed.start,
                                end: after,
                                inherited_namespaces: closed.inherited_namespaces.clone(),
                            });
                        }
                        if field.separator_run_start == Some(closed.start)
                            && field.result_start.is_none()
                        {
                            field.result_start = Some(after);
                        }
                    }
                    if let Some(span) = spans
                        .iter_mut()
                        .rev()
                        .find(|span| span.result_end == closed.start && span.end_run_end == 0)
                    {
                        span.end_run_end = after;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !elements.is_empty() {
        return Err(Error::Other(
            "table of contents XML has an unclosed element".to_owned(),
        ));
    }
    if fields.iter().any(|field| {
        generated_table_opcode(
            &Field::new(&field.instruction, "").instruction.name,
            generated,
        )
    }) {
        return Err(Error::Other(
            "table of contents field is missing its end marker".to_owned(),
        ));
    }
    spans.sort_by_key(|span| span.field_start);
    if spans
        .windows(2)
        .any(|pair| pair[1].field_start < pair[0].field_end)
    {
        return Err(Error::Other(
            "nested table of contents fields have ambiguous ownership".to_owned(),
        ));
    }
    let paragraphs = toc_paragraph_insertions(xml)?;
    for span in &mut spans {
        span.end_paragraph_content_start = paragraphs
            .get(span.end_paragraph)
            .map(|paragraph| paragraph.content_start)
            .ok_or_else(|| {
                Error::Other(
                    "table of contents end paragraph was not found in package XML".to_owned(),
                )
            })?;
    }
    Ok(spans)
}

fn collect_simple_toc_diagnostics(xml: &[u8]) -> Result<Vec<(usize, String)>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut simple_toc_starts = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    simple_toc_starts.retain(|field_start| *field_start < start);
                }
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    Some(0)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let mut is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let revision_depth = if is_typed_inline_owner {
                    elements.last().map_or(0, |parent| parent.revision_depth)
                        + usize::from(matches!(local.as_slice(), b"ins" | b"moveTo"))
                } else {
                    0
                };
                if word
                    && local == b"fldSimple"
                    && paragraph.is_some()
                    && accepted_simple_toc_parent(&elements, paragraph)
                    && resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        Field::new(&instruction, "").instruction.name == "TOC"
                    })
                {
                    simple_toc_starts.push(before);
                }
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner,
                    revision_depth,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    simple_toc_starts.retain(|field_start| *field_start < start);
                }
                let paragraph = elements.last().and_then(|element| element.paragraph);
                let typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                if word
                    && matches_local_name(element.name().as_ref(), b"fldSimple")
                    && elements
                        .last()
                        .and_then(|element| element.paragraph)
                        .is_some()
                    && accepted_simple_toc_parent(
                        &elements,
                        elements.last().and_then(|element| element.paragraph),
                    )
                    && resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        Field::new(&instruction, "").instruction.name == "TOC"
                    })
                {
                    simple_toc_starts.push(before);
                }
            }
            Event::End(_) => {
                elements.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(simple_toc_starts
        .into_iter()
        .map(|offset| {
            (
                offset,
                "simple table of contents fields are not rebuilt, stored display retained"
                    .to_owned(),
            )
        })
        .collect())
}

fn classify_typed_block_owner(
    word: bool,
    local: &[u8],
    elements: &[DynamicXmlElement],
) -> Option<TypedBlockOwner> {
    if !word {
        return None;
    }
    let parent = elements
        .last()
        .and_then(|element| element.typed_block_owner);
    let content_owner = match parent {
        Some(TypedBlockOwner::Body | TypedBlockOwner::Content(BlockControlOwner::Body)) => {
            Some(BlockControlOwner::Body)
        }
        Some(TypedBlockOwner::Table | TypedBlockOwner::Content(BlockControlOwner::Table)) => {
            Some(BlockControlOwner::Table)
        }
        Some(TypedBlockOwner::Row | TypedBlockOwner::Content(BlockControlOwner::Row)) => {
            Some(BlockControlOwner::Row)
        }
        Some(TypedBlockOwner::Cell | TypedBlockOwner::Content(BlockControlOwner::Cell)) => {
            Some(BlockControlOwner::Cell)
        }
        _ => None,
    };
    match local {
        b"document" if elements.is_empty() => Some(TypedBlockOwner::Document),
        b"body" if parent == Some(TypedBlockOwner::Document) => Some(TypedBlockOwner::Body),
        b"p" if matches!(
            content_owner,
            Some(BlockControlOwner::Body | BlockControlOwner::Cell)
        ) =>
        {
            Some(TypedBlockOwner::Paragraph)
        }
        b"tbl"
            if matches!(
                content_owner,
                Some(BlockControlOwner::Body | BlockControlOwner::Cell)
            ) =>
        {
            Some(TypedBlockOwner::Table)
        }
        b"tr" if content_owner == Some(BlockControlOwner::Table) => Some(TypedBlockOwner::Row),
        b"tc" if content_owner == Some(BlockControlOwner::Row) => Some(TypedBlockOwner::Cell),
        b"sdt" if content_owner.is_some() => {
            Some(TypedBlockOwner::ContentControl(content_owner.unwrap()))
        }
        b"sdtContent"
            if matches!(parent, Some(TypedBlockOwner::ContentControl(_)))
                && elements
                    .last()
                    .is_some_and(|control| !control.sdt_content_seen) =>
        {
            let Some(TypedBlockOwner::ContentControl(owner)) = parent else {
                unreachable!();
            };
            Some(TypedBlockOwner::Content(owner))
        }
        _ => None,
    }
}

fn direct_word_run_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    paragraph.is_some()
        && elements.last().is_some_and(|parent| {
            parent.is_typed_inline_owner
                && parent.local_name == b"r"
                && parent.paragraph == paragraph
        })
}

fn dynamic_toc_run_position(
    word: bool,
    local: &[u8],
    is_typed_inline_owner: bool,
    modeled_simple_field: bool,
    elements: &mut [DynamicXmlElement],
    paragraph: Option<usize>,
    next_boundaries: &mut [usize],
    nested_run_orders: &mut [HashMap<(usize, TocRawOrder), usize>],
    paragraph_raw_before: &mut [usize],
) -> Option<TocRunPosition> {
    let paragraph = paragraph?;
    let next = next_boundaries.get_mut(paragraph)?;
    let inherited = elements.last().and_then(|parent| parent.run_position);
    if word && local == b"r" && is_typed_inline_owner {
        let nested = elements.iter().rev().take_while(|element| {
            !element.is_typed_paragraph || element.paragraph != Some(paragraph)
        });
        if nested.into_iter().any(|element| {
            matches!(
                element.local_name.as_slice(),
                b"sdt" | b"sdtContent" | b"ins" | b"moveTo"
            )
        }) {
            let outer = inherited.unwrap_or(TocRunPosition {
                run_boundary: *next,
                raw_order: TocRawOrder::Raw(0),
                nested_order: 0,
            });
            return nested_leaf_position(paragraph, outer, nested_run_orders);
        }
        let boundary = *next;
        let position = nested_leaf_position(
            paragraph,
            TocRunPosition {
                run_boundary: boundary,
                raw_order: TocRawOrder::AfterRaw,
                nested_order: 0,
            },
            nested_run_orders,
        );
        *next += 1;
        *paragraph_raw_before.get_mut(paragraph)? = 0;
        return position;
    }
    if word
        && local == b"fldSimple"
        && modeled_simple_field
        && accepted_simple_field_parent(elements, paragraph)
    {
        if elements.last().is_some_and(|parent| {
            parent.is_typed_inline_owner
                && matches!(parent.local_name.as_slice(), b"ins" | b"moveTo")
        }) {
            let outer = inherited.unwrap_or(TocRunPosition {
                run_boundary: *next,
                raw_order: TocRawOrder::Raw(0),
                nested_order: 0,
            });
            return nested_leaf_position(paragraph, outer, nested_run_orders);
        }
        let boundary = *next;
        let position = nested_leaf_position(
            paragraph,
            TocRunPosition {
                run_boundary: boundary,
                raw_order: TocRawOrder::AfterRaw,
                nested_order: 0,
            },
            nested_run_orders,
        );
        *next += 1;
        *paragraph_raw_before.get_mut(paragraph)? = 0;
        return position;
    }
    if is_typed_inline_owner && local == b"hyperlink" {
        return inherited;
    }
    if is_typed_inline_owner && matches!(local, b"sdt" | b"sdtContent" | b"ins" | b"moveTo") {
        if let Some(inherited) = inherited {
            return Some(inherited);
        }
        let raw_order = if elements
            .last()
            .is_some_and(|parent| parent.is_typed_inline_owner && parent.local_name == b"hyperlink")
        {
            let parent = elements.last_mut()?;
            let plan = parent.hyperlink_plan.as_mut()?;
            let order = *plan.revision_orders.get(plan.next_revision)?;
            plan.next_revision += 1;
            order
        } else {
            TocRawOrder::Raw(*paragraph_raw_before.get(paragraph)?)
        };
        return Some(TocRunPosition {
            run_boundary: *next,
            raw_order,
            nested_order: 0,
        });
    }
    None
}

fn advance_direct_paragraph_raw_child(
    word: bool,
    local: &[u8],
    is_typed_inline_owner: bool,
    modeled_simple_field: bool,
    direct_paragraph_child: bool,
    paragraph: Option<usize>,
    hyperlink_plan: Option<&DynamicHyperlinkPlan>,
    paragraph_raw_before: &mut [usize],
) {
    if !direct_paragraph_child {
        return;
    }
    let modeled_without_raw = word
        && (local == b"pPr"
            || local == b"r" && is_typed_inline_owner
            || local == b"fldSimple" && modeled_simple_field
            || local == b"sdt" && is_typed_inline_owner
            || local == b"hyperlink");
    if modeled_without_raw
        || hyperlink_plan.is_some_and(|plan| plan.preserved_raw)
        || local == b"hyperlink"
    {
        return;
    }
    if let Some(raw_before) = paragraph.and_then(|index| paragraph_raw_before.get_mut(index)) {
        *raw_before += 1;
    }
}

fn dynamic_hyperlink_plan(
    xml: &[u8],
    start: usize,
    bindings: &BTreeMap<String, String>,
    raw_before: usize,
) -> Result<DynamicHyperlinkPlan> {
    let raw = dynamic_element_slice(xml, start)?;
    let mut scoped = Vec::new();
    scoped.extend_from_slice(b"<rdocx-scope");
    for (prefix, namespace) in bindings {
        if prefix == "xml" {
            continue;
        }
        if prefix.is_empty() {
            scoped.extend_from_slice(b" xmlns=\"");
        } else {
            scoped.extend_from_slice(b" xmlns:");
            scoped.extend_from_slice(prefix.as_bytes());
            scoped.extend_from_slice(b"=\"");
        }
        scoped.extend_from_slice(xml_escape_attribute(namespace).as_bytes());
        scoped.push(b'"');
    }
    scoped.push(b'>');
    scoped.extend_from_slice(raw);
    scoped.extend_from_slice(b"</rdocx-scope>");

    let mut reader = NsReader::from_reader(scoped.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut direct_runs = 0usize;
    let mut revision_boundaries = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid hyperlink XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        match event {
            Event::Start(element) => {
                if depth == 2 && word && local_name(element.name().as_ref()) == b"r" {
                    direct_runs += 1;
                } else if depth == 2
                    && word
                    && matches!(local_name(element.name().as_ref()), b"ins" | b"moveTo")
                {
                    let valid_id = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, id)| id.parse::<i32>().is_ok());
                    let has_author = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"author",
                        AttributeNamespace::Word,
                    )?
                    .is_some();
                    if valid_id && has_author {
                        revision_boundaries.push(direct_runs);
                    }
                }
                depth += 1;
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "table of contents hyperlink is not balanced".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
    let preserved_raw = direct_runs == 0;
    let revision_orders = revision_boundaries
        .into_iter()
        .map(|boundary| {
            if preserved_raw {
                TocRawOrder::Raw(raw_before)
            } else if boundary == direct_runs {
                TocRawOrder::BeforeRaw
            } else {
                TocRawOrder::AfterRaw
            }
        })
        .collect();
    Ok(DynamicHyperlinkPlan {
        revision_orders,
        next_revision: 0,
        preserved_raw,
    })
}

fn dynamic_element_slice(xml: &[u8], start: usize) -> Result<&[u8]> {
    let tail = xml.get(start..).ok_or_else(|| {
        Error::Other("table of contents element offset is outside the document".to_owned())
    })?;
    let mut reader = quick_xml::Reader::from_reader(tail);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?
        {
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => {
                return Ok(&tail[..reader.buffer_position() as usize]);
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok(&tail[..reader.buffer_position() as usize]);
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "table of contents element is not balanced".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn nested_leaf_position(
    paragraph: usize,
    outer: TocRunPosition,
    nested_run_orders: &mut [HashMap<(usize, TocRawOrder), usize>],
) -> Option<TocRunPosition> {
    let next = nested_run_orders
        .get_mut(paragraph)?
        .entry((outer.run_boundary, outer.raw_order))
        .or_default();
    let nested_order = *next;
    *next += 1;
    Some(TocRunPosition {
        nested_order,
        ..outer
    })
}

fn accepted_simple_field_parent(elements: &[DynamicXmlElement], paragraph: usize) -> bool {
    direct_typed_paragraph_parent(elements, Some(paragraph))
        || elements.last().is_some_and(|parent| {
            parent.is_typed_inline_owner
                && matches!(parent.local_name.as_slice(), b"ins" | b"moveTo")
                && parent.paragraph == Some(paragraph)
        })
}

fn typed_inline_owner(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    word: bool,
    local: &[u8],
    elements: &[DynamicXmlElement],
    paragraph: Option<usize>,
) -> Result<bool> {
    if !word || paragraph.is_none() {
        return Ok(false);
    }
    let Some(parent) = elements.last() else {
        return Ok(false);
    };
    let parent_is_paragraph = parent.is_typed_paragraph && parent.paragraph == paragraph;
    let parent_is_inline = parent.is_typed_inline_owner && parent.paragraph == paragraph;
    let parent_local = parent.local_name.as_slice();
    let valid = match local {
        b"r" => {
            parent_is_paragraph
                || (parent_is_inline
                    && matches!(
                        parent_local,
                        b"hyperlink" | b"sdtContent" | b"ins" | b"moveTo"
                    ))
        }
        b"hyperlink" => {
            parent_is_paragraph || (parent_is_inline && matches!(parent_local, b"ins" | b"moveTo"))
        }
        b"sdt" => {
            parent_is_paragraph
                || (parent_is_inline && matches!(parent_local, b"sdtContent" | b"ins" | b"moveTo"))
        }
        b"sdtContent" => parent_is_inline && parent_local == b"sdt" && !parent.sdt_content_seen,
        b"ins" | b"moveTo" => {
            let valid_parent = parent_is_paragraph
                || (parent_is_inline
                    && matches!(
                        parent_local,
                        b"hyperlink" | b"sdtContent" | b"ins" | b"moveTo"
                    ));
            let valid_id =
                resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
                    .is_some_and(|(_, id)| id.parse::<i32>().is_ok());
            let has_author =
                resolved_element_attribute(element, resolver, b"author", AttributeNamespace::Word)?
                    .is_some();
            valid_parent
                && parent.revision_depth < MAX_DYNAMIC_REVISION_NESTING_DEPTH
                && valid_id
                && has_author
        }
        _ => false,
    };
    Ok(valid)
}

fn direct_typed_paragraph_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    let Some(paragraph) = paragraph else {
        return false;
    };
    elements
        .last()
        .is_some_and(|parent| parent.is_typed_paragraph && parent.paragraph == Some(paragraph))
}

fn accepted_bookmark_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    direct_typed_paragraph_parent(elements, paragraph)
        || paragraph.is_some()
            && elements.last().is_some_and(|parent| {
                parent.is_typed_inline_owner
                    && matches!(
                        parent.local_name.as_slice(),
                        b"hyperlink" | b"sdtContent" | b"ins" | b"moveTo"
                    )
                    && parent.paragraph == paragraph
            })
}

fn accepted_simple_toc_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    direct_typed_paragraph_parent(elements, paragraph)
        || paragraph.is_some()
            && elements.last().is_some_and(|parent| {
                parent.is_typed_inline_owner
                    && matches!(parent.local_name.as_slice(), b"ins" | b"moveTo")
                    && parent.paragraph == paragraph
            })
}

fn update_dynamic_field_stack(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    event_start: usize,
    paragraph: Option<usize>,
    elements: &[DynamicXmlElement],
    fields: &mut Vec<DynamicFieldScan>,
    spans: &mut Vec<DynamicTocSpan>,
    generated: DynamicOwnerPolicy,
) -> Result<()> {
    let Some(paragraph) = paragraph else {
        return Ok(());
    };
    let kind =
        resolved_element_attribute(element, resolver, b"fldCharType", AttributeNamespace::Word)?
            .map(|(_, value)| value);
    match kind.as_deref() {
        Some("begin") => fields.push(DynamicFieldScan {
            instruction: String::new(),
            field_start: event_start,
            begin_paragraph: paragraph,
            begin_paragraph_start: elements
                .iter()
                .rev()
                .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph))
                .map(|element| element.start)
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents begin marker is not contained by a paragraph".to_owned(),
                    )
                })?,
            begin_run_start: elements
                .iter()
                .rev()
                .find(|element| {
                    element.is_word
                        && element.local_name == b"r"
                        && element.paragraph == Some(paragraph)
                })
                .map(|element| element.start)
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents begin marker is not contained by a run".to_owned(),
                    )
                })?,
            separator_paragraph: None,
            separator_run_start: None,
            result_start: None,
            result_start_position: None,
            start_paragraph_name: None,
            start_paragraph_namespaces: BTreeMap::new(),
            separator_wrapper_names: Vec::new(),
            instruction_runs: Vec::new(),
        }),
        Some("separate") => {
            let Some(field) = fields.last_mut() else {
                return Ok(());
            };
            if field.separator_paragraph.is_some() {
                return Err(Error::Other(
                    "table of contents field has more than one separator".to_owned(),
                ));
            }
            let run = elements
                .iter()
                .rev()
                .find(|element| {
                    element.is_word
                        && element.local_name == b"r"
                        && element.paragraph == Some(paragraph)
                })
                .ok_or_else(|| {
                    Error::Other("table of contents separator is not contained by a run".to_owned())
                })?;
            let para = elements
                .iter()
                .rev()
                .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph))
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents separator is not contained by a paragraph".to_owned(),
                    )
                })?;
            field.separator_paragraph = Some(paragraph);
            field.separator_run_start = Some(run.start);
            field.result_start_position = run.run_position.map(|position| {
                if position.raw_order == TocRawOrder::AfterRaw {
                    TocRunPosition {
                        run_boundary: position.run_boundary + 1,
                        raw_order: TocRawOrder::BeforeRaw,
                        nested_order: 0,
                    }
                } else {
                    TocRunPosition {
                        nested_order: position.nested_order + 1,
                        ..position
                    }
                }
            });
            field.start_paragraph_name = Some(para.qualified_name.clone());
            field.start_paragraph_namespaces = para.inherited_namespaces.clone();
            let paragraph_position = elements
                .iter()
                .position(|element| std::ptr::eq(element, para))
                .expect("paragraph was borrowed from the element stack");
            let run_position = elements
                .iter()
                .rposition(|element| std::ptr::eq(element, run))
                .expect("run was borrowed from the element stack");
            field.separator_wrapper_names = elements[paragraph_position + 1..run_position]
                .iter()
                .filter(|element| element.is_typed_inline_owner)
                .map(|element| element.qualified_name.clone())
                .collect();
        }
        Some("end") => {
            let Some(field) = fields.pop() else {
                return Ok(());
            };
            if !generated_table_opcode(
                &Field::new(&field.instruction, "").instruction.name,
                generated,
            ) {
                return Ok(());
            }
            let result_start = field.result_start.ok_or_else(|| {
                Error::Other("table of contents field is missing its separator".to_owned())
            })?;
            let separator_paragraph = field.separator_paragraph.ok_or_else(|| {
                Error::Other("table of contents field is missing its separator".to_owned())
            })?;
            if separator_paragraph != field.begin_paragraph {
                return Err(Error::Other(
                    "table of contents instruction crosses a paragraph boundary".to_owned(),
                ));
            }
            let end_run = elements
                .iter()
                .rev()
                .find(|element| {
                    element.is_word
                        && element.local_name == b"r"
                        && element.paragraph == Some(paragraph)
                })
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents end marker is not contained by a run".to_owned(),
                    )
                })?;
            let end_para = elements
                .iter()
                .rev()
                .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph))
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents end marker is not contained by a paragraph".to_owned(),
                    )
                })?;
            if event_start < result_start || end_run.start < result_start {
                return Err(Error::Other(
                    "table of contents result boundaries are reversed".to_owned(),
                ));
            }
            let paragraph_position = elements
                .iter()
                .position(|element| std::ptr::eq(element, end_para))
                .expect("end paragraph was borrowed from the element stack");
            let run_position = elements
                .iter()
                .rposition(|element| std::ptr::eq(element, end_run))
                .expect("end run was borrowed from the element stack");
            let wrapper_chain = &elements[paragraph_position + 1..run_position];
            let mut end_wrapper_prefixes = Vec::new();
            let mut wrapper_index = 0usize;
            while wrapper_index < wrapper_chain.len() {
                let wrapper = &wrapper_chain[wrapper_index];
                if !wrapper.is_typed_inline_owner {
                    wrapper_index += 1;
                    continue;
                }
                if wrapper.local_name == b"sdt"
                    && let Some(content) = wrapper_chain.get(wrapper_index + 1)
                    && content.is_typed_inline_owner
                    && content.local_name == b"sdtContent"
                {
                    end_wrapper_prefixes.push((wrapper.start, content.start_tag_end));
                    wrapper_index += 2;
                } else {
                    end_wrapper_prefixes.push((wrapper.start, wrapper.start_tag_end));
                    wrapper_index += 1;
                }
            }
            spans.push(DynamicTocSpan {
                simple_field: None,
                instruction: field.instruction,
                field_start: field.field_start,
                field_end: event_start,
                begin_paragraph: field.begin_paragraph,
                end_paragraph: paragraph,
                begin_run_start: field.begin_run_start,
                instruction_paragraph_start: field.begin_paragraph_start,
                result_start,
                result_end: end_run.start,
                result_start_position: field.result_start_position.ok_or_else(|| {
                    Error::Other(
                        "table of contents separator run position was not retained".to_owned(),
                    )
                })?,
                result_end_position: end_run.run_position.ok_or_else(|| {
                    Error::Other("table of contents end run position was not retained".to_owned())
                })?,
                end_run_end: 0,
                start_paragraph_name: field.start_paragraph_name.ok_or_else(|| {
                    Error::Other("table of contents field is missing its separator".to_owned())
                })?,
                start_paragraph_namespaces: field.start_paragraph_namespaces,
                separator_wrapper_names: field.separator_wrapper_names,
                instruction_runs: field.instruction_runs,
                end_paragraph_start: end_para.start,
                end_paragraph_content_start: 0,
                end_wrapper_prefixes,
            });
        }
        _ => {}
    }
    Ok(())
}

type ParsedDynamicTocFields = (Vec<Option<TocField>>, Vec<(usize, String)>);

fn parse_dynamic_toc_fields(
    document: &Document,
    xml: &[u8],
    spans: &[DynamicTocSpan],
) -> Result<ParsedDynamicTocFields> {
    let mut projected = document.clone_for_staging();
    prepare_physical_story_projection(&mut projected.document.body, &mut [])?;
    let document = &projected;
    let mut paragraphs = Vec::new();
    collect_body_paragraphs(&document.document.body, &mut paragraphs);
    let context = FieldEvaluationContext::default();
    let mut output = Vec::with_capacity(spans.len());
    let mut diagnostics = Vec::new();
    for span in spans {
        let field = parse_dynamic_toc_field(xml, span)?;
        let mut evaluator = Evaluator::new(document, &context);
        evaluator.ensure_numbering_layout_for_field(&field)?;
        match evaluator.evaluate_field(&field, "main", &paragraphs, span.begin_paragraph) {
            FieldOutcome::TableOfContents(toc) => output.push(Some(toc)),
            FieldOutcome::KeepStored { diagnostic } => {
                diagnostics.push((span.field_start, diagnostic));
                output.push(None);
            }
            _ => {
                return Err(Error::Other(
                    "table of contents instruction was not recognized".to_owned(),
                ));
            }
        }
    }
    Ok((output, diagnostics))
}

fn parse_dynamic_toc_field(xml: &[u8], span: &DynamicTocSpan) -> Result<Field> {
    if let Some(field) = &span.simple_field {
        return Ok(field.clone());
    }
    let prefix = span
        .start_paragraph_name
        .split_once(':')
        .map_or("w", |(prefix, _)| prefix);
    // The instruction paragraph is cut out of its part, so its start tag gets
    // the declarations it inherits there. Every run in it then resolves the
    // prefixes it resolved when the document was read.
    let mut source = Vec::new();
    append_with_inherited_namespaces(
        &mut source,
        &xml[span.instruction_paragraph_start..span.result_start],
        &span.start_paragraph_namespaces,
    )?;
    source.extend_from_slice(
        format!("<{prefix}:r><{prefix}:fldChar {prefix}:fldCharType=\"end\"/></{prefix}:r>")
            .as_bytes(),
    );
    append_toc_wrapper_closures(&mut source, &span.separator_wrapper_names);
    source.extend_from_slice(b"</");
    source.extend_from_slice(span.start_paragraph_name.as_bytes());
    source.push(b'>');
    let parse_paragraph = |source: &[u8]| -> Result<CT_P> {
        let mut reader = quick_xml::Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).map_err(|error| {
                Error::Other(format!("invalid table of contents field: {error}"))
            })? {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    return Ok(CT_P::from_xml(&mut reader)?);
                }
                Event::Eof => {
                    return Err(Error::Other(
                        "table of contents instruction paragraph was not found".to_owned(),
                    ));
                }
                _ => {}
            }
            buffer.clear();
        }
    };
    CT_P::from_xml_fragment(&source)?;

    let mut projected = format!("<w:p xmlns:w=\"{W_NS}\">").into_bytes();
    for run in &span.instruction_runs {
        append_with_inherited_namespaces(
            &mut projected,
            &xml[run.start..run.end],
            &run.inherited_namespaces,
        )?;
    }
    projected.extend_from_slice(b"<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
    let paragraph = parse_paragraph(&projected)?;
    accepted_toc_runs(&paragraph)
        .into_iter()
        .flat_map(|run| &run.run.content)
        .find_map(|content| match content {
            RunContent::Field(field)
                if field.effective_instruction().name
                    == Field::new(&span.instruction, "").instruction.name =>
            {
                Some(field.clone())
            }
            _ => None,
        })
        .ok_or_else(|| {
            Error::Other(format!(
                "table of contents instruction could not be parsed: {}",
                span.instruction.trim()
            ))
        })
}

fn append_with_inherited_namespaces(
    output: &mut Vec<u8>,
    raw: &[u8],
    inherited_namespaces: &BTreeMap<String, String>,
) -> Result<()> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (insertion, local_namespaces) =
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "invalid table of contents instruction XML: {error}"
            ))
        })? {
            Event::Start(start) | Event::Empty(start) => {
                let mut local_namespaces = HashSet::new();
                for attribute in start.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!(
                            "invalid table of contents instruction XML: {error}"
                        ))
                    })?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" {
                        local_namespaces.insert(String::new());
                    } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                        local_namespaces.insert(String::from_utf8_lossy(prefix).into_owned());
                    }
                }
                let tag_end = reader.buffer_position() as usize;
                let insertion = if tag_end >= 2 && raw[tag_end - 2] == b'/' {
                    tag_end - 2
                } else {
                    tag_end - 1
                };
                (insertion, local_namespaces)
            }
            _ => {
                return Err(Error::Other(
                    "table of contents instruction XML has no start tag".to_owned(),
                ));
            }
        };
    output.extend_from_slice(&raw[..insertion]);
    for (prefix, namespace) in inherited_namespaces {
        if prefix == "xml" || local_namespaces.contains(prefix) {
            continue;
        }
        if prefix.is_empty() {
            output.extend_from_slice(b" xmlns=\"");
        } else {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
        }
        output.extend_from_slice(xml_escape_attribute(namespace).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(&raw[insertion..]);
    Ok(())
}

#[derive(Debug, Clone)]
struct TocBookmark {
    id: i32,
    name: String,
    insert: bool,
    authored: bool,
}

#[derive(Debug)]
struct TocBookmarkState {
    max_id: i32,
    names: HashSet<String>,
    named_ranges: HashMap<String, (TocDocumentPosition, TocDocumentPosition)>,
    whole_paragraphs: HashMap<usize, (i32, String)>,
    ranges: Vec<TocBookmarkRange>,
}

#[derive(Debug, Clone)]
struct TocBookmarkRange {
    id: i32,
    name: String,
    start_offset: usize,
    end_offset: usize,
}

#[derive(Debug)]
struct TocBookmarkRepair {
    id: i32,
    name: String,
    toc_index: usize,
    replace_start: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
enum TocRawOrder {
    BeforeRaw,
    Raw(usize),
    AfterRaw,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct TocRunPosition {
    run_boundary: usize,
    raw_order: TocRawOrder,
    nested_order: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct TocDocumentPosition {
    paragraph: usize,
    accepted_run_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct TocOwnedPosition {
    paragraph: usize,
    run: TocRunPosition,
}

fn inspect_toc_bookmarks(body: &CT_Body, xml: &[u8]) -> Result<TocBookmarkState> {
    let whole_bookmarks = whole_paragraph_bookmark_ids(xml)?;
    let marker_offsets = toc_bookmark_marker_offsets(xml)?;
    let mut paragraphs = Vec::new();
    collect_body_paragraphs(body, &mut paragraphs);
    let mut starts = HashMap::<i32, (String, usize, usize)>::new();
    let mut start_order = Vec::new();
    let mut ends = HashMap::<i32, (usize, usize)>::new();
    let mut names = HashSet::new();
    let mut max_id = 0i32;
    for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
        for marker in &paragraph.bookmark_markers {
            let Some(id) = marker.id() else {
                return Err(Error::Other(
                    "table of contents source has a bookmark without an ID".to_owned(),
                ));
            };
            max_id = max_id.max(id);
            if marker.is_start() {
                let name = marker.name().ok_or_else(|| {
                    Error::Other(
                        "table of contents source has a bookmark without a name".to_owned(),
                    )
                })?;
                if !names.insert(name.to_owned()) || starts.contains_key(&id) {
                    return Err(Error::Other(
                        "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                    ));
                }
                starts.insert(
                    id,
                    (
                        name.to_owned(),
                        paragraph_index,
                        marker.projected_run_index(),
                    ),
                );
                start_order.push(id);
            } else if ends
                .insert(id, (paragraph_index, marker.projected_run_index()))
                .is_some()
            {
                return Err(Error::Other(
                    "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                ));
            }
        }
    }
    if starts.len() != ends.len() || starts.keys().any(|id| !ends.contains_key(id)) {
        return Err(Error::Other(
            "table of contents source has an unmatched bookmark".to_owned(),
        ));
    }
    let mut named_ranges = HashMap::new();
    let mut whole_paragraphs = HashMap::new();
    let mut ranges = Vec::new();
    for id in start_order {
        let (name, start_paragraph, start_run) = starts
            .remove(&id)
            .expect("bookmark start order contains each validated start once");
        let (end_paragraph, end_run) = ends[&id];
        let start_offset = marker_offsets.get(&(id, true)).copied().ok_or_else(|| {
            Error::Other("table of contents bookmark start offset was not retained".to_owned())
        })?;
        let end_offset = marker_offsets.get(&(id, false)).copied().ok_or_else(|| {
            Error::Other("table of contents bookmark end offset was not retained".to_owned())
        })?;
        if start_offset > end_offset {
            return Err(Error::Other(
                "table of contents source has a reversed bookmark".to_owned(),
            ));
        }
        named_ranges.insert(
            name.clone(),
            (
                TocDocumentPosition {
                    paragraph: start_paragraph,
                    accepted_run_index: start_run,
                },
                TocDocumentPosition {
                    paragraph: end_paragraph,
                    accepted_run_index: end_run,
                },
            ),
        );
        ranges.push(TocBookmarkRange {
            id,
            name: name.clone(),
            start_offset,
            end_offset,
        });
        if start_paragraph == end_paragraph && whole_bookmarks.contains(&id) {
            whole_paragraphs
                .entry(start_paragraph)
                .or_insert((id, name));
        }
    }
    Ok(TocBookmarkState {
        max_id,
        names,
        named_ranges,
        whole_paragraphs,
        ranges,
    })
}

fn toc_crossing_bookmark_repairs(
    state: &TocBookmarkState,
    spans: &[DynamicTocSpan],
    fields: &[Option<TocField>],
) -> Vec<TocBookmarkRepair> {
    state
        .ranges
        .iter()
        .filter_map(|range| {
            let containing_span = |offset: usize| {
                spans
                    .iter()
                    .zip(fields)
                    .enumerate()
                    .find(|(_, (span, field))| {
                        field.is_some() && offset >= span.result_start && offset < span.result_end
                    })
                    .map(|(index, _)| index)
            };
            let start_span = containing_span(range.start_offset);
            let end_span = containing_span(range.end_offset);
            match (start_span, end_span) {
                (Some(toc_index), None) => Some(TocBookmarkRepair {
                    id: range.id,
                    name: range.name.clone(),
                    toc_index,
                    replace_start: true,
                }),
                (None, Some(toc_index)) => Some(TocBookmarkRepair {
                    id: range.id,
                    name: range.name.clone(),
                    toc_index,
                    replace_start: false,
                }),
                _ => None,
            }
        })
        .collect()
}

#[derive(Clone, Copy)]
enum TocParagraphToken {
    Content,
    BookmarkStart(i32),
    BookmarkEnd(i32),
}

fn toc_bookmark_marker_offsets(xml: &[u8]) -> Result<HashMap<(i32, bool), usize>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut offsets = HashMap::new();
    let mut paragraph_count = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                invalidate_overdeep_revision_owner(word, &local, &mut elements);
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    let paragraph = paragraph_count;
                    paragraph_count += 1;
                    Some(paragraph)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let mut is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                if word
                    && matches!(local.as_slice(), b"bookmarkStart" | b"bookmarkEnd")
                    && accepted_bookmark_parent(&elements, paragraph)
                    && let Some((_, id)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    && let Ok(id) = id.parse::<i32>()
                    && offsets
                        .insert((id, local == b"bookmarkStart"), before)
                        .is_some()
                {
                    return Err(Error::Other(
                        "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                    ));
                }
                let revision_depth = if is_typed_inline_owner {
                    elements.last().map_or(0, |parent| {
                        parent.revision_depth
                            + usize::from(matches!(local.as_slice(), b"ins" | b"moveTo"))
                    })
                } else {
                    0
                };
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner,
                    revision_depth,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                invalidate_overdeep_revision_owner(word, &local, &mut elements);
                let paragraph = elements.last().and_then(|element| element.paragraph);
                let typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                if word
                    && matches!(local.as_slice(), b"bookmarkStart" | b"bookmarkEnd")
                    && accepted_bookmark_parent(&elements, paragraph)
                    && let Some((_, id)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    && let Ok(id) = id.parse::<i32>()
                    && offsets
                        .insert((id, local == b"bookmarkStart"), before)
                        .is_some()
                {
                    return Err(Error::Other(
                        "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                    ));
                }
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
            }
            Event::End(_) => {
                elements.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(offsets)
}

fn whole_paragraph_bookmark_ids(xml: &[u8]) -> Result<HashSet<i32>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut paragraphs = Vec::<Vec<TocParagraphToken>>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let mut is_typed_inline_owner = false;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    paragraphs.push(Vec::new());
                    Some(paragraphs.len() - 1)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                if let Some(index) = paragraph
                    && elements
                        .last()
                        .is_some_and(|parent| parent.is_typed_paragraph)
                    && local != b"pPr"
                {
                    paragraphs[index].push(paragraph_token(
                        &element,
                        reader.resolver(),
                        word,
                        &local,
                    )?);
                }
                mark_typed_sdt_content(&mut elements, &local, typed_block_owner, false);
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner: false,
                    revision_depth: 0,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                let typed_block_owner = classify_typed_block_owner(word, local, &elements);
                mark_typed_sdt_content(&mut elements, local, typed_block_owner, false);
                if let Some(parent) = elements.last()
                    && parent.is_typed_paragraph
                    && local != b"pPr"
                    && let Some(index) = parent.paragraph
                {
                    paragraphs[index].push(paragraph_token(
                        &element,
                        reader.resolver(),
                        word,
                        local,
                    )?);
                }
            }
            Event::End(_) => {
                elements.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut whole = HashSet::new();
    for tokens in paragraphs {
        let content = tokens
            .iter()
            .enumerate()
            .filter_map(|(index, token)| {
                matches!(token, TocParagraphToken::Content).then_some(index)
            })
            .collect::<Vec<_>>();
        let mut starts = HashMap::new();
        let mut ends = HashMap::new();
        for (index, token) in tokens.iter().enumerate() {
            match token {
                TocParagraphToken::BookmarkStart(id) => {
                    starts.insert(*id, index);
                }
                TocParagraphToken::BookmarkEnd(id) => {
                    ends.insert(*id, index);
                }
                TocParagraphToken::Content => {}
            }
        }
        for (id, start) in starts {
            let Some(end) = ends.get(&id).copied() else {
                continue;
            };
            if content
                .iter()
                .all(|position| *position > start && *position < end)
            {
                whole.insert(id);
            }
        }
    }
    Ok(whole)
}

fn paragraph_token(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    word: bool,
    local: &[u8],
) -> Result<TocParagraphToken> {
    if word && matches!(local, b"bookmarkStart" | b"bookmarkEnd") {
        let id = resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
            .and_then(|(_, id)| id.parse::<i32>().ok());
        if let Some(id) = id {
            return Ok(if local == b"bookmarkStart" {
                TocParagraphToken::BookmarkStart(id)
            } else {
                TocParagraphToken::BookmarkEnd(id)
            });
        }
    }
    Ok(TocParagraphToken::Content)
}

struct TocBookmarkAllocator {
    next_id: Option<i32>,
    next_name: Option<u32>,
    names: HashSet<String>,
    whole_paragraphs: HashMap<usize, (i32, String)>,
}

impl TocBookmarkAllocator {
    fn new(state: TocBookmarkState, authored_names: &HashSet<String>) -> Self {
        let mut names = state.names;
        names.retain(|name| !authored_names.contains(name));
        Self {
            next_id: state.max_id.checked_add(1),
            next_name: Some(1),
            names,
            whole_paragraphs: state.whole_paragraphs,
        }
    }

    fn whole_paragraph_name(&self, paragraph: usize) -> Option<(i32, String)> {
        self.whole_paragraphs.get(&paragraph).cloned()
    }

    fn allocate(&mut self, identifiers: &mut DocumentIdentifiers) -> Result<(i32, String)> {
        let preferred = self.next_id.ok_or_else(|| {
            Error::Other("table of contents exhausted the bookmark ID range".to_owned())
        })?;
        let id = identifiers.reserve_preferred_bookmark_id(preferred)?;
        let name = self.allocate_name()?;
        self.next_id = id.checked_add(1);
        Ok((id, name))
    }

    fn allocate_name(&mut self) -> Result<String> {
        loop {
            let suffix = self.next_name.ok_or_else(|| {
                Error::Other("table of contents exhausted the bookmark name range".to_owned())
            })?;
            let name = format!("_Toc{suffix}");
            self.next_name = suffix.checked_add(1);
            if self.names.insert(name.clone()) {
                return Ok(name);
            }
        }
    }
}

#[derive(Debug, Clone)]
struct TocSource {
    paragraph_index: usize,
    level: u8,
    title: String,
    omit_page_number: bool,
    needs_bookmark: bool,
    sequence_prefix: Option<String>,
    numbering_prefix: Option<TocNumberingPrefix>,
}

#[derive(Debug, Clone)]
struct TocNumberingPrefix {
    marker: String,
    suffix: ST_LvlSuffix,
}

#[derive(Debug, Clone)]
struct TocEntryStyle {
    style_id: String,
    has_right_tab: bool,
}

fn discover_toc_sources(
    document: &Document,
    spans: &[DynamicTocSpan],
    fields: &[Option<TocField>],
    bookmarks: &TocBookmarkState,
    numbering_layout: &Arc<rdocx_layout::WordLayoutResult>,
) -> Result<Vec<Vec<TocSource>>> {
    let mut paragraphs = Vec::new();
    collect_body_paragraphs(&document.document.body, &mut paragraphs);
    let context = FieldEvaluationContext::default();
    let sequence_snapshot = Arc::new(
        rdocx_layout::engine::evaluate_sequence_fields(&document.build_layout_input()?).map_err(
            |error| Error::Other(format!("TOC sequence source evaluation failed: {error}")),
        )?,
    );
    let mut sequence_nodes = Vec::new();
    for index in 1u32.. {
        let Some(node) = rdocx_layout::SourceNodeId::new(index) else {
            break;
        };
        let Some(path) = sequence_snapshot.source_node(node) else {
            break;
        };
        if matches!(path.story, rdocx_layout::WordStory::Document) {
            sequence_nodes.push(node);
        }
    }
    if sequence_nodes.len() != paragraphs.len() {
        return Err(Error::Other(
            "TOC sequence paragraph source inventory is incomplete".into(),
        ));
    }
    let mut all_sources = Vec::with_capacity(fields.len());
    for toc in fields {
        let Some(toc) = toc else {
            all_sources.push(Vec::new());
            continue;
        };
        let bookmark_range = toc
            .bookmark
            .as_ref()
            .map(|name| {
                bookmarks.named_ranges.get(name).copied().ok_or_else(|| {
                    Error::Other(format!(
                        "table of contents source bookmark {name} was not found"
                    ))
                })
            })
            .transpose()?;
        let mut sources = Vec::new();
        let mut evaluator = Evaluator::new(document, &context);
        evaluator.bind_numbering_layout(Arc::clone(numbering_layout));
        evaluator.sequence_snapshot = Some(Arc::clone(&sequence_snapshot));
        let mut sequence_value = None;
        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            let paragraph_fully_owned = spans.iter().any(|span| {
                paragraph_index > span.begin_paragraph && paragraph_index < span.end_paragraph
            });
            if paragraph_fully_owned {
                continue;
            }
            evaluator.register_paragraph_fields(paragraph, Some(sequence_nodes[paragraph_index]));
            let accepted_runs = accepted_toc_runs(paragraph);
            for (accepted_run_index, run) in accepted_runs.iter().enumerate() {
                let position = TocDocumentPosition {
                    paragraph: paragraph_index,
                    accepted_run_index,
                };
                let in_bookmark =
                    bookmark_range.is_none_or(|(start, end)| position >= start && position < end);
                for content in &run.run.content {
                    let RunContent::Field(field) = content else {
                        continue;
                    };
                    let owned_position = TocOwnedPosition {
                        paragraph: paragraph_index,
                        run: TocRunPosition {
                            run_boundary: run.run_boundary,
                            raw_order: run.raw_order,
                            nested_order: run.nested_order,
                        },
                    };
                    if toc_source_position_is_owned(spans, owned_position) {
                        continue;
                    }
                    let outcome =
                        evaluator.evaluate_field(field, "main", &paragraphs, paragraph_index);
                    if field.instruction.name == "SEQ"
                        && let Some(identifier) = field.instruction.arguments.first()
                        && field_argument_text(identifier).is_some_and(|name| {
                            toc.sequence_identifier
                                .as_deref()
                                .is_some_and(|selected| selected.eq_ignore_ascii_case(name))
                        })
                        && let FieldOutcome::Resolved(ref value) = outcome
                    {
                        sequence_value = Some(value.clone());
                    }
                    if !in_bookmark {
                        continue;
                    }
                    let FieldOutcome::TableOfContentsEntry(tc) = outcome else {
                        continue;
                    };
                    if toc_accepts_tc(toc, &tc) {
                        let omit_page_number = tc.omit_page_number
                            || toc
                                .omit_page_number_levels
                                .is_some_and(|(start, end)| (start..=end).contains(&tc.level));
                        sources.push(TocSource {
                            paragraph_index,
                            level: tc.level,
                            title: tc.entry,
                            omit_page_number,
                            needs_bookmark: toc.hyperlink || !omit_page_number,
                            sequence_prefix: sequence_value.clone(),
                            numbering_prefix: toc_numbering_prefix(
                                numbering_layout,
                                paragraph_index,
                            ),
                        });
                    }
                }
            }
            let paragraph_in_bookmark = bookmark_range.is_none_or(|(start, end)| {
                accepted_runs
                    .iter()
                    .enumerate()
                    .any(|(accepted_run_index, run)| {
                        let position = TocDocumentPosition {
                            paragraph: paragraph_index,
                            accepted_run_index,
                        };
                        let owned_position = TocOwnedPosition {
                            paragraph: paragraph_index,
                            run: TocRunPosition {
                                run_boundary: run.run_boundary,
                                raw_order: run.raw_order,
                                nested_order: run.nested_order,
                            },
                        };
                        !toc_source_position_is_owned(spans, owned_position)
                            && position >= start
                            && position < end
                    })
            });
            if !paragraph_in_bookmark {
                continue;
            }
            let effective_heading = evaluator
                .numbering_source(sequence_nodes[paragraph_index])
                .and_then(|node| numbering_layout.paragraph_navigation_source(node))
                .map(|(_, level, ..)| level);
            let Some(level) = toc_paragraph_level(document, toc, paragraph, effective_heading)?
            else {
                continue;
            };
            let title = accepted_runs
                .iter()
                .filter(|run| {
                    !toc_source_position_is_owned(
                        spans,
                        TocOwnedPosition {
                            paragraph: paragraph_index,
                            run: TocRunPosition {
                                run_boundary: run.run_boundary,
                                raw_order: run.raw_order,
                                nested_order: run.nested_order,
                            },
                        },
                    )
                })
                .map(|run| run.run.text())
                .collect::<String>();
            if title.is_empty() {
                continue;
            }
            let omit_page_number = toc
                .omit_page_number_levels
                .is_some_and(|(start, end)| (start..=end).contains(&level));
            sources.push(TocSource {
                paragraph_index,
                level,
                title,
                omit_page_number,
                needs_bookmark: toc.hyperlink || !omit_page_number,
                sequence_prefix: sequence_value.clone(),
                numbering_prefix: toc_numbering_prefix(numbering_layout, paragraph_index),
            });
        }
        sources.sort_by_key(|source| source.paragraph_index);
        all_sources.push(sources);
    }
    Ok(all_sources)
}

fn toc_numbering_prefix(
    layout: &rdocx_layout::WordLayoutResult,
    paragraph_index: usize,
) -> Option<TocNumberingPrefix> {
    let numbering = layout.document_paragraph_numbering(paragraph_index)?;
    if numbering.marker_text.is_empty() {
        return None;
    }
    Some(TocNumberingPrefix {
        marker: numbering.marker_text.clone(),
        suffix: numbering.suffix,
    })
}

fn ensure_toc_entry_styles(
    document: &mut Document,
    sources: &[Vec<TocSource>],
) -> Result<(BTreeMap<u8, TocEntryStyle>, Vec<String>)> {
    let source_styles = toc_style_view(&document.styles);
    let mut levels = sources
        .iter()
        .flatten()
        .map(|source| source.level)
        .collect::<Vec<_>>();
    levels.sort_unstable();
    levels.dedup();
    let mut resolved = BTreeMap::new();
    for level in levels {
        let built_in_name = format!("toc {level}");
        let canonical_id = format!("TOC{level}");
        let style_id = document
            .styles
            .styles
            .iter()
            .find(|style| {
                // Entries reference their style by identifier, so a producer
                // style without one cannot be the entry style.
                style.style_type == StyleType::Paragraph
                    && !style.style_id.is_empty()
                    && style
                        .name
                        .as_deref()
                        .is_some_and(|name| name.trim().eq_ignore_ascii_case(&built_in_name))
            })
            .map(|style| style.style_id.clone())
            .or_else(|| {
                document
                    .styles
                    .get_by_id(&canonical_id)
                    .map(|style| style.style_id.clone())
            })
            .unwrap_or_else(|| {
                let (style, _, _) =
                    style::StyleBuilder::paragraph(&canonical_id, &format!("TOC {level}")).build();
                document.styles.styles.push(style);
                canonical_id
            });
        let effective = style::resolve_paragraph_properties(Some(&style_id), &document.styles);
        let has_right_tab = effective
            .tabs
            .as_ref()
            .is_some_and(|tabs| tabs.tabs.iter().any(|tab| tab.val == ST_TabJc::Right));
        resolved.insert(
            level,
            TocEntryStyle {
                style_id,
                has_right_tab,
            },
        );
    }
    // Producer defects the read surface accepts stay in the package and are
    // reported. Only a defect the staged entry styles introduce rejects.
    let retained =
        style::validate_style_graph_change(&source_styles, &toc_style_view(&document.styles))?
            .into_iter()
            .map(|defect| format!("{defect}, retained while rebuilding TOC"))
            .collect();
    Ok((resolved, retained))
}

/// The style graph a TOC rebuild resolves. It keeps the first definition of
/// each style ID and, for each style type, the default that
/// `CT_Styles::get_default` gives layout. The package keeps every definition.
fn toc_style_view(styles: &CT_Styles) -> CT_Styles {
    let mut view = styles.clone();
    let mut seen = HashSet::new();
    view.styles
        .retain(|style| seen.insert(style.style_id.clone()));
    for style in &mut view.styles {
        style.is_default = styles
            .get_default(style.style_type)
            .is_some_and(|default| default.style_id == style.style_id);
    }
    view
}

fn toc_default_style_diagnostics(styles: &CT_Styles) -> Vec<String> {
    [
        StyleType::Paragraph,
        StyleType::Character,
        StyleType::Table,
        StyleType::Numbering,
    ]
    .into_iter()
    .filter_map(|style_type| {
        let first = styles.get_default(style_type)?;
        styles
            .styles
            .iter()
            .any(|style| {
                style.style_type == style_type
                    && style.is_default
                    && style.style_id != first.style_id
            })
            .then(|| {
                format!(
                    "multiple default {} styles used first default '{}' while rebuilding TOC",
                    style_type.to_str(),
                    first.style_id
                )
            })
    })
    .collect()
}

fn toc_duplicate_style_diagnostics(styles: &CT_Styles) -> Vec<String> {
    let mut seen = HashSet::new();
    let mut reported = HashSet::new();
    styles
        .styles
        .iter()
        .filter_map(|style| {
            if seen.insert(style.style_id.as_str()) || !reported.insert(style.style_id.as_str()) {
                return None;
            }
            Some(format!(
                "duplicate style ID '{}' used first definition while rebuilding TOC",
                style.style_id
            ))
        })
        .collect()
}

fn toc_section_text_width(body: &CT_Body, paragraph_index: usize) -> i32 {
    const DEFAULT_PAGE_WIDTH: i32 = 12_240;
    const DEFAULT_MARGIN: i32 = 1_440;
    const DEFAULT_TEXT_WIDTH: i32 = DEFAULT_PAGE_WIDTH - 2 * DEFAULT_MARGIN;

    let mut paragraphs = Vec::new();
    collect_body_paragraphs(body, &mut paragraphs);
    let section = paragraphs
        .iter()
        .skip(paragraph_index)
        .find_map(|paragraph| {
            paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.sect_pr.as_ref())
        })
        .or(body.sect_pr.as_ref());
    let Some(section) = section else {
        return DEFAULT_TEXT_WIDTH;
    };
    let page_width = section
        .page_width
        .map(|width| width.0)
        .unwrap_or(DEFAULT_PAGE_WIDTH);
    let left = section
        .margin_left
        .map(|margin| margin.0)
        .unwrap_or(DEFAULT_MARGIN);
    let right = section
        .margin_right
        .map(|margin| margin.0)
        .unwrap_or(DEFAULT_MARGIN);
    page_width
        .checked_sub(left)
        .and_then(|width| width.checked_sub(right))
        .filter(|width| *width > 0)
        .unwrap_or(DEFAULT_TEXT_WIDTH)
}

fn toc_source_position_is_owned(spans: &[DynamicTocSpan], position: TocOwnedPosition) -> bool {
    spans.iter().any(|span| {
        #[cfg(test)]
        NAVIGATION_OWNER_VISITS.set(NAVIGATION_OWNER_VISITS.get() + 1);
        if span.begin_paragraph == span.end_paragraph && position.paragraph == span.begin_paragraph
        {
            return position.run >= span.result_start_position
                && position.run < span.result_end_position;
        }
        if position.paragraph > span.begin_paragraph && position.paragraph < span.end_paragraph {
            return true;
        }
        if position.paragraph == span.begin_paragraph {
            return position.run >= span.result_start_position;
        }
        if position.paragraph == span.end_paragraph {
            return position.run < span.result_end_position;
        }
        false
    })
}

#[derive(Clone, Copy)]
struct AcceptedTocRun<'a> {
    run: &'a CT_R,
    run_boundary: usize,
    raw_order: TocRawOrder,
    nested_order: usize,
}

fn accepted_toc_runs(paragraph: &CT_P) -> Vec<AcceptedTocRun<'_>> {
    let mut runs = Vec::new();
    append_accepted_paragraph_runs(paragraph, None, &mut runs);
    let mut nested_run_orders = HashMap::<(usize, TocRawOrder), usize>::new();
    for run in &mut runs {
        let next = nested_run_orders
            .entry((run.run_boundary, run.raw_order))
            .or_default();
        run.nested_order = *next;
        *next += 1;
    }
    runs
}

fn accepted_paragraph_owners(
    paragraph: &CT_P,
    boundary: usize,
) -> Vec<(TocRawOrder, AcceptedParagraphOwner<'_>)> {
    let mut owners = paragraph
        .content_controls
        .iter()
        .filter(|(at, _, _, _)| *at == boundary)
        .map(|(_, raw_before, _, control)| {
            (
                TocRawOrder::Raw(*raw_before),
                0u8,
                AcceptedParagraphOwner::Control(control),
            )
        })
        .chain(
            paragraph
                .revisions
                .iter()
                .filter(|(at, _, _)| *at == boundary)
                .map(|(_, slot, revision)| {
                    (
                        accepted_revision_raw_order(paragraph, boundary, *slot),
                        1u8,
                        AcceptedParagraphOwner::Revision(revision),
                    )
                }),
        )
        .collect::<Vec<_>>();
    owners.sort_by_key(|(raw_order, kind, _)| (*raw_order, *kind));
    owners
        .into_iter()
        .map(|(order, _, owner)| (order, owner))
        .collect()
}

fn append_accepted_paragraph_runs<'a>(
    paragraph: &'a CT_P,
    inherited_position: Option<(usize, TocRawOrder)>,
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=paragraph.runs.len() {
        for (raw_order, owner) in accepted_paragraph_owners(paragraph, boundary) {
            let position = inherited_position.unwrap_or((boundary, raw_order));
            match owner {
                AcceptedParagraphOwner::Control(control) => {
                    append_accepted_control_runs(control, position, runs);
                }
                AcceptedParagraphOwner::Revision(revision) => {
                    append_accepted_revision_runs(revision, position, runs);
                }
            }
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            let (run_boundary, raw_order) =
                inherited_position.unwrap_or((boundary, TocRawOrder::AfterRaw));
            runs.push(AcceptedTocRun {
                run,
                run_boundary,
                raw_order,
                nested_order: 0,
            });
        }
    }
}

fn accepted_revision_raw_order(paragraph: &CT_P, boundary: usize, slot: usize) -> TocRawOrder {
    let Some(index) = hyperlink_revision_index(slot) else {
        return TocRawOrder::Raw(slot);
    };
    if let Some(raw_before) = paragraph
        .hyperlinks
        .get(index)
        .and_then(|hyperlink| hyperlink.preserved_raw_before)
    {
        TocRawOrder::Raw(raw_before)
    } else if paragraph
        .hyperlinks
        .get(index)
        .is_some_and(|hyperlink| boundary == hyperlink.run_end)
    {
        TocRawOrder::BeforeRaw
    } else {
        TocRawOrder::AfterRaw
    }
}

enum AcceptedParagraphOwner<'a> {
    Control(&'a CT_Sdt),
    Revision(&'a CT_Revision),
}

fn append_accepted_control_runs<'a>(
    control: &'a CT_Sdt,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=control.content.len() {
        for (_, revision) in control.revisions().iter().filter(|(at, _)| *at == boundary) {
            append_accepted_revision_runs(revision, position, runs);
        }
        if let Some(content) = control.content.get(boundary) {
            match content {
                SdtContent::Run(run) => runs.push(AcceptedTocRun {
                    run,
                    run_boundary: position.0,
                    raw_order: position.1,
                    nested_order: 0,
                }),
                SdtContent::ContentControl(control) => {
                    append_accepted_control_runs(control, position, runs)
                }
                SdtContent::Paragraph(paragraph) => {
                    append_accepted_paragraph_runs(paragraph, Some(position), runs)
                }
                SdtContent::Table(table) => append_accepted_table_runs(table, position, runs),
                SdtContent::Row(row) => append_accepted_row_runs(row, position, runs),
                SdtContent::Cell(cell) => append_accepted_cell_runs(cell, position, runs),
                SdtContent::RawXml(_) => {}
            }
        }
    }
}

fn append_accepted_revision_runs<'a>(
    revision: &'a CT_Revision,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    match revision.kind() {
        RevisionKind::Insertion | RevisionKind::MoveTo => {
            if let Some(paragraph) = revision.content_paragraph() {
                append_accepted_paragraph_runs(paragraph, Some(position), runs);
                return;
            }
            let RevisionContent::Runs(direct) = revision.content() else {
                return;
            };
            for boundary in 0..=direct.len() {
                for (_, nested) in revision
                    .nested_revisions()
                    .iter()
                    .filter(|(at, _)| *at == boundary)
                {
                    append_accepted_revision_runs(nested, position, runs);
                }
                if let Some(run) = direct.get(boundary) {
                    runs.push(AcceptedTocRun {
                        run,
                        run_boundary: position.0,
                        raw_order: position.1,
                        nested_order: 0,
                    });
                }
            }
        }
        RevisionKind::Deletion
        | RevisionKind::MoveFrom
        | RevisionKind::RunPropertyChange
        | RevisionKind::ParagraphPropertyChange
        | RevisionKind::TablePropertyChange
        | RevisionKind::SectionPropertyChange => {}
    }
}

fn append_accepted_table_runs<'a>(
    table: &'a CT_Tbl,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_accepted_control_runs(control, position, runs);
        }
        if let Some(row) = table.rows.get(boundary) {
            append_accepted_row_runs(row, position, runs);
        }
    }
}

fn append_accepted_row_runs<'a>(
    row: &'a CT_Row,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_accepted_control_runs(control, position, runs);
        }
        if let Some(cell) = row.cells.get(boundary) {
            append_accepted_cell_runs(cell, position, runs);
        }
    }
}

fn append_accepted_cell_runs<'a>(
    cell: &'a CT_Tc,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                append_accepted_paragraph_runs(paragraph, Some(position), runs)
            }
            CellContent::Table(table) => append_accepted_table_runs(table, position, runs),
            CellContent::ContentControl(control) => {
                append_accepted_control_runs(control, position, runs)
            }
        }
    }
}

fn field_argument_text(argument: &FieldArgument) -> Option<&str> {
    match argument {
        FieldArgument::Text(value) => Some(value),
        FieldArgument::Nested(_) => None,
    }
}

fn toc_accepts_tc(toc: &TocField, tc: &TcField) -> bool {
    match &toc.entries {
        TocEntrySelection::None => false,
        TocEntrySelection::All => true,
        TocEntrySelection::Identifier(identifier) => tc
            .table_identifier
            .as_deref()
            .is_some_and(|candidate| candidate.eq_ignore_ascii_case(identifier)),
    }
}

fn toc_paragraph_level(
    document: &Document,
    toc: &TocField,
    paragraph: &CT_P,
    effective_heading: Option<Option<u32>>,
) -> Result<Option<u8>> {
    let style_id = paragraph
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref());
    if let Some(style_id) = style_id {
        for (name, level) in &toc.custom_styles {
            let matches_id = style_id.eq_ignore_ascii_case(name);
            let matches_name = document
                .styles
                .get_by_id(style_id)
                .and_then(|style| style.name.as_deref())
                .is_some_and(|style_name| style_name.eq_ignore_ascii_case(name));
            if matches_id || matches_name {
                return Ok(Some(*level));
            }
        }
    }
    let outline = if toc.use_outline_levels {
        match effective_heading {
            Some(Some(level)) => u8::try_from(level)
                .ok()
                .filter(|level| (1..=9).contains(level)),
            Some(None) => return Ok(None),
            None => {
                return Err(Error::Other(
                    "TOC outline source has no qualified current layout facts".into(),
                ));
            }
        }
    } else {
        None
    };
    if let Some(style_id) = style_id
        && let Some((start, end)) = toc.heading_levels
        && let Some(level) = heading_style_level(document, style_id)
        && (start..=end).contains(&level)
    {
        return Ok(Some(level));
    }
    // Apply an explicit heading range to outline-derived entries as well.
    Ok(outline.filter(|level| {
        toc.heading_levels
            .is_none_or(|(start, end)| (start..=end).contains(level))
    }))
}

fn heading_style_level(document: &Document, style_id: &str) -> Option<u8> {
    let candidates = std::iter::once(style_id).chain(
        document
            .styles
            .get_by_id(style_id)
            .and_then(|style| style.name.as_deref()),
    );
    for candidate in candidates {
        let normalized = candidate
            .chars()
            .filter(|character| !character.is_whitespace())
            .collect::<String>()
            .to_ascii_lowercase();
        if let Some(suffix) = normalized.strip_prefix("heading")
            && let Ok(level) = suffix.parse::<u8>()
            && (1..=9).contains(&level)
        {
            return Some(level);
        }
    }
    None
}

#[derive(Debug)]
struct TocParagraphInsertion {
    start: usize,
    end: usize,
    name: String,
    namespaces: BTreeMap<String, String>,
    has_non_namespace_attributes: bool,
    start_tag_end: usize,
    content_start: usize,
    content_end: usize,
}

fn insert_toc_bookmarks_xml(
    xml: &[u8],
    bookmarks: &BTreeMap<usize, TocBookmark>,
    repairs: &[TocBookmarkRepair],
    all_toc_spans: &[DynamicTocSpan],
    toc_spans: &[DynamicTocSpan],
) -> Result<Vec<u8>> {
    let paragraphs = toc_paragraph_insertions(xml)?;
    if bookmarks.keys().any(|index| *index >= paragraphs.len()) {
        return Err(Error::Other(
            "table of contents source paragraph was not found in package XML".to_owned(),
        ));
    }
    let mut insertions = BTreeMap::<usize, Vec<u8>>::new();
    for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
        let bookmark = bookmarks
            .get(&paragraph_index)
            .filter(|bookmark| bookmark.insert);
        if bookmark.is_none() {
            continue;
        }
        let end_boundary = toc_spans
            .iter()
            .any(|span| span.end_paragraph == paragraph_index);
        let fragment_start = toc_spans
            .iter()
            .filter(|span| span.end_paragraph == paragraph_index)
            .map(|span| span.end_run_end)
            .max()
            .unwrap_or(paragraph.content_start);
        let content_end = toc_spans
            .iter()
            .filter(|span| span.begin_paragraph == paragraph_index)
            .map(|span| span.begin_run_start)
            .min()
            .unwrap_or(paragraph.content_end);
        if fragment_start > content_end {
            return Err(Error::Other(
                "table of contents source bookmark has no unowned paragraph range".to_owned(),
            ));
        }
        if !end_boundary && let Some(bookmark) = bookmark {
            let start_insertion = insertions.entry(fragment_start).or_default();
            start_insertion.extend_from_slice(
                format!(
                    "<w:bookmarkStart w:id=\"{}\" w:name=\"{}\" xmlns:w=\"{W_NS}\"/>",
                    bookmark.id,
                    xml_escape_attribute(&bookmark.name)
                )
                .as_bytes(),
            );
        }

        let end_insertion = insertions.entry(content_end).or_default();
        if let Some(bookmark) = bookmark {
            end_insertion.extend_from_slice(
                format!(
                    "<w:bookmarkEnd w:id=\"{}\" xmlns:w=\"{W_NS}\"/>",
                    bookmark.id
                )
                .as_bytes(),
            );
        }
    }
    let mut repair_ends = BTreeMap::<usize, Vec<i32>>::new();
    for repair in repairs.iter().filter(|repair| !repair.replace_start) {
        let span = all_toc_spans.get(repair.toc_index).ok_or_else(|| {
            Error::Other("table of contents bookmark repair owner was not found".to_owned())
        })?;
        repair_ends
            .entry(span.begin_run_start)
            .or_default()
            .push(repair.id);
    }
    for (position, ids) in repair_ends {
        let insertion = insertions.entry(position).or_default();
        for id in ids.into_iter().rev() {
            insertion.extend_from_slice(
                format!("<w:bookmarkEnd w:id=\"{id}\" xmlns:w=\"{W_NS}\"/>").as_bytes(),
            );
        }
    }
    let mut output = xml.to_vec();
    for (position, replacement) in insertions.into_iter().rev() {
        output.splice(position..position, replacement);
    }
    Ok(output)
}

fn toc_paragraph_insertions(xml: &[u8]) -> Result<Vec<TocParagraphInsertion>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut paragraphs = Vec::<TocParagraphInsertion>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let mut is_typed_inline_owner = false;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    let index = paragraphs.len();
                    paragraphs.push(TocParagraphInsertion {
                        start: before,
                        end: after,
                        name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                        namespaces: namespace_bindings.clone(),
                        has_non_namespace_attributes: element.attributes().any(|attribute| {
                            attribute.is_ok_and(|attribute| {
                                attribute.key.as_ref() != b"xmlns"
                                    && !attribute.key.as_ref().starts_with(b"xmlns:")
                            })
                        }),
                        start_tag_end: after,
                        content_start: after,
                        content_end: after,
                    });
                    Some(index)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                mark_typed_sdt_content(&mut elements, &local, typed_block_owner, false);
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner: false,
                    revision_depth: 0,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                let typed_block_owner = classify_typed_block_owner(word, local, &elements);
                if typed_block_owner == Some(TypedBlockOwner::Paragraph) {
                    paragraphs.push(TocParagraphInsertion {
                        start: before,
                        end: after,
                        name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                        namespaces: BTreeMap::new(),
                        has_non_namespace_attributes: false,
                        start_tag_end: before,
                        content_start: before,
                        content_end: before,
                    });
                }
                mark_typed_sdt_content(&mut elements, local, typed_block_owner, false);
                if word
                    && local == b"pPr"
                    && let Some(parent) = elements.last()
                    && parent.is_typed_paragraph
                    && let Some(paragraph) = parent.paragraph
                {
                    paragraphs[paragraph].content_start = after;
                }
            }
            Event::End(element) => {
                let Some(closed) = elements.pop() else {
                    return Err(Error::Other(
                        "table of contents XML has an unmatched end element".to_owned(),
                    ));
                };
                if word
                    && matches_local_name(element.name().as_ref(), b"pPr")
                    && let Some(paragraph) = closed.paragraph
                    && elements
                        .last()
                        .is_some_and(|parent| parent.is_word && parent.local_name == b"p")
                {
                    paragraphs[paragraph].content_start = after;
                }
                if word
                    && matches_local_name(element.name().as_ref(), b"p")
                    && let Some(paragraph) = closed.paragraph
                {
                    paragraphs[paragraph].content_end = before;
                    paragraphs[paragraph].end = after;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(paragraphs)
}

#[derive(Debug)]
struct TocPagePlaceholder {
    toc_index: usize,
    token: String,
    bookmark: String,
}

fn render_toc_entries(
    toc_index: usize,
    toc: &TocField,
    sources: &[TocSource],
    entry_styles: &BTreeMap<u8, TocEntryStyle>,
    fallback_right_tab: i32,
    bookmarks: &BTreeMap<usize, TocBookmark>,
    source_xml: &[u8],
    placeholders: &mut Vec<TocPagePlaceholder>,
) -> Result<Vec<u8>> {
    let mut output = String::new();
    for source in sources {
        let bookmark = bookmarks.get(&source.paragraph_index);
        if source.needs_bookmark && bookmark.is_none() {
            return Err(Error::Other(
                "table of contents source bookmark was not allocated".to_owned(),
            ));
        }
        let entry_style = entry_styles.get(&source.level).ok_or_else(|| {
            Error::Other(format!(
                "table of contents level {} has no entry style",
                source.level
            ))
        })?;
        output.push_str("<w:p><w:pPr><w:pStyle w:val=\"");
        output.push_str(&xml_escape_attribute(&entry_style.style_id));
        output.push_str("\"/>");
        // Word gives the entry of a heading numbered with a tab its own left
        // stop after the number, so the title does not go to the page number
        // stop.
        let number_stop = match source.numbering_prefix.as_ref() {
            Some(prefix) if prefix.suffix == ST_LvlSuffix::Tab => {
                Some(toc_number_tab_stop(&prefix.marker, source.level))
            }
            _ => None,
        };
        let fallback_right = !source.omit_page_number
            && toc.page_number_separator.is_none()
            && !entry_style.has_right_tab;
        if number_stop.is_some() || fallback_right {
            output.push_str("<w:tabs>");
            if let Some(stop) = number_stop {
                output.push_str("<w:tab w:val=\"left\" w:pos=\"");
                output.push_str(&stop.to_string());
                output.push_str("\"/>");
            }
            if fallback_right {
                output.push_str("<w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"");
                output.push_str(&fallback_right_tab.to_string());
                output.push_str("\"/>");
            }
            output.push_str("</w:tabs>");
        }
        output.push_str("</w:pPr>");
        if toc.hyperlink {
            output.push_str("<w:hyperlink w:anchor=\"");
            output.push_str(&xml_escape_attribute(
                &bookmark.expect("required above").name,
            ));
            output.push_str("\">");
        }
        if let Some(prefix) = source.numbering_prefix.as_ref() {
            output.push_str("<w:r><w:t>");
            output.push_str(&xml_escape_text(&prefix.marker));
            output.push_str("</w:t></w:r>");
            match prefix.suffix {
                ST_LvlSuffix::Tab => output.push_str("<w:r><w:tab/></w:r>"),
                ST_LvlSuffix::Space => {
                    output.push_str("<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>")
                }
                ST_LvlSuffix::Nothing => {}
            }
        }
        output.push_str("<w:r><w:t>");
        output.push_str(&xml_escape_text(&source.title));
        output.push_str("</w:t></w:r>");
        if toc.hyperlink {
            output.push_str("</w:hyperlink>");
        }
        if !source.omit_page_number {
            if let Some(separator) = toc.page_number_separator.as_deref() {
                output.push_str("<w:r><w:t xml:space=\"preserve\">");
                output.push_str(&xml_escape_text(separator));
                output.push_str("</w:t></w:r>");
            } else {
                output.push_str("<w:r><w:tab/></w:r>");
            }
            if let Some(sequence) = source.sequence_prefix.as_deref() {
                output.push_str("<w:r><w:t>");
                output.push_str(&xml_escape_text(sequence));
                output.push_str("</w:t></w:r><w:r><w:t>");
                output.push_str(&xml_escape_text(
                    toc.entry_page_separator.as_deref().unwrap_or("-"),
                ));
                output.push_str("</w:t></w:r>");
            }
            let bookmark = &bookmark.expect("required above").name;
            let token = collision_safe_toc_page_token(source_xml, output.as_bytes(), placeholders);
            output.push_str("<w:fldSimple w:instr=\" PAGEREF ");
            output.push_str(&xml_escape_attribute(bookmark));
            output.push_str(" \\h \" w:dirty=\"0\"><w:r><w:t>");
            output.push_str(&token);
            output.push_str("</w:t></w:r></w:fldSimple>");
            placeholders.push(TocPagePlaceholder {
                toc_index,
                token,
                bookmark: bookmark.clone(),
            });
        }
        output.push_str("</w:p>");
    }
    Ok(output.into_bytes())
}

/// Advance widths of Aptos, Word's default font, for the printable ASCII
/// characters from space to `~`, in its 2048 units per em.
///
/// Word measures a TOC entry's number in this font, so these values are
/// facts about Word's output rather than about the document. They are the
/// font's own advances: Word writes 2509 twips for "WWWWWWWWWW1", which is
/// exactly their sum at 12 points plus 240.
const APTOS_ADVANCES: [u16; 95] = [
    416, 600, 757, 1100, 1094, 1692, 1317, 431, 600, 600, 936, 1094, 585, 697, 585, 695, 1094,
    1094, 1094, 1094, 1094, 1094, 1094, 1094, 1094, 1094, 585, 585, 1094, 1094, 1094, 1027, 1832,
    1207, 1237, 1418, 1405, 1139, 1074, 1451, 1448, 533, 678, 1164, 1023, 1618, 1446, 1500, 1182,
    1500, 1241, 1159, 980, 1395, 1199, 1827, 1132, 1105, 1056, 603, 695, 603, 1094, 942, 1130,
    1088, 1149, 1075, 1148, 1079, 617, 992, 1129, 489, 489, 997, 533, 1747, 1129, 1130, 1149, 1148,
    685, 995, 662, 1144, 926, 1476, 906, 926, 897, 603, 553, 603, 1094,
];

/// The left stop Word writes after the number of a TOC entry, in twips.
///
/// Word measures the number in its own default font, 12 pt Aptos in current
/// releases, whatever the document's fonts, and puts the stop on the first
/// 12 pt boundary at least 12 pt past the number, counting 12 pt of indent
/// for each level below the first. A character outside the table counts as
/// one em.
fn toc_number_tab_stop(marker: &str, level: u8) -> i32 {
    const STEP: u64 = 240;
    let units = marker.chars().fold(0u64, |sum, ch| {
        let advance = u32::from(ch)
            .checked_sub(0x20)
            .and_then(|index| APTOS_ADVANCES.get(index as usize))
            .map_or(2048, |advance| u32::from(*advance));
        sum.saturating_add(u64::from(advance))
    });
    // The indent is `level - 1` steps and the stop is one step past the
    // number, so the two cancel to `level` plus the rounded-up width.
    let steps = u64::from(level)
        .saturating_add(units.saturating_add(2047) / 2048)
        .min((i32::MAX as u64) / STEP);
    (steps as i32) * STEP as i32
}

fn collision_safe_toc_page_token(
    source_xml: &[u8],
    generated: &[u8],
    placeholders: &[TocPagePlaceholder],
) -> String {
    let index = placeholders.len();
    for nonce in 0u64.. {
        let token = format!("__RDOCX_TOC_PAGE_{index}_{nonce}__");
        if find_bytes(source_xml, token.as_bytes()).is_none()
            && find_bytes(generated, token.as_bytes()).is_none()
            && placeholders
                .iter()
                .all(|placeholder| placeholder.token != token)
        {
            return token;
        }
    }
    unreachable!("u64 placeholder nonce space is finite but non-empty")
}

fn xml_escape_text(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn xml_escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn end_boundary_bookmark_starts(
    toc_index: usize,
    span: &DynamicTocSpan,
    spans: &[DynamicTocSpan],
    fields: &[Option<TocField>],
    bookmarks: &BTreeMap<usize, TocBookmark>,
    repairs: &[TocBookmarkRepair],
) -> Vec<u8> {
    let mut output = Vec::new();
    for repair in repairs
        .iter()
        .filter(|repair| repair.toc_index == toc_index && repair.replace_start)
    {
        output.extend_from_slice(
            format!(
                "<w:bookmarkStart w:id=\"{}\" w:name=\"{}\" xmlns:w=\"{W_NS}\"/>",
                repair.id,
                xml_escape_attribute(&repair.name)
            )
            .as_bytes(),
        );
    }
    let owns_last_end_boundary = fields.get(toc_index).is_some_and(Option::is_some)
        && spans
            .iter()
            .zip(fields)
            .filter(|(candidate, field)| {
                field.is_some() && candidate.end_paragraph == span.end_paragraph
            })
            .all(|(candidate, _)| candidate.end_run_end <= span.end_run_end);
    if owns_last_end_boundary
        && let Some(bookmark) = bookmarks
            .get(&span.end_paragraph)
            .filter(|bookmark| bookmark.insert)
    {
        output.extend_from_slice(
            format!(
                "<w:bookmarkStart w:id=\"{}\" w:name=\"{}\" xmlns:w=\"{W_NS}\"/>",
                bookmark.id,
                xml_escape_attribute(&bookmark.name)
            )
            .as_bytes(),
        );
    }
    output
}

fn dynamic_toc_replacement(
    xml: &[u8],
    span: &DynamicTocSpan,
    generated: &[u8],
    end_boundary_bookmark_starts: &[u8],
) -> Result<Vec<u8>> {
    if span.result_start > span.end_paragraph_start
        || span.end_paragraph_start > span.end_paragraph_content_start
        || span.end_paragraph_content_start > span.result_end
        || span.result_end > xml.len()
    {
        return Err(Error::Other(
            "table of contents result paragraph boundaries are invalid".to_owned(),
        ));
    }
    let mut prior_wrapper_end = span.end_paragraph_content_start;
    for &(start, end) in &span.end_wrapper_prefixes {
        if start < prior_wrapper_end || start >= end || end > span.result_end {
            return Err(Error::Other(
                "table of contents end-marker wrappers are invalid".to_owned(),
            ));
        }
        prior_wrapper_end = end;
    }
    let mut replacement = Vec::new();
    append_toc_wrapper_closures(&mut replacement, &span.separator_wrapper_names);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(span.start_paragraph_name.as_bytes());
    replacement.push(b'>');
    replacement.extend_from_slice(generated);
    replacement.extend_from_slice(&xml[span.end_paragraph_start..span.end_paragraph_content_start]);
    replacement.extend_from_slice(end_boundary_bookmark_starts);
    for &(start, end) in &span.end_wrapper_prefixes {
        replacement.extend_from_slice(&xml[start..end]);
    }
    Ok(replacement)
}

/// Reuse eligible entry paragraph roots while replacing only owned entry facts.
fn entry_boundary_toc_edits(
    xml: &[u8],
    span: &DynamicTocSpan,
    paragraphs: &[TocParagraphInsertion],
    generated: &[u8],
    entry_style_ids: &BTreeSet<String>,
    end_boundary_bookmark_starts: &[u8],
) -> Result<Option<Vec<FieldSourceEdit>>> {
    if !span.separator_wrapper_names.is_empty()
        || !span.end_wrapper_prefixes.is_empty()
        || !end_boundary_bookmark_starts.is_empty()
    {
        return Ok(None);
    }
    let begin = &paragraphs[span.begin_paragraph];
    let end = &paragraphs[span.end_paragraph];
    let single = span.begin_paragraph == span.end_paragraph;
    let begin_properties = &xml[begin.start_tag_end..begin.content_start];
    let end_properties = &xml[end.start_tag_end..end.content_start];
    let begin_scan = toc_fragment_scan(begin_properties, &begin.namespaces)?;
    let end_scan = toc_fragment_scan(end_properties, &end.namespaces)?;
    let begin_content = toc_fragment_scan(
        &xml[span.result_start..if single {
            span.result_end
        } else {
            begin.content_end
        }],
        &begin.namespaces,
    )?;
    let after_end = toc_fragment_scan(&xml[span.end_run_end..end.content_end], &end.namespaces)?;
    let is_entry = |scan: &TocFragmentScan, has_run: bool| {
        !scan.blocks_replacement
            && has_run
            && scan
                .style
                .as_ref()
                .is_some_and(|(id, _)| entry_style_ids.contains(id))
    };
    let merge_begin = begin.content_start == span.begin_run_start
        && (is_entry(&begin_scan, begin_content.has_run)
            || (single && !begin_content.has_run && !begin_scan.blocks_replacement));
    let merge_end = !after_end.has_run
        && is_entry(
            &end_scan,
            toc_fragment_scan(&xml[end.content_start..span.result_end], &end.namespaces)?.has_run,
        );
    let mut entries = Vec::new();
    let mut offset = 0;
    for close in byte_match_offsets(generated, b"</w:p>") {
        let next = close + b"</w:p>".len();
        entries.push(&generated[offset..next]);
        offset = next;
    }
    if offset != generated.len() {
        return Ok(None);
    }
    if single && (!merge_begin || after_end.has_run) {
        return Ok(None);
    }
    if !single && !merge_begin && !merge_end {
        return Ok(None);
    }
    if entries.is_empty() {
        let mut edits = Vec::new();
        if merge_begin {
            edits.push(FieldSourceEdit {
                start: begin.start_tag_end,
                end: begin.content_start,
                replacement: toc_overlay_entry_properties(begin_properties, &begin_scan, None)?,
            });
        }
        let replacement = if single {
            Vec::new()
        } else {
            // Zero entries retain the balanced two-root topology. Remove only
            // the entry style on an eligible owner, preserving its other facts.
            let mut replacement = format!("</{}>", begin.name).into_bytes();
            replacement.extend_from_slice(&xml[end.start..end.start_tag_end]);
            if merge_end {
                replacement.extend_from_slice(&toc_overlay_entry_properties(
                    end_properties,
                    &end_scan,
                    None,
                )?);
            } else {
                replacement.extend_from_slice(end_properties);
            }
            replacement
        };
        edits.push(FieldSourceEdit {
            start: span.result_start,
            end: span.result_end,
            replacement,
        });
        return Ok(Some(edits));
    }
    // Combining two roots cannot silently discard identity or foreign attributes
    // on the removed end root. Keep that root in the generic topology instead.
    let keep_end_root = entries.len() == 1
        && merge_begin
        && merge_end
        && !single
        && (end.has_non_namespace_attributes || end_scan.has_opaque_properties);
    let merge_end = merge_end && !keep_end_root;
    let generated_bindings = BTreeMap::from([("w".to_owned(), W_NS.to_owned())]);
    let mut edits = Vec::new();
    let mut replacement = Vec::new();
    if !merge_begin {
        replacement.extend_from_slice(format!("</{}>", begin.name).as_bytes());
    }
    for (index, entry) in entries.iter().enumerate() {
        let properties_end = find_bytes(entry, b"</w:pPr>").ok_or_else(|| {
            Error::Other("generated TOC entry has no paragraph properties".to_owned())
        })? + b"</w:pPr>".len();
        let generated_properties = &entry[b"<w:p>".len()..properties_end];
        let first_reused = index == 0 && merge_begin;
        let last_reused = !single && index + 1 == entries.len() && merge_end;
        if first_reused {
            edits.push(FieldSourceEdit {
                start: begin.start_tag_end,
                end: begin.content_start,
                replacement: toc_overlay_entry_properties(
                    begin_properties,
                    &begin_scan,
                    Some(generated_properties),
                )?,
            });
        } else if last_reused {
            replacement.extend_from_slice(&xml[end.start..end.start_tag_end]);
            replacement.extend_from_slice(&toc_overlay_entry_properties(
                end_properties,
                &end_scan,
                Some(generated_properties),
            )?);
        } else {
            replacement.extend_from_slice(format!("<w:p xmlns:w=\"{W_NS}\">").as_bytes());
            replacement.extend_from_slice(generated_properties);
        }
        replacement.extend_from_slice(&toc_close_fragment_namespaces(
            &entry[properties_end..entry.len() - b"</w:p>".len()],
            &generated_bindings,
        )?);
        if index + 1 < entries.len() || (!single && !last_reused) {
            replacement.extend_from_slice(
                if first_reused {
                    format!("</{}>", begin.name).into_bytes()
                } else {
                    b"</w:p>".to_vec()
                }
                .as_slice(),
            );
        }
    }
    if !single && !merge_end {
        replacement.extend_from_slice(&xml[end.start..end.content_start]);
    }
    edits.push(FieldSourceEdit {
        start: span.result_start,
        end: span.result_end,
        replacement,
    });
    if single && entries.len() > 1 {
        // The final generated entry no longer inherits the original paragraph
        // scope. Retained end content must carry that scope on every moved root.
        edits.push(FieldSourceEdit {
            start: span.result_end,
            end: end.content_end,
            replacement: toc_close_fragment_namespaces(
                &xml[span.result_end..end.content_end],
                &end.namespaces,
            )?,
        });
        edits.push(FieldSourceEdit {
            start: end.content_end,
            end: end.end,
            replacement: b"</w:p>".to_vec(),
        });
    }
    if !single && merge_begin && merge_end && entries.len() == 1 {
        // The end run and every trailing sibling move out of their original
        // namespace scope. Close each retained root, without rewriting tokens.
        edits.push(FieldSourceEdit {
            start: span.result_end,
            end: end.content_end,
            replacement: toc_close_fragment_namespaces(
                &xml[span.result_end..end.content_end],
                &end.namespaces,
            )?,
        });
        edits.push(FieldSourceEdit {
            start: end.content_end,
            end: end.end,
            replacement: format!("</{}>", begin.name).into_bytes(),
        });
    }
    edits.sort_by_key(|edit| edit.start);
    Ok(Some(edits))
}

/// Resolve a fragment in its actual paragraph scope, retaining exact byte ranges.
struct TocFragmentScan {
    namespaces: BTreeMap<String, String>,
    style: Option<(String, std::ops::Range<usize>)>,
    has_run: bool,
    blocks_replacement: bool,
    has_opaque_properties: bool,
    property_root: Option<(usize, usize, usize, usize, String)>,
    children: Vec<(Option<String>, std::ops::Range<usize>)>,
}

fn toc_fragment_scan(
    fragment: &[u8],
    bindings: &BTreeMap<String, String>,
) -> Result<TocFragmentScan> {
    let mut wrapper = b"<toc-fragment".to_vec();
    for (prefix, uri) in bindings {
        if prefix == "xml" {
            continue;
        }
        wrapper.extend_from_slice(
            if prefix.is_empty() {
                " xmlns".to_owned()
            } else {
                format!(" xmlns:{prefix}")
            }
            .as_bytes(),
        );
        wrapper.extend_from_slice(format!("=\"{}\"", xml_escape_attribute(uri)).as_bytes());
    }
    wrapper.push(b'>');
    let origin = wrapper.len();
    wrapper.extend_from_slice(fragment);
    wrapper.extend_from_slice(b"</toc-fragment>");
    let mut reader = NsReader::from_reader(wrapper.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut scan = TocFragmentScan {
        namespaces: bindings.clone(),
        style: None,
        has_run: false,
        blocks_replacement: false,
        has_opaque_properties: false,
        property_root: None,
        children: Vec::new(),
    };
    let mut child: Option<(Option<String>, usize)> = None;
    let mut owned = HashSet::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid TOC boundary XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let name = element.name();
                let local = local_name(name.as_ref());
                if depth >= 1 && word && local == b"r" {
                    scan.has_run = true;
                }
                if depth == 1 && word && matches!(local, b"pPr" | b"tabs") {
                    if scan.property_root.is_some() {
                        return Err(Error::Other(
                            "duplicate owned TOC paragraph properties".to_owned(),
                        ));
                    }
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| {
                            Error::Other(format!("invalid TOC boundary namespace: {error}"))
                        })?;
                        let key = attribute.key.as_ref();
                        let prefix = if key == b"xmlns" {
                            Some(String::new())
                        } else {
                            key.strip_prefix(b"xmlns:")
                                .map(|prefix| String::from_utf8_lossy(prefix).into_owned())
                        };
                        if let Some(prefix) = prefix {
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    element.decoder(),
                                )
                                .map_err(|error| {
                                    Error::Other(format!("invalid TOC boundary namespace: {error}"))
                                })?
                                .into_owned();
                            if value.is_empty() {
                                scan.namespaces.remove(&prefix);
                            } else {
                                scan.namespaces.insert(prefix, value);
                            }
                        }
                    }
                    let start = before - origin;
                    let opening = after - origin;
                    scan.property_root = Some((
                        start,
                        opening,
                        opening,
                        opening,
                        String::from_utf8_lossy(name.as_ref()).into_owned(),
                    ));
                    if empty {
                        scan.property_root.as_mut().unwrap().2 = start;
                    }
                }
                // A generated stop cannot erase opaque facts attached to an
                // existing stop. Leave this boundary in the generic topology.
                let inside_tab_stop = depth >= 3
                    && child
                        .as_ref()
                        .is_some_and(|(key, _)| key.as_deref() == Some("tabs"));
                if inside_tab_stop && (depth > 3 || !word || local != b"tab") {
                    scan.blocks_replacement = true;
                }
                let inside_style = child
                    .as_ref()
                    .is_some_and(|(key, _)| key.as_deref() == Some("pStyle"));
                if inside_style && depth >= 3 {
                    scan.blocks_replacement = true;
                }
                if scan.property_root.is_some() && depth >= 1 {
                    if depth >= 2 && !word {
                        scan.has_opaque_properties = true;
                    }
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| {
                            Error::Other(format!("invalid TOC property attribute: {error}"))
                        })?;
                        let key = attribute.key.as_ref();
                        let declaration = key == b"xmlns" || key.starts_with(b"xmlns:");
                        // Namespace closure on a generated Word fact is not an
                        // opaque producer declaration. Its exact binding is
                        // already required by the qualified owner itself.
                        if key == b"xmlns:w"
                            && word
                            && ((depth == 2 && name.as_ref() == b"w:pStyle")
                                || (inside_tab_stop && depth == 3 && name.as_ref() == b"w:tab"))
                            && attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    element.decoder(),
                                )
                                .map_err(|error| {
                                    Error::Other(format!("invalid TOC property namespace: {error}"))
                                })?
                                == W_NS
                        {
                            continue;
                        }
                        let (attribute_namespace, attribute_local) =
                            reader.resolver().resolve_attribute(attribute.key);
                        // Only these attributes are generated entry facts. A
                        // recognized property name does not own its other data.
                        let owned_fact = !declaration
                            && word
                            && namespace_is_word(&attribute_namespace)
                            && ((depth == 2
                                && local == b"pStyle"
                                && attribute_local.as_ref() == b"val")
                                || (inside_tab_stop
                                    && depth == 3
                                    && local == b"tab"
                                    && matches!(
                                        attribute_local.as_ref(),
                                        b"val" | b"leader" | b"pos"
                                    )));
                        if !owned_fact {
                            scan.has_opaque_properties = true;
                            if inside_tab_stop || (depth == 2 && word && local == b"pStyle") {
                                // These owners are replaced or removed, including
                                // zero-entry style clearing. Keep opaque owners.
                                scan.blocks_replacement = true;
                            }
                        }
                    }
                }
                if depth == 2 && scan.property_root.is_some() {
                    let key = word.then(|| String::from_utf8_lossy(local).into_owned());
                    if word
                        && scan
                            .property_root
                            .as_ref()
                            .is_some_and(|root| !root.4.ends_with(":tabs") && root.4 != "tabs")
                        && !matches!(local, b"pStyle" | b"tabs")
                    {
                        scan.has_opaque_properties = true;
                    }
                    if let Some(ref key) = key {
                        if key != "tab" && !owned.insert(key.clone()) {
                            return Err(Error::Other(format!(
                                "duplicate owned TOC paragraph property {key}"
                            )));
                        }
                        if matches!(local, b"sectPr" | b"numPr" | b"pPrChange") {
                            scan.blocks_replacement = true;
                        }
                    }
                    let range = before - origin..after - origin;
                    if word && local == b"pStyle" {
                        let mut value = None;
                        for attribute in element.attributes() {
                            let attribute = attribute.map_err(|error| {
                                Error::Other(format!("invalid TOC property attribute: {error}"))
                            })?;
                            let (namespace, local) =
                                reader.resolver().resolve_attribute(attribute.key);
                            if namespace_is_word(&namespace) && local.as_ref() == b"val" {
                                if value.is_some() {
                                    return Err(Error::Other(
                                        "duplicate owned TOC style value".to_owned(),
                                    ));
                                }
                                value = Some(
                                    attribute
                                        .decoded_and_normalized_value(
                                            XmlVersion::Implicit1_0,
                                            element.decoder(),
                                        )
                                        .map_err(|error| {
                                            Error::Other(format!(
                                                "invalid TOC style value: {error}"
                                            ))
                                        })?
                                        .into_owned(),
                                );
                            }
                        }
                        let value = value.ok_or_else(|| {
                            Error::Other("TOC entry style has no owned value".to_owned())
                        })?;
                        scan.style = Some((value, range.clone()));
                    }
                    if empty {
                        scan.children.push((key, range));
                    } else {
                        child = Some((key, range.start));
                    }
                }
                if word && depth > 2 && matches!(local, b"ins" | b"del" | b"moveFrom" | b"moveTo") {
                    scan.blocks_replacement = true;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Other("unbalanced TOC boundary XML".to_owned()))?;
                if depth == 2
                    && let Some((key, start)) = child.take()
                {
                    let range = start..after - origin;
                    if key.as_deref() == Some("pStyle")
                        && let Some((_, old)) = scan.style.as_mut()
                    {
                        *old = range.clone();
                    }
                    scan.children.push((key, range));
                }
                if depth == 1
                    && let Some(root) = scan.property_root.as_mut()
                {
                    // Only the pPr root is active at depth two.
                    if root.3 == root.1 {
                        root.2 = before - origin;
                        root.3 = after - origin;
                    }
                }
            }
            Event::Comment(_)
            | Event::PI(_)
            | Event::CData(_)
            | Event::Text(_)
            | Event::GeneralRef(_) => {
                if depth >= 2
                    && scan
                        .property_root
                        .as_ref()
                        .is_some_and(|root| root.3 == root.1)
                {
                    // Ordinary XML indentation does not protect an otherwise
                    // mergeable root. All other body tokens are imported facts.
                    let opaque = match &event {
                        Event::Text(text) => !text
                            .iter()
                            .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n')),
                        _ => true,
                    };
                    scan.has_opaque_properties |= opaque;
                }
                if (depth >= 4
                    && child
                        .as_ref()
                        .is_some_and(|(key, _)| key.as_deref() == Some("tabs")))
                    || (depth >= 3
                        && child
                            .as_ref()
                            .is_some_and(|(key, _)| key.as_deref() == Some("pStyle")))
                {
                    // These tokens belong to the imported stop, not the generated
                    // entry facts. Replacing the stop would discard their bytes.
                    scan.blocks_replacement = true;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(scan)
}

/// Overlay generated style and tab facts. Retain all other source particles.
fn toc_overlay_entry_properties(
    original: &[u8],
    scan: &TocFragmentScan,
    generated: Option<&[u8]>,
) -> Result<Vec<u8>> {
    let canonical = BTreeMap::from([("w".to_owned(), W_NS.to_owned())]);
    let mut replacements = Vec::new();
    if let Some(generated) = generated {
        let new_scan = toc_fragment_scan(generated, &canonical)?;
        for (name, range) in new_scan.children {
            let name = name.unwrap();
            let replacement = if name == "pStyle"
                && let Some((old_id, old_range)) = &scan.style
            {
                let generated_scan = toc_fragment_scan(generated, &canonical)?;
                let new_id = &generated_scan
                    .style
                    .as_ref()
                    .ok_or_else(|| Error::Other("generated TOC style is absent".to_owned()))?
                    .0;
                if old_id == new_id {
                    original[old_range.clone()].to_vec()
                } else {
                    toc_retarget_style(&original[old_range.clone()], new_id, &scan.namespaces)?
                }
            } else {
                toc_close_fragment_namespaces(&generated[range], &canonical)?
            };
            replacements.push((name, replacement));
        }
    }
    let Some((start, opening_end, closing_start, end, ref name)) = scan.property_root else {
        return generated.map_or_else(
            || Ok(original.to_vec()),
            |properties| toc_close_fragment_namespaces(properties, &canonical),
        );
    };
    let mut output = original[..start].to_vec();
    let empty = opening_end == end && closing_start == start;
    if empty {
        output.extend_from_slice(&original[start..opening_end - 2]);
        output.push(b'>');
    } else {
        output.extend_from_slice(&original[start..opening_end]);
    }
    // pStyle comes first in CT_PPr. Insert it before retaining source particles.
    if let Some((_, replacement)) = replacements.iter().find(|(key, _)| key == "pStyle") {
        output.extend_from_slice(replacement);
    }
    let mut cursor = opening_end;
    let mut tabs_written = false;
    for (key, range) in &scan.children {
        output.extend_from_slice(&original[cursor..range.start]);
        if key.as_deref() == Some("pStyle") {
            // Remove the complete old style, including a started element's end.
        } else if key.as_deref() == Some("tabs") {
            if let Some((_, replacement)) = replacements.iter().find(|(key, _)| key == "tabs") {
                output.extend_from_slice(&toc_overlay_tab_particles(
                    &original[range.clone()],
                    replacement,
                    &scan.namespaces,
                )?);
            } else {
                output.extend_from_slice(&original[range.clone()]);
            }
            tabs_written = true;
        } else {
            // Tabs precede spacing and later CT_PPr properties. Unknown particles
            // keep their original order and exact source bytes.
            if !tabs_written
                && matches!(
                    key.as_deref(),
                    Some(
                        "suppressAutoHyphens"
                            | "kinsoku"
                            | "wordWrap"
                            | "overflowPunct"
                            | "topLinePunct"
                            | "autoSpaceDE"
                            | "autoSpaceDN"
                            | "bidi"
                            | "adjustRightInd"
                            | "snapToGrid"
                            | "spacing"
                            | "ind"
                            | "contextualSpacing"
                            | "mirrorIndents"
                            | "suppressOverlap"
                            | "jc"
                            | "textDirection"
                            | "textAlignment"
                            | "textboxTightWrap"
                            | "outlineLvl"
                            | "divId"
                            | "cnfStyle"
                            | "rPr"
                            | "sectPr"
                            | "pPrChange"
                    )
                )
            {
                if let Some((_, replacement)) = replacements.iter().find(|(key, _)| key == "tabs") {
                    output.extend_from_slice(replacement);
                }
                tabs_written = true;
            }
            output.extend_from_slice(&original[range.clone()]);
        }
        cursor = range.end;
    }
    if !tabs_written
        && let Some((_, replacement)) = replacements.iter().find(|(key, _)| key == "tabs")
    {
        output.extend_from_slice(replacement);
    }
    if !empty {
        output.extend_from_slice(&original[cursor..closing_start]);
    }
    if empty {
        output.extend_from_slice(format!("</{name}>").as_bytes());
    } else {
        output.extend_from_slice(&original[closing_start..end]);
    }
    output.extend_from_slice(&original[end..]);
    Ok(output)
}

/// Change only the qualified style value, retaining opaque attributes and body.
fn toc_retarget_style(
    original: &[u8],
    value: &str,
    bindings: &BTreeMap<String, String>,
) -> Result<Vec<u8>> {
    let closed = xml_fragment_with_namespaces(original, bindings, "TOC style")?;
    let mut reader = NsReader::from_reader(closed.as_slice());
    let mut buffer = Vec::new();
    let (_, event) = reader
        .read_resolved_event_into(&mut buffer)
        .map_err(|error| Error::Other(format!("invalid TOC style: {error}")))?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Err(Error::Other("TOC style has no root".to_owned())),
    };
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Other(format!("invalid TOC style attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_is_word(&namespace) && local.as_ref() == b"val" {
            let (_, start, end) = attribute_source_span(original, attribute.key.as_ref())
                .ok_or_else(|| Error::Other("TOC style source value is unavailable".to_owned()))?;
            let mut output = original[..start].to_vec();
            output.extend_from_slice(xml_escape_attribute(value).as_bytes());
            output.extend_from_slice(&original[end..]);
            return Ok(output);
        }
    }
    Err(Error::Other("TOC style has no qualified value".to_owned()))
}

/// Replace owned tab stops while keeping the source tabs root and opaque particles.
fn toc_overlay_tab_particles(
    original: &[u8],
    generated: &[u8],
    bindings: &BTreeMap<String, String>,
) -> Result<Vec<u8>> {
    let old = toc_fragment_scan(original, bindings)?;
    let canonical = BTreeMap::from([("w".to_owned(), W_NS.to_owned())]);
    let new = toc_fragment_scan(generated, &canonical)?;
    let (start, opening, closing, end, name) = old
        .property_root
        .as_ref()
        .ok_or_else(|| Error::Other("TOC tabs have no owned root".to_owned()))?;
    let (new_start, new_opening, new_closing, _, _) = new
        .property_root
        .as_ref()
        .ok_or_else(|| Error::Other("generated TOC tabs have no owned root".to_owned()))?;
    let _ = (start, new_start);
    let replacement =
        toc_close_fragment_namespaces(&generated[*new_opening..*new_closing], &canonical)?;
    let empty = *opening == *end && *closing == *start;
    let mut output = Vec::new();
    if empty {
        output.extend_from_slice(&original[..*opening - 2]);
        output.push(b'>');
    } else {
        output.extend_from_slice(&original[..*opening]);
    }
    let mut cursor = *opening;
    let mut written = false;
    for (key, range) in old.children {
        output.extend_from_slice(&original[cursor..range.start]);
        if key.as_deref() == Some("tab") {
            if !written {
                output.extend_from_slice(&replacement);
                written = true;
            }
        } else {
            output.extend_from_slice(&original[range.clone()]);
        }
        cursor = range.end;
    }
    if !written {
        output.extend_from_slice(&replacement);
    }
    if !empty {
        output.extend_from_slice(&original[cursor..*closing]);
        output.extend_from_slice(&original[*closing..*end]);
    } else {
        output.extend_from_slice(format!("</{name}>").as_bytes());
    }
    output.extend_from_slice(&original[*end..]);
    Ok(output)
}

/// Close every top-level fragment root using its original inherited namespace map.
fn toc_close_fragment_namespaces(
    fragment: &[u8],
    bindings: &BTreeMap<String, String>,
) -> Result<Vec<u8>> {
    let mut reader = quick_xml::Reader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut cursor = 0usize;
    let mut root = 0usize;
    let mut output = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid retained TOC fragment: {error}")))?
        {
            Event::Start(_) => {
                if depth == 0 {
                    root = before;
                }
                depth += 1;
            }
            Event::Empty(_) if depth == 0 => {
                let end = reader.buffer_position() as usize;
                output.extend_from_slice(&fragment[cursor..before]);
                output.extend_from_slice(&xml_fragment_with_namespaces(
                    &fragment[before..end],
                    bindings,
                    "TOC fragment",
                )?);
                cursor = end;
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Other("unbalanced retained TOC fragment".to_owned()))?;
                if depth == 0 {
                    let end = reader.buffer_position() as usize;
                    output.extend_from_slice(&fragment[cursor..root]);
                    output.extend_from_slice(&xml_fragment_with_namespaces(
                        &fragment[root..end],
                        bindings,
                        "TOC fragment",
                    )?);
                    cursor = end;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Other("unclosed retained TOC fragment".to_owned()));
    }
    output.extend_from_slice(&fragment[cursor..]);
    Ok(output)
}

/// Every style identifier a TOC entry of any level may carry: the producer
/// styles named `toc 1` to `toc 9`, the canonical `TOC1` to `TOC9`, and the
/// styles this rebuild resolved.
fn toc_entry_style_ids(
    styles: &CT_Styles,
    resolved: &BTreeMap<u8, TocEntryStyle>,
) -> BTreeSet<String> {
    let mut ids = (1..=9)
        .map(|level| format!("TOC{level}"))
        .collect::<BTreeSet<_>>();
    for style in &styles.styles {
        let is_entry_name = style.name.as_deref().is_some_and(|name| {
            (1..=9).any(|level| name.trim().eq_ignore_ascii_case(&format!("toc {level}")))
        });
        if style.style_type == StyleType::Paragraph && !style.style_id.is_empty() && is_entry_name {
            ids.insert(style.style_id.clone());
        }
    }
    ids.extend(resolved.values().map(|style| style.style_id.clone()));
    ids
}

fn relocate_end_boundary_bookmark_starts(
    mut xml: Vec<u8>,
    spans: &[DynamicTocSpan],
    marker_starts: &[Vec<u8>],
) -> Result<Vec<u8>> {
    let mut edits = Vec::new();
    for (toc_index, span) in spans.iter().enumerate() {
        let markers = marker_starts.get(toc_index).ok_or_else(|| {
            Error::Other("table of contents end bookmark owner was not retained".to_owned())
        })?;
        if markers.is_empty() {
            continue;
        }
        let search_end = span
            .end_wrapper_prefixes
            .first()
            .map(|(start, _)| *start)
            .unwrap_or(span.result_end);
        if span.end_paragraph_content_start > search_end || search_end > span.end_run_end {
            return Err(Error::Other(
                "table of contents end bookmark boundaries are invalid".to_owned(),
            ));
        }
        let matches =
            byte_match_offsets(&xml[span.end_paragraph_content_start..search_end], markers);
        if matches.len() != 1 {
            return Err(Error::Other(
                "table of contents end bookmark start was not uniquely staged".to_owned(),
            ));
        }
        let start = span.end_paragraph_content_start + matches[0];
        let end = start + markers.len();
        if end > span.end_run_end {
            return Err(Error::Other(
                "table of contents end bookmark start crosses its field boundary".to_owned(),
            ));
        }
        let mut replacement = xml[end..span.end_run_end].to_vec();
        replacement.extend_from_slice(markers);
        edits.push(FieldSourceEdit {
            start,
            end: span.end_run_end,
            replacement,
        });
    }
    edits.sort_by_key(|edit| edit.start);
    if edits.windows(2).any(|pair| pair[0].end > pair[1].start) {
        return Err(Error::Other(
            "table of contents end bookmark repairs overlap".to_owned(),
        ));
    }
    for edit in edits.into_iter().rev() {
        xml.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(xml)
}

fn append_toc_wrapper_closures(output: &mut Vec<u8>, wrappers: &[String]) {
    for name in wrappers.iter().rev() {
        output.extend_from_slice(b"</");
        output.extend_from_slice(name.as_bytes());
        output.push(b'>');
    }
}

fn reopen_staged_document(document: Document) -> Result<Document> {
    document.reopen_prepared_staged()
}

fn deterministic_toc_page_values(document: &Document) -> Result<HashMap<String, String>> {
    let layout = document.layout_deterministic()?;
    let mut output = HashMap::new();
    let mut invalid_target = None;
    for page in &layout.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element
                && let Some(oxml_layout::FieldKind::TargetPage(target)) = run.field_kind
                && let Some(name) = layout.page_reference_name(target)
            {
                let valid_page = run
                    .text
                    .parse::<usize>()
                    .is_ok_and(|value| (1..=layout.layout.pages.len()).contains(&value));
                if !valid_page
                    || output
                        .get(name)
                        .is_some_and(|existing| existing != &run.text)
                {
                    invalid_target.get_or_insert_with(|| name.to_owned());
                } else {
                    output
                        .entry(name.to_owned())
                        .or_insert_with(|| run.text.clone());
                }
            }
        });
    }
    if let Some(target) = invalid_target {
        return Err(Error::Other(format!(
            "table of contents page target {target} was not resolved"
        )));
    }
    Ok(output)
}

fn byte_match_offsets(input: &[u8], needle: &[u8]) -> Vec<usize> {
    if needle.is_empty() {
        return Vec::new();
    }
    let mut output = Vec::new();
    let mut start = 0usize;
    while let Some(relative) = find_bytes(&input[start..], needle) {
        let found = start + relative;
        output.push(found);
        start = found + needle.len();
    }
    output
}

fn non_body_story_parts(document: &Document) -> Result<Vec<(String, Vec<u8>)>> {
    let mut parts = merge_referenced_header_footer_parts(document)?
        .into_iter()
        .chain(
            relationship_parts(document, rel_types::FOOTNOTES)
                .into_iter()
                .map(|(name, xml)| (format!("footnotes:{name}"), xml)),
        )
        .chain(
            relationship_parts(document, rel_types::ENDNOTES)
                .into_iter()
                .map(|(name, xml)| (format!("endnotes:{name}"), xml)),
        )
        .collect::<Vec<_>>();
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    Ok(parts)
}

fn merge_referenced_header_footer_parts(document: &Document) -> Result<Vec<(String, Vec<u8>)>> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Ok(Vec::new());
    };
    let xml = document.document.to_xml()?;
    let mut reader = NsReader::from_reader(xml.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut references = Vec::<(String, bool)>::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid main document XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if namespace_is_word(&reader.resolver().resolve_element(element.name()).0) =>
            {
                let name = element.name();
                let local = local_name(name.as_ref());
                let is_header = if local == b"headerReference" {
                    true
                } else if local == b"footerReference" {
                    false
                } else {
                    buffer.clear();
                    continue;
                };
                if let Some((_, rel_id)) = resolved_element_attribute(
                    &element,
                    reader.resolver(),
                    b"id",
                    AttributeNamespace::Relationship,
                )? {
                    references.push((rel_id, is_header));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut seen = HashSet::new();
    let mut parts = Vec::new();
    for (rel_id, is_header) in references {
        let Some(relationship) = relationships.get_by_id(&rel_id) else {
            continue;
        };
        let relationship_type = if is_header {
            rel_types::HEADER
        } else {
            rel_types::FOOTER
        };
        if relationship.rel_type != relationship_type
            || !crate::document::relationship_is_internal(relationship)
        {
            continue;
        }
        let part_name =
            OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
        if seen.insert(part_name.clone())
            && let Some(xml) = document.package.get_part(&part_name)
        {
            let story = if is_header { "header" } else { "footer" };
            parts.push((format!("{story}:{part_name}"), xml.to_vec()));
        }
    }
    Ok(parts)
}

fn reject_varying_non_body_merge_fields(
    document: &Document,
    records: &[BTreeMap<String, String>],
) -> Result<()> {
    if records.is_empty() {
        return Ok(());
    }
    let mut names = HashSet::new();
    for (_, xml) in non_body_story_parts(document)? {
        collect_raw_merge_field_names(&xml, &mut names)?;
    }
    for name in names {
        let expected = records[0].get(&name).map(String::as_str).unwrap_or("");
        if records
            .iter()
            .skip(1)
            .any(|record| record.get(&name).map(String::as_str).unwrap_or("") != expected)
        {
            return Err(Error::Other(
                "sectioned mail merge cannot vary fields in headers, footers, footnotes, or endnotes"
                    .to_owned(),
            ));
        }
    }
    Ok(())
}

struct ComplexInstruction {
    text: String,
    collecting: bool,
}

fn collect_raw_merge_field_names(xml: &[u8], names: &mut HashSet<String>) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut complex = Vec::<ComplexInstruction>::new();
    let mut in_instruction_text = false;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package story XML: {error}")))?;
        let word = match &event {
            Event::Start(element) | Event::Empty(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            Event::End(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            _ => false,
        };
        match event {
            Event::Start(element) if word => {
                if collect_raw_field_element(&element, reader.resolver(), names, &mut complex)? {
                    in_instruction_text = true;
                }
            }
            Event::Empty(element) if word => {
                collect_raw_field_element(&element, reader.resolver(), names, &mut complex)?;
            }
            Event::Text(text) if in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    let decoded = text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::Other(format!("invalid field instruction entity: {error}"))
                    })?;
                    instruction.text.push_str(&unescaped);
                }
            }
            Event::CData(text) if in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    instruction.text.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?);
                }
            }
            Event::End(element)
                if word && matches_local_name(element.name().as_ref(), b"instrText") =>
            {
                in_instruction_text = false;
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_raw_field_element(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    names: &mut HashSet<String>,
    complex: &mut Vec<ComplexInstruction>,
) -> Result<bool> {
    let name = element.name();
    let local = local_name(name.as_ref());
    if local == b"fldSimple" {
        if let Some((_, instruction)) =
            resolved_element_attribute(element, resolver, b"instr", AttributeNamespace::Word)?
        {
            collect_merge_field_name(&instruction, names);
        }
    } else if local == b"fldChar" {
        match resolved_element_attribute(
            element,
            resolver,
            b"fldCharType",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value)
        .as_deref()
        {
            Some("begin") => complex.push(ComplexInstruction {
                text: String::new(),
                collecting: true,
            }),
            Some("separate") => {
                if let Some(instruction) = complex.last_mut() {
                    instruction.collecting = false;
                }
            }
            Some("end") => {
                if let Some(instruction) = complex.pop() {
                    collect_merge_field_name(&instruction.text, names);
                }
            }
            _ => {}
        }
    }
    Ok(local == b"instrText")
}

fn collect_merge_field_name(instruction: &str, names: &mut HashSet<String>) {
    let field = Field::new(instruction, "");
    if field.instruction.name == "MERGEFIELD"
        && let Some(FieldArgument::Text(name)) = field.instruction.arguments.first()
    {
        names.insert(name.clone());
    }
}

#[derive(Default)]
struct BodyIdentityValues {
    bookmark_ids: Vec<String>,
    bookmark_events: Vec<(String, bool)>,
    comment_ids: Vec<String>,
    comment_events: Vec<(String, u8)>,
    content_control_ids: Vec<String>,
    drawing_ids: Vec<String>,
    non_visual_drawing_ids: Vec<String>,
    bookmark_names: Vec<String>,
    reference_names: Vec<String>,
}

/// Merge-local identities that are outside F-249's document allocator scope.
///
/// Bookmark and `wp:docPr` ids are reserved from `DocumentIdentifiers` by
/// `remap_body_identities`. Content-control and non-visual drawing ids have
/// distinct OOXML scopes and remain local to the rich-merge combiner.
struct BodyIdentityState {
    used_content_control_ids: HashSet<u32>,
    used_non_visual_drawing_ids: HashSet<u32>,
    used_names: HashSet<String>,
    next_content_control_id: u32,
    next_non_visual_drawing_id: u32,
    next_name: u32,
}

impl BodyIdentityState {
    fn from_documents(documents: &[Document]) -> Result<Self> {
        let mut state = Self {
            used_content_control_ids: HashSet::new(),
            used_non_visual_drawing_ids: HashSet::new(),
            used_names: HashSet::new(),
            next_content_control_id: 1,
            next_non_visual_drawing_id: 1,
            next_name: 1,
        };
        for document in documents {
            let xml = document.document.to_xml()?;
            let values = body_identity_values(&xml)?;
            state.used_content_control_ids.extend(
                values
                    .content_control_ids
                    .iter()
                    .filter_map(|value| value.parse::<u32>().ok()),
            );
            state.used_non_visual_drawing_ids.extend(
                values
                    .non_visual_drawing_ids
                    .iter()
                    .filter_map(|value| value.parse::<u32>().ok()),
            );
            state.used_names.extend(values.bookmark_names);
            state.used_names.extend(values.reference_names);
            for bytes in document.package.parts.values() {
                if let Ok(values) = body_identity_values(&wrap_fragment_companion(bytes)) {
                    state.used_content_control_ids.extend(
                        values
                            .content_control_ids
                            .iter()
                            .filter_map(|value| value.parse::<u32>().ok()),
                    );
                    state.used_non_visual_drawing_ids.extend(
                        values
                            .non_visual_drawing_ids
                            .iter()
                            .filter_map(|value| value.parse::<u32>().ok()),
                    );
                    state.used_names.extend(values.bookmark_names);
                    state.used_names.extend(values.reference_names);
                }
            }
        }
        Ok(state)
    }

    fn allocate_content_control_id(&mut self) -> Result<String> {
        allocate_merge_local_id(
            &mut self.used_content_control_ids,
            &mut self.next_content_control_id,
            "content-control",
        )
    }

    fn allocate_non_visual_drawing_id(&mut self) -> Result<String> {
        allocate_merge_local_id(
            &mut self.used_non_visual_drawing_ids,
            &mut self.next_non_visual_drawing_id,
            "non-visual drawing",
        )
    }

    fn allocate_name(&mut self) -> Result<String> {
        loop {
            let candidate = format!("MailMerge{}", self.next_name);
            self.next_name = self.next_name.checked_add(1).ok_or_else(|| {
                Error::Other("mail merge exhausted the bookmark name range".to_owned())
            })?;
            if self.used_names.insert(candidate.clone()) {
                return Ok(candidate);
            }
        }
    }
}

fn allocate_merge_local_id(
    occupied: &mut HashSet<u32>,
    cursor: &mut u32,
    category: &str,
) -> Result<String> {
    loop {
        let candidate = *cursor;
        *cursor = cursor.checked_add(1).ok_or_else(|| {
            Error::Other(format!(
                "mail merge exhausted the {category} identity range"
            ))
        })?;
        if occupied.insert(candidate) {
            return Ok(candidate.to_string());
        }
    }
}

#[derive(Default)]
struct BodyIdentityRemap {
    bookmark_ids: BTreeMap<String, String>,
    comment_ids: BTreeMap<String, String>,
    content_control_ids: BTreeMap<String, String>,
    drawing_ids: BTreeMap<String, String>,
    non_visual_drawing_ids: BTreeMap<String, String>,
    bookmark_names: BTreeMap<String, String>,
    /// Remove `w14:paraId` and `w14:textId` from paragraphs and table rows.
    /// A copy must not share them with its source, and Word assigns new ones
    /// to an element that has none.
    drop_paragraph_identities: bool,
}

fn remap_body_identities(
    document: &mut Document,
    identifiers: &mut DocumentIdentifiers,
    state: &mut BodyIdentityState,
) -> Result<BodyIdentityRemap> {
    let (updated, remap) =
        remap_fragment_xml_identities(&document.document.to_xml()?, identifiers, state, true)?;
    document.document = CT_Document::from_xml(&updated)?;
    Ok(remap)
}

fn remap_fragment_xml_identities(
    source: &[u8],
    identifiers: &mut DocumentIdentifiers,
    state: &mut BodyIdentityState,
    drop_paragraph_identities: bool,
) -> Result<(Vec<u8>, BodyIdentityRemap)> {
    // The drawing remap is keyed by value, so a source that repeats a
    // `wp:docPr/@id` is first made unique and every copy gets its own id.
    let xml = uniquify_drawing_ids_in_xml(source)?;
    let values = body_identity_values(&xml)?;
    let mut remap = BodyIdentityRemap {
        drop_paragraph_identities,
        ..Default::default()
    };
    for value in values.bookmark_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.bookmark_ids.entry(value) {
            entry.insert(identifiers.reserve_bookmark_id()?.to_string());
        }
    }
    for value in values.content_control_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) =
            remap.content_control_ids.entry(value)
        {
            entry.insert(state.allocate_content_control_id()?);
        }
    }
    for value in values.drawing_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.drawing_ids.entry(value) {
            entry.insert(identifiers.reserve_drawing_id()?.to_string());
        }
    }
    for value in values.non_visual_drawing_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) =
            remap.non_visual_drawing_ids.entry(value)
        {
            entry.insert(state.allocate_non_visual_drawing_id()?);
        }
    }
    for value in values.bookmark_names {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.bookmark_names.entry(value)
        {
            entry.insert(state.allocate_name()?);
        }
    }
    let updated = patch_body_identity_attributes(&xml, &remap)?;
    Ok((updated, remap))
}

fn body_identity_values(xml: &[u8]) -> Result<BodyIdentityValues> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut body_depth = None;
    let mut sdt_properties_depth = None;
    let mut values = BodyIdentityValues::default();
    let mut complex = Vec::<ComplexInstruction>::new();
    let mut in_instruction_text = false;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid document identity XML: {error}")))?;
        let word = match &event {
            Event::Start(element) | Event::Empty(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            Event::End(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            _ => false,
        };
        match event {
            Event::Start(element) => {
                if body_depth.is_none()
                    && word
                    && matches_local_name(element.name().as_ref(), b"body")
                {
                    body_depth = Some(depth);
                } else if body_depth.is_some() {
                    collect_body_identity_values(
                        &element,
                        reader.resolver(),
                        sdt_properties_depth.is_some(),
                        &mut values,
                    )?;
                    if collect_body_reference_values(
                        &element,
                        reader.resolver(),
                        &mut complex,
                        &mut values.reference_names,
                    )? {
                        in_instruction_text = true;
                    }
                    if word && matches_local_name(element.name().as_ref(), b"sdtPr") {
                        sdt_properties_depth = Some(depth);
                    }
                }
                depth += 1;
            }
            Event::Empty(element) if body_depth.is_some() => {
                collect_body_identity_values(
                    &element,
                    reader.resolver(),
                    sdt_properties_depth.is_some(),
                    &mut values,
                )?;
                collect_body_reference_values(
                    &element,
                    reader.resolver(),
                    &mut complex,
                    &mut values.reference_names,
                )?;
            }
            Event::Text(text) if body_depth.is_some() && in_instruction_text => {
                append_complex_instruction_text(
                    &text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?,
                    &mut complex,
                )?;
            }
            Event::CData(text) if body_depth.is_some() && in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    instruction.text.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?);
                }
            }
            Event::End(element) => {
                if word && matches_local_name(element.name().as_ref(), b"instrText") {
                    in_instruction_text = false;
                }
                depth = depth.saturating_sub(1);
                if sdt_properties_depth == Some(depth) {
                    sdt_properties_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
            }
            Event::Eof => return Ok(values),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_body_identity_values(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    in_sdt_properties: bool,
    values: &mut BodyIdentityValues,
) -> Result<()> {
    let name = element.name();
    let local = local_name(name.as_ref());
    let namespace = resolver.resolve_element(element.name()).0;
    if namespace_is_word(&namespace) && matches!(local, b"bookmarkStart" | b"bookmarkEnd") {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
        {
            values.bookmark_ids.push(value.clone());
            values
                .bookmark_events
                .push((value, local == b"bookmarkStart"));
        }
        if local == b"bookmarkStart"
            && let Some((_, value)) =
                resolved_element_attribute(element, resolver, b"name", AttributeNamespace::Word)?
        {
            values.bookmark_names.push(value);
        }
    } else if namespace_is_word(&namespace)
        && matches!(
            local,
            b"commentRangeStart" | b"commentRangeEnd" | b"commentReference"
        )
    {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
        {
            let kind = match local {
                b"commentRangeStart" => 0,
                b"commentRangeEnd" => 1,
                b"commentReference" => 2,
                _ => unreachable!("comment identity element was matched"),
            };
            values.comment_ids.push(value.clone());
            values.comment_events.push((value, kind));
        }
    } else if namespace_is_word(&namespace) && in_sdt_properties && local == b"id" {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"val", AttributeNamespace::Word)?
        {
            values.content_control_ids.push(value);
        }
    } else if namespace_matches(&namespace, WP_NS) && local == b"docPr" {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Unbound)?
        {
            values.drawing_ids.push(value);
        }
    } else if namespace_is_non_visual_drawing(&namespace)
        && local == b"cNvPr"
        && let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Unbound)?
    {
        values.non_visual_drawing_ids.push(value);
    }
    Ok(())
}

fn collect_body_reference_values(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    complex: &mut Vec<ComplexInstruction>,
    references: &mut Vec<String>,
) -> Result<bool> {
    let (namespace, local) = resolver.resolve_element(element.name());
    if !namespace_is_word(&namespace) {
        return Ok(false);
    }
    if local.as_ref() == b"hyperlink" {
        if let Some((_, anchor)) =
            resolved_element_attribute(element, resolver, b"anchor", AttributeNamespace::Word)?
        {
            references.push(anchor);
        }
    } else if local.as_ref() == b"fldSimple" {
        if let Some((_, instruction)) =
            resolved_element_attribute(element, resolver, b"instr", AttributeNamespace::Word)?
        {
            collect_reference_field_name(&instruction, references);
        }
    } else if local.as_ref() == b"fldChar" {
        match resolved_element_attribute(
            element,
            resolver,
            b"fldCharType",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value)
        .as_deref()
        {
            Some("begin") => complex.push(ComplexInstruction {
                text: String::new(),
                collecting: true,
            }),
            Some("separate") => {
                if let Some(instruction) = complex.last_mut() {
                    instruction.collecting = false;
                }
            }
            Some("end") => {
                if let Some(instruction) = complex.pop() {
                    collect_reference_field_name(&instruction.text, references);
                }
            }
            _ => {}
        }
    }
    Ok(local.as_ref() == b"instrText")
}

fn append_complex_instruction_text(text: &str, complex: &mut [ComplexInstruction]) -> Result<()> {
    if let Some(instruction) = complex.last_mut()
        && instruction.collecting
    {
        let unescaped = quick_xml::escape::unescape(text)
            .map_err(|error| Error::Other(format!("invalid field instruction entity: {error}")))?;
        instruction.text.push_str(&unescaped);
    }
    Ok(())
}

fn collect_reference_field_name(instruction: &str, references: &mut Vec<String>) {
    if let Some(name) = reference_field_name(instruction) {
        references.push(name);
    }
}

fn reference_field_name(instruction: &str) -> Option<String> {
    let field = Field::new(instruction, "");
    if matches!(field.instruction.name.as_str(), "REF" | "PAGEREF") {
        field
            .instruction
            .arguments
            .first()
            .and_then(argument_text)
            .map(str::to_owned)
    } else {
        None
    }
}

#[derive(Clone, Copy)]
enum AttributeNamespace {
    Relationship,
    Word,
    Unbound,
    Bound(&'static str),
}

fn resolved_element_attribute(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    local: &[u8],
    expected: AttributeNamespace,
) -> Result<Option<(Vec<u8>, String)>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| Error::Other(format!("invalid XML attribute: {error}")))?;
        let (namespace, resolved_local) = resolver.resolve_attribute(attribute.key);
        let namespace_matches = match expected {
            AttributeNamespace::Relationship => namespace_matches(&namespace, R_NS),
            AttributeNamespace::Word => namespace_is_word(&namespace),
            AttributeNamespace::Unbound => matches!(namespace, ResolveResult::Unbound),
            AttributeNamespace::Bound(expected) => namespace_matches(&namespace, expected),
        };
        if namespace_matches && resolved_local.as_ref() == local {
            let raw = std::str::from_utf8(attribute.value.as_ref())
                .map_err(|error| Error::Other(format!("invalid XML attribute value: {error}")))?;
            let decoded = quick_xml::escape::unescape(raw)
                .map_err(|error| Error::Other(format!("invalid XML attribute entity: {error}")))?;
            return Ok(Some((
                attribute.key.as_ref().to_vec(),
                decoded.into_owned(),
            )));
        }
    }
    Ok(None)
}

const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

fn namespace_matches(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected.as_bytes())
}

fn namespace_is_non_visual_drawing(namespace: &ResolveResult<'_>) -> bool {
    [PIC_NS, DRAWING_NS, WPS_NS]
        .iter()
        .any(|expected| namespace_matches(namespace, expected))
}

fn patch_body_identity_attributes(xml: &[u8], remap: &BodyIdentityRemap) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut body_depth = None;
    let mut sdt_properties_depth = None;
    let mut edits = Vec::new();
    let mut complex = Vec::<ComplexInstructionEdit>::new();
    let mut in_instruction_text = false;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid document identity XML: {error}")))?;
        let after = reader.buffer_position() as usize;
        let identity_namespace = match &event {
            Event::Start(element) | Event::Empty(element) => BodyIdentityNamespace::from_resolved(
                &reader.resolver().resolve_element(element.name()).0,
            ),
            Event::End(element) => BodyIdentityNamespace::from_resolved(
                &reader.resolver().resolve_element(element.name()).0,
            ),
            _ => BodyIdentityNamespace::default(),
        };
        match event {
            Event::Start(element) => {
                if body_depth.is_none()
                    && identity_namespace.word
                    && matches_local_name(element.name().as_ref(), b"body")
                {
                    body_depth = Some(depth);
                } else if body_depth.is_some() {
                    collect_body_identity_edits(
                        xml,
                        before,
                        after,
                        &element,
                        reader.resolver(),
                        identity_namespace,
                        sdt_properties_depth.is_some(),
                        remap,
                        &mut edits,
                    )?;
                    if collect_body_reference_edits(
                        xml,
                        before,
                        after,
                        &element,
                        reader.resolver(),
                        remap,
                        &mut complex,
                        &mut edits,
                    )? {
                        in_instruction_text = true;
                    }
                    if identity_namespace.word
                        && matches_local_name(element.name().as_ref(), b"sdtPr")
                    {
                        sdt_properties_depth = Some(depth);
                    }
                }
                depth += 1;
            }
            Event::Empty(element) if body_depth.is_some() => {
                collect_body_identity_edits(
                    xml,
                    before,
                    after,
                    &element,
                    reader.resolver(),
                    identity_namespace,
                    sdt_properties_depth.is_some(),
                    remap,
                    &mut edits,
                )?;
                collect_body_reference_edits(
                    xml,
                    before,
                    after,
                    &element,
                    reader.resolver(),
                    remap,
                    &mut complex,
                    &mut edits,
                )?;
            }
            Event::Text(text) if body_depth.is_some() && in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    let decoded = text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::Other(format!("invalid field instruction entity: {error}"))
                    })?;
                    instruction.text.push_str(&decoded);
                    instruction.spans.push((before, after));
                }
            }
            Event::CData(text) if body_depth.is_some() && in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    instruction.text.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?);
                    instruction.spans.push((before, after));
                }
            }
            Event::End(element) => {
                if identity_namespace.word
                    && matches_local_name(element.name().as_ref(), b"instrText")
                {
                    in_instruction_text = false;
                }
                depth = depth.saturating_sub(1);
                if sdt_properties_depth == Some(depth) {
                    sdt_properties_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut updated = xml.to_vec();
    edits.sort_by_key(|edit: &FieldSourceEdit| edit.start);
    for edit in edits.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

pub(crate) fn patch_bookmark_ids(
    xml: &[u8],
    bookmark_ids: &BTreeMap<String, String>,
) -> Result<Vec<u8>> {
    patch_body_identity_attributes(
        xml,
        &BodyIdentityRemap {
            bookmark_ids: bookmark_ids.clone(),
            ..Default::default()
        },
    )
}

pub(crate) fn freshen_content_fragment_identities(
    document: &mut Document,
    wrapped_fragment: &[u8],
) -> Result<Vec<u8>> {
    let wrapped_fragment = &uniquify_drawing_ids_in_xml(wrapped_fragment)?;
    let values = body_identity_values(wrapped_fragment)?;
    let mut open_bookmarks = Vec::new();
    let mut seen_bookmarks = HashSet::new();
    for (id, start) in &values.bookmark_events {
        if *start {
            if !seen_bookmarks.insert(id.clone()) {
                return Err(Error::Other(
                    "content fragment bookmark ownership is incomplete or ambiguous".to_owned(),
                ));
            }
            open_bookmarks.push(id.as_str());
        } else if open_bookmarks.pop() != Some(id.as_str()) {
            return Err(Error::Other(
                "content fragment bookmark ownership is incomplete or ambiguous".to_owned(),
            ));
        }
    }
    if !open_bookmarks.is_empty() {
        return Err(Error::Other(
            "content fragment bookmark ownership is incomplete or ambiguous".to_owned(),
        ));
    }
    let mut state = BodyIdentityState::from_documents(std::slice::from_ref(document))?;
    let mut remap = BodyIdentityRemap {
        drop_paragraph_identities: true,
        ..Default::default()
    };
    for value in values.bookmark_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.bookmark_ids.entry(value) {
            entry.insert(document.identifiers.reserve_bookmark_id()?.to_string());
        }
    }
    for value in values.content_control_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) =
            remap.content_control_ids.entry(value)
        {
            entry.insert(state.allocate_content_control_id()?);
        }
    }
    for value in values.drawing_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.drawing_ids.entry(value) {
            entry.insert(document.identifiers.reserve_drawing_id()?.to_string());
        }
    }
    for value in values.non_visual_drawing_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) =
            remap.non_visual_drawing_ids.entry(value)
        {
            entry.insert(state.allocate_non_visual_drawing_id()?);
        }
    }
    for value in values.bookmark_names {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.bookmark_names.entry(value)
        {
            entry.insert(state.allocate_name()?);
        }
    }
    patch_body_identity_attributes(wrapped_fragment, &remap)
}

fn validate_fragment_identity_ownership(values: &BodyIdentityValues) -> Result<()> {
    let mut open_bookmarks = Vec::new();
    let mut seen_bookmarks = HashSet::new();
    for (id, start) in &values.bookmark_events {
        if *start {
            if !seen_bookmarks.insert(id.clone()) {
                return Err(Error::Other(
                    "document fragment bookmark ownership is incomplete or ambiguous".to_owned(),
                ));
            }
            open_bookmarks.push(id.as_str());
        } else if open_bookmarks.pop() != Some(id.as_str()) {
            return Err(Error::Other(
                "document fragment bookmark ownership is incomplete or ambiguous".to_owned(),
            ));
        }
    }
    if !open_bookmarks.is_empty() {
        return Err(Error::Other(
            "document fragment bookmark ownership is incomplete or ambiguous".to_owned(),
        ));
    }

    let mut comment_counts = BTreeMap::<&str, [usize; 3]>::new();
    for (id, kind) in &values.comment_events {
        comment_counts.entry(id).or_default()[usize::from(*kind)] += 1;
    }
    if comment_counts.values().any(|counts| *counts != [1, 1, 1]) {
        return Err(Error::Other(
            "document fragment comment ownership is incomplete or ambiguous".to_owned(),
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn collect_body_identity_edits(
    xml: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    namespace: BodyIdentityNamespace,
    in_sdt_properties: bool,
    remap: &BodyIdentityRemap,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<()> {
    let name = element.name();
    let local = local_name(name.as_ref());
    if namespace.word && matches!(local, b"bookmarkStart" | b"bookmarkEnd") {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"id",
            AttributeNamespace::Word,
            &remap.bookmark_ids,
            edits,
        )?;
        if local == b"bookmarkStart" {
            add_identity_attribute_edit(
                xml,
                start,
                end,
                element,
                resolver,
                b"name",
                AttributeNamespace::Word,
                &remap.bookmark_names,
                edits,
            )?;
        }
    } else if namespace.word
        && matches!(
            local,
            b"commentRangeStart" | b"commentRangeEnd" | b"commentReference"
        )
    {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"id",
            AttributeNamespace::Word,
            &remap.comment_ids,
            edits,
        )?;
    } else if namespace.word && in_sdt_properties && local == b"id" {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"val",
            AttributeNamespace::Word,
            &remap.content_control_ids,
            edits,
        )?;
    } else if namespace.wordprocessing_drawing && local == b"docPr" {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"id",
            AttributeNamespace::Unbound,
            &remap.drawing_ids,
            edits,
        )?;
    } else if namespace.non_visual_drawing && local == b"cNvPr" {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"id",
            AttributeNamespace::Unbound,
            &remap.non_visual_drawing_ids,
            edits,
        )?;
    } else if remap.drop_paragraph_identities && namespace.word && matches!(local, b"p" | b"tr") {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::Other(format!("invalid XML attribute: {error}")))?;
            let (attribute_namespace, attribute_local) = resolver.resolve_attribute(attribute.key);
            if !namespace_matches(&attribute_namespace, W14_NS)
                || !matches!(attribute_local.as_ref(), b"paraId" | b"textId")
            {
                continue;
            }
            let Some((name_start, _, value_end)) =
                attribute_source_span(&xml[start..end], attribute.key.as_ref())
            else {
                return Err(Error::Other(
                    "document identity attribute source was not found".to_owned(),
                ));
            };
            // Remove the whitespace before the name with the attribute.
            let removed_start = xml[start..start + name_start]
                .iter()
                .rposition(|byte| !byte.is_ascii_whitespace())
                .map_or(0, |last| last + 1);
            edits.push(FieldSourceEdit {
                start: start + removed_start,
                end: start + value_end + 1,
                replacement: Vec::new(),
            });
        }
    }
    Ok(())
}

#[derive(Clone, Copy, Default)]
struct BodyIdentityNamespace {
    word: bool,
    wordprocessing_drawing: bool,
    non_visual_drawing: bool,
}

impl BodyIdentityNamespace {
    fn from_resolved(namespace: &ResolveResult<'_>) -> Self {
        Self {
            word: namespace_is_word(namespace),
            wordprocessing_drawing: namespace_matches(namespace, WP_NS),
            non_visual_drawing: namespace_is_non_visual_drawing(namespace),
        }
    }
}

fn add_identity_attribute_edit(
    xml: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    local: &[u8],
    expected: AttributeNamespace,
    replacements: &BTreeMap<String, String>,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<()> {
    let Some((key, old)) = resolved_element_attribute(element, resolver, local, expected)? else {
        return Ok(());
    };
    let Some(replacement) = replacements.get(&old) else {
        return Ok(());
    };
    let Some((relative_start, relative_end)) = attribute_value_span(&xml[start..end], &key) else {
        return Err(Error::Other(
            "document identity attribute source was not found".to_owned(),
        ));
    };
    edits.push(FieldSourceEdit {
        start: start + relative_start,
        end: start + relative_end,
        replacement: replacement.as_bytes().to_vec(),
    });
    Ok(())
}

struct ComplexInstructionEdit {
    text: String,
    collecting: bool,
    spans: Vec<(usize, usize)>,
}

#[allow(clippy::too_many_arguments)]
fn collect_body_reference_edits(
    xml: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    remap: &BodyIdentityRemap,
    complex: &mut Vec<ComplexInstructionEdit>,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<bool> {
    let (namespace, local) = resolver.resolve_element(element.name());
    if !namespace_is_word(&namespace) {
        return Ok(false);
    }
    if local.as_ref() == b"hyperlink" {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"anchor",
            AttributeNamespace::Word,
            &remap.bookmark_names,
            edits,
        )?;
    } else if local.as_ref() == b"fldSimple" {
        if let Some((key, instruction)) =
            resolved_element_attribute(element, resolver, b"instr", AttributeNamespace::Word)?
            && let Some(updated) = remap_reference_instruction(&instruction, &remap.bookmark_names)
        {
            add_attribute_value_edit(xml, start, end, &key, &updated, true, edits)?;
        }
    } else if local.as_ref() == b"fldChar" {
        match resolved_element_attribute(
            element,
            resolver,
            b"fldCharType",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value)
        .as_deref()
        {
            Some("begin") => complex.push(ComplexInstructionEdit {
                text: String::new(),
                collecting: true,
                spans: Vec::new(),
            }),
            Some("separate") => {
                if let Some(instruction) = complex.last_mut() {
                    instruction.collecting = false;
                }
            }
            Some("end") => {
                if let Some(instruction) = complex.pop()
                    && let Some(updated) =
                        remap_reference_instruction(&instruction.text, &remap.bookmark_names)
                    && let Some((first, remaining)) = instruction.spans.split_first()
                {
                    edits.push(FieldSourceEdit {
                        start: first.0,
                        end: first.1,
                        replacement: quick_xml::escape::escape(&updated)
                            .into_owned()
                            .into_bytes(),
                    });
                    edits.extend(remaining.iter().map(|(start, end)| FieldSourceEdit {
                        start: *start,
                        end: *end,
                        replacement: Vec::new(),
                    }));
                }
            }
            _ => {}
        }
    }
    Ok(local.as_ref() == b"instrText")
}

fn add_attribute_value_edit(
    xml: &[u8],
    start: usize,
    end: usize,
    key: &[u8],
    replacement: &str,
    escape: bool,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<()> {
    let Some((relative_start, relative_end)) = attribute_value_span(&xml[start..end], key) else {
        return Err(Error::Other(
            "document identity attribute source was not found".to_owned(),
        ));
    };
    let replacement = if escape {
        quick_xml::escape::escape(replacement)
            .into_owned()
            .into_bytes()
    } else {
        replacement.as_bytes().to_vec()
    };
    edits.push(FieldSourceEdit {
        start: start + relative_start,
        end: start + relative_end,
        replacement,
    });
    Ok(())
}

fn remap_reference_instruction(
    instruction: &str,
    names: &BTreeMap<String, String>,
) -> Option<String> {
    let old = reference_field_name(instruction)?;
    let replacement = names.get(&old)?;
    let command_end = instruction
        .find(char::is_whitespace)
        .unwrap_or(instruction.len());
    let relative = instruction[command_end..].find(&old)?;
    let start = command_end + relative;
    let mut updated = instruction.to_owned();
    updated.replace_range(start..start + old.len(), replacement);
    Some(updated)
}

pub(crate) fn attribute_value_span(
    element: &[u8],
    attribute_name: &[u8],
) -> Option<(usize, usize)> {
    attribute_source_span(element, attribute_name)
        .map(|(_, value_start, value_end)| (value_start, value_end))
}

/// Return where the name of one attribute starts in a start tag, and where
/// its value starts and ends.
fn attribute_source_span(element: &[u8], attribute_name: &[u8]) -> Option<(usize, usize, usize)> {
    let mut index = 1usize;
    while index < element.len() && !element[index].is_ascii_whitespace() {
        index += 1;
    }
    while index < element.len() {
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= element.len() || matches!(element[index], b'>' | b'/') {
            return None;
        }
        let name_start = index;
        while index < element.len()
            && !element[index].is_ascii_whitespace()
            && element[index] != b'='
        {
            index += 1;
        }
        let name_end = index;
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        if element.get(index) != Some(&b'=') {
            return None;
        }
        index += 1;
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *element.get(index)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        index += 1;
        let value_start = index;
        while index < element.len() && element[index] != quote {
            index += 1;
        }
        let value_end = index;
        if &element[name_start..name_end] == attribute_name {
            return Some((name_start, value_start, value_end));
        }
        index += 1;
    }
    None
}

fn empty_section_properties() -> CT_SectPr {
    CT_SectPr {
        page_width: None,
        page_height: None,
        orientation: None,
        margin_top: None,
        margin_right: None,
        margin_bottom: None,
        margin_left: None,
        gutter: None,
        header_distance: None,
        footer_distance: None,
        section_type: None,
        columns: None,
        page_number: None,
        footnote_pr: None,
        endnote_pr: None,
        paper_source: None,
        page_borders: None,
        line_numbers: None,
        vertical_alignment: None,
        text_direction: None,
        doc_grid: None,
        title_pg: None,
        header_refs: Vec::new(),
        footer_refs: Vec::new(),
        extra_xml: Vec::new(),
        extra_xml_positions: Vec::new(),
        change: None,
    }
}

struct TextBoxCachePatch {
    location: crate::ContentLocation,
    replacements: Vec<(Vec<u8>, Vec<u8>)>,
    updated: usize,
}

fn collect_ref_cached_note_references(field: &Field, output: &mut Vec<(StoryKind, i32)>) {
    if let Some(runs) = field.cached_result_runs() {
        for content in runs.iter().flat_map(|run| &run.content) {
            match content {
                RunContent::FootnoteRef { id, .. } => output.push((StoryKind::Footnote, *id)),
                RunContent::EndnoteRef { id, .. } => output.push((StoryKind::Endnote, *id)),
                RunContent::Field(field) => collect_ref_cached_note_references(field, output),
                _ => {}
            }
        }
    }
}

fn ref_note_reference_counts(package: &OpcPackage) -> Result<HashMap<(StoryKind, i32), usize>> {
    let mut counts = HashMap::new();
    for (name, xml) in &package.parts {
        if !name.ends_with(".xml")
            || ![
                b"footnoteReference".as_slice(),
                b"endnoteReference".as_slice(),
            ]
            .iter()
            .any(|needle| xml.windows(needle.len()).any(|window| window == *needle))
        {
            continue;
        }
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).map_err(|error| {
                Error::Other(format!(
                    "REF f source reference graph is malformed: {error}"
                ))
            })? {
                Event::Start(element) | Event::Empty(element) => {
                    let (namespace, local) = reader.resolver().resolve_element(element.name());
                    if namespace_is_word(&namespace) {
                        let kind = match local.as_ref() {
                            b"footnoteReference" => Some(StoryKind::Footnote),
                            b"endnoteReference" => Some(StoryKind::Endnote),
                            _ => None,
                        };
                        if let Some(kind) = kind {
                            let (_, id) = resolved_element_attribute(
                                &element,
                                reader.resolver(),
                                b"id",
                                AttributeNamespace::Word,
                            )?
                            .ok_or_else(|| {
                                Error::Other("REF f source reference has no ID".into())
                            })?;
                            let id = id.parse::<i32>().map_err(|_| {
                                Error::Other("REF f source reference ID is invalid".into())
                            })?;
                            *counts.entry((kind, id)).or_insert(0) += 1;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    Ok(counts)
}

#[derive(Clone, Copy)]
enum RefCommentCompanionKind {
    Extended,
    Ids,
}

impl RefCommentCompanionKind {
    fn namespace(self) -> &'static str {
        match self {
            Self::Extended => rdocx_oxml::comments_extended::W15_NS,
            Self::Ids => "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        }
    }

    fn root(self) -> &'static [u8] {
        match self {
            Self::Extended => b"commentsEx",
            Self::Ids => b"commentsIds",
        }
    }

    fn item(self) -> &'static [u8] {
        match self {
            Self::Extended => b"commentEx",
            Self::Ids => b"commentId",
        }
    }

    fn relationship(self) -> &'static str {
        match self {
            Self::Extended => crate::comments::COMMENTS_EXTENDED_REL_TYPE,
            Self::Ids => "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        }
    }
}

struct RefCommentCompanionEntry {
    para_id: String,
    parent: Option<String>,
    durable_id: Option<String>,
    span: std::ops::Range<usize>,
    xml: Vec<u8>,
}

fn ref_comment_companion_entries(
    xml: &[u8],
    kind: RefCommentCompanionKind,
) -> Result<(Vec<RefCommentCompanionEntry>, usize)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut root_end = None;
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    let mut durable_seen = HashSet::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "REF f annotation companion XML is malformed: {error}"
            ))
        })?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = xml.get(end.saturating_sub(2)..end) == Some(b"/>");
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                let owned = namespace_matches(&namespace, kind.namespace());
                if depth == 0 {
                    if !owned || local.as_ref() != kind.root() || saw_root || empty {
                        return Err(Error::Other(
                            "REF f annotation companion has no unique qualified root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 1 && owned && local.as_ref() == kind.item() {
                    let attribute = |name| {
                        resolved_element_attribute(
                            &element,
                            reader.resolver(),
                            name,
                            AttributeNamespace::Bound(kind.namespace()),
                        )
                        .map(|value| value.map(|(_, value)| value))
                    };
                    let para_id = attribute(b"paraId")?.ok_or_else(|| {
                        Error::Other("REF f annotation companion entry has no paragraph ID".into())
                    })?;
                    if para_id.len() != 8
                        || u32::from_str_radix(&para_id, 16).is_err()
                        || !seen.insert(para_id.to_ascii_uppercase())
                    {
                        return Err(Error::Other(
                            "REF f annotation companion paragraph ID is invalid or ambiguous"
                                .into(),
                        ));
                    }
                    let parent = attribute(b"paraIdParent")?;
                    let durable_id = attribute(b"durableId")?;
                    if matches!(kind, RefCommentCompanionKind::Ids)
                        && !durable_id.as_ref().is_some_and(|id| {
                            id.len() == 8
                                && u32::from_str_radix(id, 16).is_ok()
                                && durable_seen.insert(id.to_ascii_uppercase())
                        })
                    {
                        return Err(Error::Other(
                            "REF f annotation companion durable ID is absent or invalid".into(),
                        ));
                    }
                    if !empty {
                        reader
                            .read_to_end_into(element.name(), &mut Vec::new())
                            .map_err(|error| {
                                Error::Other(format!(
                                    "REF f annotation companion entry is unclosed: {error}"
                                ))
                            })?;
                    }
                    let span = start..reader.buffer_position() as usize;
                    let scope = crate::document::story_namespace_scope_at(xml, start)?;
                    entries.push(RefCommentCompanionEntry {
                        para_id,
                        parent,
                        durable_id,
                        xml: crate::document::close_content_fragment_namespaces(
                            &xml[span.clone()],
                            &scope,
                        )?,
                        span,
                    });
                    buffer.clear();
                    continue;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Other("REF f annotation companion XML is unbalanced".into())
                })?;
                if depth == 0 {
                    let (namespace, local) = reader.resolver().resolve_element(element.name());
                    if !namespace_matches(&namespace, kind.namespace())
                        || local.as_ref() != kind.root()
                        || root_end.replace(start).is_some()
                    {
                        return Err(Error::Other(
                            "REF f annotation companion root is ambiguous".into(),
                        ));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if depth != 0 || !saw_root {
        return Err(Error::Other(
            "REF f annotation companion XML is unclosed".into(),
        ));
    }
    Ok((
        entries,
        root_end
            .ok_or_else(|| Error::Other("REF f annotation companion root end is missing".into()))?,
    ))
}

fn ref_comment_paragraph_ids(xml: &[u8]) -> Result<Vec<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut ids = Vec::new();
    let mut seen = HashSet::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("REF f annotation payload is malformed: {error}"))
        })? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace)
                    && matches!(
                        local.as_ref(),
                        b"commentReference" | b"commentRangeStart" | b"commentRangeEnd"
                    )
                {
                    return Err(Error::Other("REF f annotation payload contains an unqualified recursive annotation graph".into()));
                }
                if namespace_is_word(&namespace) && local.as_ref() == b"p" {
                    let (_, id) = resolved_element_attribute(&element, reader.resolver(), b"paraId", AttributeNamespace::Bound(W14_NS))?
                        .ok_or_else(|| Error::Other("REF f annotation paragraph lacks qualified identity for companion closure".into()))?;
                    if id.len() != 8
                        || u32::from_str_radix(&id, 16).is_err()
                        || !seen.insert(id.to_ascii_uppercase())
                    {
                        return Err(Error::Other(
                            "REF f annotation paragraph identity is invalid or duplicated".into(),
                        ));
                    }
                    ids.push(id);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(ids)
}

fn ref_comment_paragraph_identity_counts(xml: &[u8]) -> Result<HashMap<String, usize>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut counts = HashMap::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "REF f annotation owner identity inventory is malformed: {error}"
            ))
        })? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace)
                    && local.as_ref() == b"p"
                    && let Some((_, id)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"paraId",
                        AttributeNamespace::Bound(W14_NS),
                    )?
                {
                    *counts.entry(id.to_ascii_uppercase()).or_insert(0) += 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(counts)
}

fn ref_comment_companion_part(
    document: &Document,
    kind: RefCommentCompanionKind,
) -> Result<Option<String>> {
    let relationships = document.package.get_part_rels(&document.doc_part_name);
    let mut matches = relationships
        .into_iter()
        .flat_map(|relationships| &relationships.items)
        .filter(|relationship| relationship.rel_type == kind.relationship());
    let first = matches.next();
    if matches.next().is_some()
        || first
            .is_some_and(|relationship| !crate::document::relationship_is_internal(relationship))
    {
        return Err(Error::Other(
            "REF f annotation companion relationship owner is ambiguous or external".into(),
        ));
    }
    first
        .map(|relationship| {
            let part =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            if document.package.get_part(&part).is_none() {
                return Err(Error::Other(
                    "REF f annotation companion relationship target is missing".into(),
                ));
            }
            Ok(part)
        })
        .transpose()
}

struct RefCommentCompanionCopy {
    kind: RefCommentCompanionKind,
    part: String,
    xml: Vec<u8>,
    replace_para_id: Option<String>,
}

struct RefCommentCopy {
    part: String,
    xml: Vec<u8>,
    replace_id: Option<i32>,
    companions: Vec<RefCommentCompanionCopy>,
}

fn ref_comment_occupied_identity_values(package: &OpcPackage) -> Result<HashSet<u32>> {
    let mut occupied = HashSet::new();
    for (part, xml) in &package.parts {
        if !part.ends_with(".xml") {
            continue;
        }
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).map_err(|error| {
                Error::Other(format!(
                    "REF f annotation identity inventory is malformed: {error}"
                ))
            })? {
                Event::Start(element) | Event::Empty(element) => {
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| {
                            Error::Other(format!(
                                "REF f annotation identity attribute is malformed: {error}"
                            ))
                        })?;
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        let namespace_owned = namespace_matches(&namespace, W14_NS)
                            || namespace_matches(
                                &namespace,
                                RefCommentCompanionKind::Extended.namespace(),
                            )
                            || namespace_matches(
                                &namespace,
                                RefCommentCompanionKind::Ids.namespace(),
                            );
                        if namespace_owned
                            && matches!(local.as_ref(), b"paraId" | b"paraIdParent" | b"durableId")
                        {
                            let value = attribute
                                .decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Implicit1_0,
                                    element.decoder(),
                                )
                                .map_err(|error| {
                                    Error::Other(format!(
                                        "REF f annotation identity value is malformed: {error}"
                                    ))
                                })?;
                            if value.len() != 8 {
                                return Err(Error::Other("REF f annotation identity is not an eight-digit hexadecimal value".into()));
                            }
                            occupied.insert(u32::from_str_radix(&value, 16).map_err(|_| {
                                Error::Other("REF f annotation identity is invalid".into())
                            })?);
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    Ok(occupied)
}

fn patch_ref_comment_identity_attributes(
    xml: &[u8],
    comment_ids: &BTreeMap<String, String>,
    paragraph_ids: &BTreeMap<String, String>,
    durable_ids: &BTreeMap<String, String>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut edits = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "REF f annotation identity patch XML is malformed: {error}"
            ))
        })?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if namespace_is_word(&namespace) && local.as_ref() == b"comment" {
                    add_identity_attribute_edit(
                        xml,
                        start,
                        end,
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                        comment_ids,
                        &mut edits,
                    )?;
                }
                if namespace_is_word(&namespace) && local.as_ref() == b"p" {
                    add_identity_attribute_edit(
                        xml,
                        start,
                        end,
                        &element,
                        reader.resolver(),
                        b"paraId",
                        AttributeNamespace::Bound(W14_NS),
                        paragraph_ids,
                        &mut edits,
                    )?;
                }
                for kind in [
                    RefCommentCompanionKind::Extended,
                    RefCommentCompanionKind::Ids,
                ] {
                    if namespace_matches(&namespace, kind.namespace())
                        && local.as_ref() == kind.item()
                    {
                        add_identity_attribute_edit(
                            xml,
                            start,
                            end,
                            &element,
                            reader.resolver(),
                            b"paraId",
                            AttributeNamespace::Bound(kind.namespace()),
                            paragraph_ids,
                            &mut edits,
                        )?;
                        add_identity_attribute_edit(
                            xml,
                            start,
                            end,
                            &element,
                            reader.resolver(),
                            b"durableId",
                            AttributeNamespace::Bound(kind.namespace()),
                            durable_ids,
                            &mut edits,
                        )?;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut result = xml.to_vec();
    for edit in edits.into_iter().rev() {
        result.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(result)
}

fn prepare_ref_comment_copy(
    document: &mut Document,
    original: &Document,
    source_id: i32,
    replace_id: Option<i32>,
    occupied_comments: &mut HashSet<i32>,
    occupied_identities: &mut HashSet<u32>,
) -> Result<(i32, RefCommentCopy)> {
    let (part, source) = original.fragment_comment_dependency(source_id)?;
    if document.comments_part_name.as_deref() != Some(&part)
        || !fragment_note_references(&source)?.is_empty()
    {
        return Err(Error::Other(
            "REF f annotation has an unqualified relationship owner or recursive note payload"
                .into(),
        ));
    }
    let paragraph_ids = ref_comment_paragraph_ids(&source)?;
    let paragraph_owners = ref_comment_paragraph_identity_counts(
        original
            .package
            .get_part(&part)
            .ok_or_else(|| Error::Other("REF f annotation owner part is absent".into()))?,
    )?;
    if paragraph_ids
        .iter()
        .any(|id| paragraph_owners.get(&id.to_ascii_uppercase()) != Some(&1))
    {
        return Err(Error::Other(
            "REF f annotation paragraph identity is shared with another physical owner".into(),
        ));
    }
    let source_last = paragraph_ids.last().ok_or_else(|| {
        Error::Other("REF f annotation source has no qualified last paragraph identity".into())
    })?;
    let old_last = replace_id
        .map(|id| {
            original
                .fragment_comment_dependency(id)
                .and_then(|(old_part, xml)| {
                    if old_part != part || id == source_id {
                        return Err(Error::Other(
                            "REF f cached annotation is borrowed or differently owned".into(),
                        ));
                    }
                    let old_paragraphs = ref_comment_paragraph_ids(&xml)?;
                    if old_paragraphs.iter().any(|id|paragraph_owners.get(&id.to_ascii_uppercase())!=Some(&1)) {
                        return Err(Error::Other("REF f cached annotation paragraph identity is shared with another physical owner".into()));
                    }
                    old_paragraphs
                        .last()
                        .cloned()
                        .ok_or_else(|| {
                            Error::Other(
                                "REF f cached annotation has no qualified last paragraph identity"
                                    .into(),
                            )
                        })
                })
        })
        .transpose()?;
    let mut para_map = BTreeMap::new();
    for old in &paragraph_ids {
        para_map.insert(
            old.clone(),
            crate::comments::allocate_para_id_from_occupied(occupied_identities)?,
        );
    }
    let id = if let Some(id) = replace_id {
        if !occupied_comments.contains(&id) {
            return Err(Error::Other(
                "REF f cached annotation owner is absent".into(),
            ));
        }
        id
    } else {
        let mut id = 0i32;
        while occupied_comments.contains(&id) {
            id = id
                .checked_add(1)
                .ok_or_else(|| Error::Other("REF f annotation IDs are exhausted".into()))?;
        }
        occupied_comments.insert(id);
        id
    };
    let mut companions = Vec::new();
    for kind in [
        RefCommentCompanionKind::Extended,
        RefCommentCompanionKind::Ids,
    ] {
        let Some(companion_part) = ref_comment_companion_part(original, kind)? else {
            continue;
        };
        let xml = original.package.get_part(&companion_part).unwrap();
        let (entries, _) = ref_comment_companion_entries(xml, kind)?;
        let selected = entries
            .iter()
            .filter(|entry| entry.para_id.eq_ignore_ascii_case(source_last))
            .collect::<Vec<_>>();
        if selected.len() != 1
            || selected[0].parent.is_some()
            || entries.iter().any(|entry| {
                entry
                    .parent
                    .as_ref()
                    .is_some_and(|parent| parent.eq_ignore_ascii_case(source_last))
            })
        {
            return Err(Error::Other("REF f annotation companion source is missing, ambiguous or part of an unqualified reply graph".into()));
        }
        let replace_para_id = if let Some(old_last) = &old_last {
            let old = entries
                .iter()
                .filter(|entry| entry.para_id.eq_ignore_ascii_case(old_last))
                .collect::<Vec<_>>();
            if old.len() != 1
                || old[0].parent.is_some()
                || entries.iter().any(|entry| {
                    entry
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.eq_ignore_ascii_case(old_last))
                })
            {
                return Err(Error::Other(
                    "REF f cached annotation companion closure is ambiguous".into(),
                ));
            }
            Some(old[0].para_id.clone())
        } else {
            None
        };
        let mut durable_map = BTreeMap::new();
        if let Some(old) = &selected[0].durable_id {
            durable_map.insert(
                old.clone(),
                crate::comments::allocate_para_id_from_occupied(occupied_identities)?,
            );
        }
        let mut companion_para_map = para_map.clone();
        // Companion producers may use a different hexadecimal spelling of the same ID.
        companion_para_map.insert(selected[0].para_id.clone(), para_map[source_last].clone());
        companions.push(RefCommentCompanionCopy {
            kind,
            part: companion_part,
            xml: patch_ref_comment_identity_attributes(
                &selected[0].xml,
                &BTreeMap::new(),
                &companion_para_map,
                &durable_map,
            )?,
            replace_para_id,
        });
    }
    for relationship_id in relationship_ids_in_xml(&source)? {
        let relationship = original
            .package
            .get_part_rels(&part)
            .and_then(|relationships| relationships.get_by_id(&relationship_id))
            .ok_or_else(|| {
                Error::Other("REF f annotation payload relationship is absent".into())
            })?;
        if !crate::document::relationship_is_internal(relationship) {
            return Err(Error::Other(
                "REF f annotation external relationship payload lacks authenticated controls"
                    .into(),
            ));
        }
        discover_fragment_part_closure(
            &original.package,
            &OpcPackage::resolve_rel_target(&part, &relationship.target),
            &mut HashSet::new(),
        )?;
    }
    let mut xml = patch_ref_comment_identity_attributes(
        &source,
        &BTreeMap::from([(source_id.to_string(), id.to_string())]),
        &para_map,
        &BTreeMap::new(),
    )?;
    let wrapper = format!("<w:body xmlns:w=\"{W_NS}\">");
    let mut body = wrapper.as_bytes().to_vec();
    body.extend_from_slice(&xml);
    body.extend_from_slice(b"</w:body>");
    let mut remap = BodyIdentityRemap::default();
    for old in body_identity_values(&body)?.drawing_ids {
        if remap
            .drawing_ids
            .insert(old, document.identifiers.reserve_drawing_id()?.to_string())
            .is_some()
        {
            return Err(Error::Other(
                "REF f annotation source has duplicate drawing identities".into(),
            ));
        }
    }
    let mapped = patch_body_identity_attributes(&body, &remap)?;
    xml = mapped[wrapper.len()..mapped.len() - b"</w:body>".len()].to_vec();
    Ok((
        id,
        RefCommentCopy {
            part,
            xml,
            replace_id,
            companions,
        },
    ))
}

fn publish_ref_comment_copy(document: &mut Document, copy: RefCommentCopy) -> Result<()> {
    document.publish_fragment_comment_staged(&copy.part, copy.xml, copy.replace_id)?;
    for companion in copy.companions {
        let source = document.package.get_part(&companion.part).ok_or_else(|| {
            Error::Other("REF f annotation companion publication target is missing".into())
        })?;
        let (entries, end) = ref_comment_companion_entries(source, companion.kind)?;
        let mut updated = source.to_vec();
        if let Some(id) = companion.replace_para_id {
            let entries = entries
                .iter()
                .filter(|entry| entry.para_id == id)
                .collect::<Vec<_>>();
            if entries.len() != 1 {
                return Err(Error::Other(
                    "REF f cached annotation companion publication owner is ambiguous".into(),
                ));
            }
            updated.splice(entries[0].span.clone(), companion.xml);
        } else {
            updated.splice(end..end, companion.xml);
        }
        ref_comment_companion_entries(&updated, companion.kind)?;
        if matches!(companion.kind, RefCommentCompanionKind::Extended) {
            document.comments_extended = Some(
                rdocx_oxml::comments_extended::CT_CommentsEx::from_xml(&updated)?,
            );
        }
        document.package.set_part(&companion.part, updated);
    }
    Ok(())
}

fn ref_comment_reference_ids(cache: &RefBookmarkCache) -> Result<Vec<i32>> {
    let ids = cache
        .runs
        .iter()
        .flat_map(|run| &run.content)
        .filter_map(|content| match content {
            RunContent::CommentReference { id, .. } => Some(*id),
            _ => None,
        })
        .collect::<Vec<_>>();
    let mut unique = HashSet::new();
    if cache.comment_ranges.len() != ids.len() * 2 {
        return Err(Error::Other(
            "REF f annotation target lacks closed owned range endpoints".into(),
        ));
    }
    for id in &ids {
        let starts = cache.comment_ranges.iter().filter(|marker| matches!(marker,CommentRangeMarker::Start{id:marker_id,..} if marker_id==id)).count();
        let ends = cache
            .comment_ranges
            .iter()
            .filter(
                |marker| matches!(marker,CommentRangeMarker::End{id:marker_id,..} if marker_id==id),
            )
            .count();
        if !unique.insert(*id) || (starts, ends) != (1, 1) {
            return Err(Error::Other(
                "REF f annotation target range ownership is ambiguous".into(),
            ));
        }
    }
    Ok(ids)
}

fn ref_comment_graph_counts(package: &OpcPackage) -> Result<HashMap<i32, [usize; 3]>> {
    let mut counts = HashMap::new();
    for (part, xml) in &package.parts {
        if !part.ends_with(".xml")
            || !xml
                .windows(b"comment".len())
                .any(|value| value == b"comment")
        {
            continue;
        }
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).map_err(|error| {
                Error::Other(format!("REF f annotation graph is malformed: {error}"))
            })? {
                Event::Start(element) | Event::Empty(element) => {
                    let (namespace, local) = reader.resolver().resolve_element(element.name());
                    if namespace_is_word(&namespace) {
                        let index = match local.as_ref() {
                            b"commentRangeStart" => Some(0),
                            b"commentRangeEnd" => Some(1),
                            b"commentReference" => Some(2),
                            _ => None,
                        };
                        if let Some(index) = index {
                            let (_, value) = resolved_element_attribute(
                                &element,
                                reader.resolver(),
                                b"id",
                                AttributeNamespace::Word,
                            )?
                            .ok_or_else(|| {
                                Error::Other("REF f annotation graph edge has no ID".into())
                            })?;
                            let id = value.parse::<i32>().map_err(|_| {
                                Error::Other("REF f annotation graph edge has an invalid ID".into())
                            })?;
                            counts.entry(id).or_insert([0; 3])[index] += 1;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    Ok(counts)
}

fn ref_comment_replacement_ids(
    document: &Document,
    field: &Field,
    source: &RefBookmarkCache,
) -> Result<Vec<Option<i32>>> {
    let mut old = RefBookmarkCache {
        runs: Vec::new(),
        comment_ranges: Vec::new(),
    };
    let mut run = CT_R::new("");
    run.content = vec![RunContent::Field(field.clone())];
    append_ref_cached_run(&run, &mut old)?;
    let source_ids = ref_comment_reference_ids(source)?;
    let old_ids = ref_comment_reference_ids(&old)?;
    if old_ids.is_empty() {
        return Ok(vec![None; source_ids.len()]);
    }
    if old_ids.len() != source_ids.len() {
        return Err(Error::Other(
            "REF f cached annotation reference inventory differs from its current source".into(),
        ));
    }
    let counts = ref_comment_graph_counts(&document.package)?;
    let mut unique = HashSet::new();
    let mut replacements = Vec::new();
    for (id, source_id) in old_ids.iter().zip(&source_ids) {
        let start = old.comment_ranges.iter().filter(|marker| matches!(marker, CommentRangeMarker::Start {id: marker_id,..} if marker_id==id)).count();
        let end = old.comment_ranges.iter().filter(|marker| matches!(marker, CommentRangeMarker::End {id: marker_id,..} if marker_id==id)).count();
        if id == source_id
            || !unique.insert(*id)
            || (start, end) != (1, 1)
            || counts.get(id) != Some(&[1, 1, 1])
        {
            return Err(Error::Other("REF f cached annotation owner is borrowed, shared or has ambiguous range endpoints".into()));
        }
        document.fragment_comment_dependency(*id)?;
        replacements.push(Some(*id));
    }
    Ok(replacements)
}

fn prepare_ref_note_copy(
    document: &mut Document,
    original: &Document,
    kind: StoryKind,
    old_id: i32,
    replace_id: Option<i32>,
    occupied: &mut HashMap<StoryKind, HashSet<i32>>,
) -> Result<(i32, FragmentNoteCopy)> {
    let (source_part, source) = original.fragment_note_dependency(kind, old_id)?;
    if !fragment_note_references(&source)?.is_empty() {
        return Err(Error::Other(
            "REF f note payload contains recursive note references".into(),
        ));
    }
    let part = document.ensure_fragment_note_part_staged(kind)?;
    if part != source_part {
        return Err(Error::Other(
            "REF f note copy has no same-owner relationship scope".into(),
        ));
    }
    let used = if let Some(used) = occupied.get_mut(&kind) {
        used
    } else {
        let notes = CT_Footnotes::from_xml(
            document
                .package
                .get_part(&part)
                .ok_or_else(|| Error::Other("REF f note part is missing".into()))?,
        )?;
        occupied
            .entry(kind)
            .or_insert_with(|| notes.footnotes.iter().map(|note| note.id).collect())
    };
    let id = if let Some(id) = replace_id {
        let (old_part, _) = original.fragment_note_dependency(kind, id)?;
        if old_part != part || !used.contains(&id) {
            return Err(Error::Other(
                "REF f cached note replacement is not a unique existing owner".into(),
            ));
        }
        id
    } else {
        let mut id = 1i32;
        while used.contains(&id) {
            id = id
                .checked_add(1)
                .ok_or_else(|| Error::Other("REF f note IDs are exhausted".into()))?;
        }
        used.insert(id);
        id
    };
    let mut xml = transform_ref_note_copy(original, &source_part, &source)?;
    for relationship_id in relationship_ids_in_xml(&xml)? {
        let relationships = original
            .package
            .get_part_rels(&source_part)
            .ok_or_else(|| Error::Other("REF f note relationship owner is missing".into()))?;
        let relationship = relationships.get_by_id(&relationship_id).ok_or_else(|| {
            Error::Other(format!(
                "REF f note relationship {relationship_id} is missing"
            ))
        })?;
        if !crate::document::relationship_is_internal(relationship) {
            return Err(Error::Other(
                "REF f retained note payload has an external relationship".into(),
            ));
        }
        let target = OpcPackage::resolve_rel_target(&source_part, &relationship.target);
        discover_fragment_part_closure(&original.package, &target, &mut HashSet::new())?;
    }
    let wrapper = format!("<w:body xmlns:w=\"{W_NS}\">").into_bytes();
    let mut body = wrapper.clone();
    body.extend_from_slice(&xml);
    body.extend_from_slice(b"</w:body>");
    let values = body_identity_values(&body)?;
    let mut remap = BodyIdentityRemap::default();
    for old in values.drawing_ids {
        if remap.drawing_ids.contains_key(&old) {
            return Err(Error::Other(
                "REF f source has duplicate drawing IDs".into(),
            ));
        }
        remap
            .drawing_ids
            .insert(old, document.identifiers.reserve_drawing_id()?.to_string());
    }
    let mapped = patch_body_identity_attributes(&body, &remap)?;
    xml = mapped[wrapper.len()..mapped.len() - b"</w:body>".len()].to_vec();
    let maps = HashMap::from([(kind, BTreeMap::from([(old_id.to_string(), id.to_string())]))]);
    xml = patch_fragment_note_ids(&xml, &maps)?;
    Ok((
        id,
        FragmentNoteCopy {
            kind,
            destination_part: part,
            xml,
            replace_id,
        },
    ))
}

fn transform_ref_note_copy(document: &Document, part: &str, source: &[u8]) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<bool>::new();
    let mut edits = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("REF f note XML is malformed: {error}")))?;
        let word = namespace_is_word(&namespace);
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if word && matches!(local.as_ref(), b"bookmarkStart" | b"bookmarkEnd") {
                    reader
                        .read_to_end_into(element.name(), &mut Vec::new())
                        .map_err(|error| {
                            Error::Other(format!("REF f note bookmark is malformed: {error}"))
                        })?;
                    edits.push(start..reader.buffer_position() as usize);
                } else {
                    let flatten = if word && local.as_ref() == b"hyperlink" {
                        let relationship_id = resolved_element_attribute(&element, reader.resolver(), b"id", AttributeNamespace::Relationship)?.map(|(_, value)| value).ok_or_else(|| Error::Other("REF f internal note hyperlink is unsupported without native controls".into()))?;
                        let relationship = document
                            .package
                            .get_part_rels(part)
                            .and_then(|relationships| relationships.get_by_id(&relationship_id))
                            .ok_or_else(|| {
                                Error::Other("REF f note hyperlink relationship is missing".into())
                            })?;
                        if crate::document::relationship_is_internal(relationship) {
                            return Err(Error::Other("REF f internal note hyperlink is unsupported without native controls".into()));
                        }
                        edits.push(start..end);
                        true
                    } else {
                        false
                    };
                    stack.push(flatten);
                }
            }
            Event::Empty(element) => {
                if word
                    && matches!(
                        element.local_name().as_ref(),
                        b"bookmarkStart" | b"bookmarkEnd"
                    )
                {
                    edits.push(start..end);
                }
            }
            Event::End(_) => {
                if stack
                    .pop()
                    .ok_or_else(|| Error::Other("REF f note XML is unbalanced".into()))?
                {
                    edits.push(start..end);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(Error::Other("REF f note XML is unclosed".into()));
    }
    let mut result = source.to_vec();
    for range in edits.into_iter().rev() {
        result.drain(range);
    }
    Ok(result)
}

struct RefCopyField {
    field: Field,
    locked: bool,
    generated: bool,
    suppress_copied_references: bool,
}

fn collect_ref_copy_fields(
    field: &Field,
    inherited_lock: bool,
    generated: bool,
    suppress_copied_references: bool,
    fields: &mut Vec<RefCopyField>,
) {
    let locked = inherited_lock || field.locked() == Some(true);
    fields.push(RefCopyField {
        field: field.clone(),
        locked,
        generated,
        suppress_copied_references,
    });
    for child in field.nested_fields_in_source_order() {
        collect_ref_copy_fields(child, locked, generated, suppress_copied_references, fields);
    }
    for child in field.cached_fields_in_source_order() {
        collect_ref_copy_fields(child, locked, true, suppress_copied_references, fields);
    }
}

fn append_ref_copy_paragraph(
    paragraph: &CT_P,
    suppress_copied_references: bool,
    fields: &mut Vec<RefCopyField>,
) {
    for content in accepted_toc_runs(paragraph)
        .into_iter()
        .flat_map(|run| &run.run.content)
    {
        if let RunContent::Field(field) = content {
            collect_ref_copy_fields(field, false, false, suppress_copied_references, fields);
        }
    }
}

struct RefBookmarkCache {
    runs: Vec<CT_R>,
    comment_ranges: Vec<CommentRangeMarker>,
}

fn ref_cache_marker_at(marker: &CommentRangeMarker, position: usize) -> Result<CommentRangeMarker> {
    match marker {
        CommentRangeMarker::Start {
            id,
            has_child_content: false,
            ..
        } => Ok(CommentRangeMarker::Start {
            id: *id,
            run_index: position,
            raw_before: 0,
            has_child_content: false,
        }),
        CommentRangeMarker::End {
            id,
            has_child_content: false,
            ..
        } => Ok(CommentRangeMarker::End {
            id: *id,
            run_index: position,
            raw_before: 0,
            has_child_content: false,
        }),
        _ => Err(Error::Other(
            "REF f annotation boundary has unmodeled child content".into(),
        )),
    }
}

fn append_ref_cached_run(run: &CT_R, cache: &mut RefBookmarkCache) -> Result<()> {
    if let [RunContent::Field(field)] = run.content.as_slice() {
        if let Some(runs) = field.cached_result_runs() {
            let base = cache.runs.len();
            let mut boundaries = Vec::with_capacity(runs.len() + 1);
            for child in runs {
                boundaries.push(cache.runs.len());
                let mut child = child.clone();
                if child.properties.is_none() {
                    child.properties = run.properties.clone();
                }
                append_ref_cached_run(&child, cache)?;
            }
            boundaries.push(cache.runs.len());
            for marker in field.cached_result_comment_ranges() {
                let position = *boundaries
                    .get(match marker {
                        CommentRangeMarker::Start { run_index, .. }
                        | CommentRangeMarker::End { run_index, .. } => *run_index,
                    })
                    .ok_or_else(|| {
                        Error::Other("REF f cached annotation boundary is stale".into())
                    })?;
                cache
                    .comment_ranges
                    .push(ref_cache_marker_at(marker, position)?);
            }
            if cache.runs.len() == base {
                cache.runs.push(CT_R::new(""));
            }
        } else {
            let segments = field.cached_display_segments();
            if segments.is_empty() {
                let mut stored = CT_R::new(&field.cached_result);
                stored.properties = run.properties.clone();
                cache.runs.push(stored);
            } else {
                for (text, properties) in segments {
                    let mut stored = CT_R::new(text);
                    stored.properties = properties.cloned().or_else(|| run.properties.clone());
                    cache.runs.push(stored);
                }
            }
        }
    } else if run
        .content
        .iter()
        .any(|content| matches!(content, RunContent::Field(_)))
    {
        return Err(Error::Other(
            "REF f mixed field target lacks a source-qualified cache boundary".into(),
        ));
    } else {
        cache.runs.push(run.clone());
    }
    Ok(())
}

fn append_ref_target_paragraph(
    paragraph: &CT_P,
    start: usize,
    end: usize,
    cache: &mut RefBookmarkCache,
) -> Result<()> {
    let projected = paragraph.accepted_bookmark_runs();
    let selected = projected
        .get(start..end)
        .ok_or_else(|| Error::Other("REF f target run range is stale".into()))?;
    let mut boundaries = Vec::with_capacity(selected.len() + 1);
    for run in selected {
        boundaries.push(cache.runs.len());
        append_ref_cached_run(run, cache)?;
    }
    boundaries.push(cache.runs.len());
    if !paragraph.comment_ranges.is_empty() {
        if !paragraph.content_controls.is_empty()
            || !paragraph.revisions.is_empty()
            || projected.len() != paragraph.runs.len()
        {
            return Err(Error::Other(
                "REF f annotation target lacks an exact accepted boundary projection".into(),
            ));
        }
        let mut starts = BTreeMap::new();
        let mut pairs = Vec::new();
        for marker in &paragraph.comment_ranges {
            match marker {
                CommentRangeMarker::Start {
                    id,
                    run_index,
                    has_child_content: false,
                    ..
                } => {
                    if starts.insert(*id, *run_index).is_some() {
                        return Err(Error::Other(
                            "REF f annotation source range is ambiguous".into(),
                        ));
                    }
                }
                CommentRangeMarker::End {
                    id,
                    run_index,
                    has_child_content: false,
                    ..
                } => {
                    let first = starts.remove(id).ok_or_else(|| {
                        Error::Other("REF f annotation source range is unpaired".into())
                    })?;
                    if first > *run_index {
                        return Err(Error::Other(
                            "REF f annotation source range is reversed".into(),
                        ));
                    }
                    pairs.push((*id, first, *run_index));
                }
                _ => {
                    return Err(Error::Other(
                        "REF f annotation source range contains unmodeled payload".into(),
                    ));
                }
            }
        }
        if !starts.is_empty() {
            return Err(Error::Other(
                "REF f annotation source range crosses an unqualified target paragraph".into(),
            ));
        }
        for (id, first, last) in pairs {
            let clipped_start = first.max(start);
            let clipped_end = last.min(end);
            if clipped_start >= clipped_end {
                continue;
            }
            let first = boundaries[clipped_start - start];
            let last = boundaries[clipped_end - start];
            cache.comment_ranges.push(CommentRangeMarker::Start {
                id,
                run_index: first,
                raw_before: 0,
                has_child_content: false,
            });
            cache.comment_ranges.push(CommentRangeMarker::End {
                id,
                run_index: last,
                raw_before: 0,
                has_child_content: false,
            });
            if !cache.runs.iter().flat_map(|run| &run.content).any(|content| matches!(content,RunContent::CommentReference{id:reference,..} if *reference==id)) {
                let mut reference=CT_R::new("");reference.content=vec![RunContent::CommentReference{id,raw_before:0}];cache.runs.push(reference);
            }
        }
    }
    cache.comment_ranges.sort_by_key(|marker| match marker {
        CommentRangeMarker::Start { run_index, .. } | CommentRangeMarker::End { run_index, .. } => {
            *run_index
        }
    });
    Ok(())
}

fn ref_bookmark_runs(document: &Document, target: &str) -> Result<Option<RefBookmarkCache>> {
    let bookmarks = document.bookmarks();
    let matching = bookmarks
        .iter()
        .filter(|bookmark| bookmark.name() == Some(target))
        .collect::<Vec<_>>();
    if !matching.is_empty() {
        if matching.len() != 1 || matching[0].issue().is_some() {
            return Err(Error::Other(
                "REF f target bookmark is ambiguous or malformed".into(),
            ));
        }
        let range = matching[0]
            .range()
            .ok_or_else(|| Error::Other("REF f target lacks a range".into()))?;
        let mut paragraphs = Vec::new();
        collect_body_paragraphs(&document.document.body, &mut paragraphs);
        let mut cache = RefBookmarkCache {
            runs: Vec::new(),
            comment_ranges: Vec::new(),
        };
        for index in range.start.body_index..=range.end.body_index {
            let paragraph = paragraphs
                .get(index)
                .ok_or_else(|| Error::Other("REF f target paragraph is absent".into()))?;
            let projected = paragraph.accepted_bookmark_runs();
            let start = if index == range.start.body_index {
                range.start.run_index
            } else {
                0
            };
            let end = if index == range.end.body_index {
                range.end.run_index
            } else {
                projected.len()
            };
            if index > range.start.body_index {
                let mut line = CT_R::new("");
                line.content = vec![RunContent::Break(rdocx_oxml::text::BreakType::Line)];
                cache.runs.push(line);
            }
            append_ref_target_paragraph(paragraph, start, end, &mut cache)?;
        }
        return Ok(Some(cache));
    }
    let (ranges, across) = document.story_range_inventory()?;
    if across
        .iter()
        .any(|kind| matches!(kind, crate::StoryRangeKind::Bookmark { name, .. } if name == target))
    {
        return Err(Error::Other(
            "REF f target bookmark runs across story owners".into(),
        ));
    }
    let range = ranges.into_iter().find(|range| matches!(range.kind(), crate::StoryRangeKind::Bookmark { name, .. } if name == target));
    let Some(range) = range else {
        return Ok(None);
    };
    let crate::StoryRangeKind::Bookmark { id, .. } = range.kind() else {
        unreachable!()
    };
    let mut selected = false;
    let mut cache = RefBookmarkCache {
        runs: Vec::new(),
        comment_ranges: Vec::new(),
    };
    for (location, xml) in document.story_range_paragraphs()? {
        if location == range.range().start.location {
            selected = true;
        }
        if !selected {
            continue;
        }
        if location.story() != range.range().start.location.story() {
            return Err(Error::Other("REF f target crosses story owners".into()));
        }
        let paragraph = CT_P::from_xml_fragment(&xml)?;
        let runs = paragraph.accepted_bookmark_runs();
        let first = if location == range.range().start.location {
            paragraph
                .bookmark_markers
                .iter()
                .find(|marker| {
                    marker.is_start() && marker.id() == Some(*id) && marker.name() == Some(target)
                })
                .ok_or_else(|| {
                    Error::Other(
                        "REF f target start marker is absent from its physical projection".into(),
                    )
                })?
                .projected_run_index()
        } else {
            0
        };
        let last = if location == range.range().end.location {
            paragraph
                .bookmark_markers
                .iter()
                .find(|marker| !marker.is_start() && marker.id() == Some(*id))
                .ok_or_else(|| {
                    Error::Other(
                        "REF f target end marker is absent from its physical projection".into(),
                    )
                })?
                .projected_run_index()
        } else {
            runs.len()
        };
        if !cache.runs.is_empty() {
            let mut line = CT_R::new("");
            line.content = vec![RunContent::Break(rdocx_oxml::text::BreakType::Line)];
            cache.runs.push(line);
        }
        append_ref_target_paragraph(&paragraph, first, last, &mut cache)?;
        if location == range.range().end.location {
            return Ok(Some(cache));
        }
    }
    Err(Error::Other("REF f related target owner is absent".into()))
}

struct CachedFieldUpdate {
    cached_result: String,
    dirty: bool,
    typed_runs: Option<Vec<CT_R>>,
    comment_ranges: Vec<CommentRangeMarker>,
}

fn valid_xml_character(value: char) -> bool {
    matches!(value, '\u{0009}' | '\u{000A}' | '\u{000D}')
        || ('\u{0020}'..='\u{D7FF}').contains(&value)
        || ('\u{E000}'..='\u{FFFD}').contains(&value)
        || ('\u{10000}'..='\u{10FFFF}').contains(&value)
}

#[derive(Debug, Clone, Copy)]
struct MailMergeStoryState {
    record_number: Option<u32>,
    sequence_number: Option<u32>,
}

struct Evaluator<'a> {
    document: &'a Document,
    context: &'a FieldEvaluationContext,
    bookmarks: BTreeMap<String, BookmarkValue>,
    // Related-story bookmarks whose two markers are in two owners. A REF to
    // one refuses by name instead of reading as missing.
    cross_owner_bookmarks: HashSet<String>,
    numbering_layout: Option<Arc<rdocx_layout::WordLayoutResult>>,
    numbering_nodes: HashMap<rdocx_layout::WordSourcePath, rdocx_layout::SourceNodeId>,
    navigation_styles: HashMap<String, Vec<rdocx_layout::SourceNodeId>>,
    navigation_style_pages: HashMap<String, BTreeMap<usize, Vec<rdocx_layout::SourceNodeId>>>,
    navigation_field_pages: HashMap<oxml_layout::FieldSource, BTreeSet<usize>>,
    results: Vec<FieldEvaluation>,
    visible_results: Vec<bool>,
    general_field_contexts: Vec<bool>,
    general_context: bool,
    mail_merge_stories: BTreeMap<String, MailMergeStoryState>,
    nested_outcomes: Vec<BTreeMap<usize, FieldOutcome>>,
    missing_merge_fields_as_empty: bool,
    current_field_source: Option<oxml_layout::FieldSource>,
    sequence_snapshot: Option<Arc<rdocx_layout::WordSequenceSnapshot>>,
    sequence_nodes: Vec<Option<rdocx_layout::SourceNodeId>>,
    field_sources: HashMap<usize, oxml_layout::FieldSource>,
    ref_note_counts: Option<HashMap<(StoryKind, i32), usize>>,
    locked_fields: HashSet<usize>,
}

struct BookmarkValue {
    text: String,
    range: Option<crate::RunRange>,
}

fn field_needs_numbering_layout(field: &Field) -> bool {
    let instruction = field.effective_instruction();
    let own = unsupported_switch(&instruction).is_none()
        && validate_instruction_shape(&instruction).is_ok()
        && (instruction.name == "STYLEREF"
            || (instruction.name == "REF"
                && ["n", "r", "w"]
                    .iter()
                    .any(|name| has_switch(&instruction, name))));
    own || field
        .effective_nested_fields_in_source_order(&instruction)
        .into_iter()
        .any(field_needs_numbering_layout)
}

impl<'a> Evaluator<'a> {
    fn new(document: &'a Document, context: &'a FieldEvaluationContext) -> Self {
        let mut bookmarks: BTreeMap<String, BookmarkValue> = document
            .bookmarks()
            .into_iter()
            .filter(|bookmark| bookmark.issue().is_none())
            .filter_map(|bookmark| {
                Some((
                    bookmark.name()?.to_owned(),
                    BookmarkValue {
                        text: bookmark.text().to_owned(),
                        range: bookmark.range(),
                    },
                ))
            })
            .collect();
        // Main bookmark ranges keep their established accepted projection.
        // Related ranges are admitted only by the checked physical inventory.
        // One paragraph inventory serves the ranges and the name counts.
        let paragraphs = document.story_range_paragraphs().ok();
        let mut cross_owner_bookmarks = HashSet::new();
        if let Some(Ok((ranges, across))) = paragraphs.as_deref().map(Document::story_ranges_in) {
            for kind in across {
                if let crate::StoryRangeKind::Bookmark { name, .. } = kind
                    && !bookmarks.contains_key(&name)
                {
                    cross_owner_bookmarks.insert(name);
                }
            }
            for range in ranges {
                let crate::StoryRangeKind::Bookmark { name, .. } = range.kind() else {
                    continue;
                };
                if bookmarks.contains_key(name)
                    || range.range().start.location.story().kind() == StoryKind::Body
                {
                    continue;
                }
                if let Ok(Some(cache)) = ref_bookmark_runs(document, name) {
                    bookmarks.insert(
                        name.clone(),
                        BookmarkValue {
                            text: cache.runs.iter().map(CT_R::text).collect(),
                            range: None,
                        },
                    );
                }
            }
        }
        if let Some(paragraphs) = paragraphs {
            let mut names = HashMap::<String, usize>::new();
            for (_, xml) in paragraphs {
                if let Ok(paragraph) = CT_P::from_xml_fragment(&xml) {
                    for marker in paragraph
                        .bookmark_markers
                        .iter()
                        .filter(|marker| marker.is_start())
                    {
                        if let Some(name) = marker.name() {
                            *names.entry(name.to_owned()).or_default() += 1;
                        }
                    }
                }
            }
            bookmarks.retain(|name, _| names.get(name).is_some_and(|count| *count == 1));
        }
        Self {
            document,
            context,
            bookmarks,
            cross_owner_bookmarks,
            numbering_layout: None,
            numbering_nodes: HashMap::new(),
            navigation_styles: HashMap::new(),
            navigation_style_pages: HashMap::new(),
            navigation_field_pages: HashMap::new(),
            results: Vec::new(),
            visible_results: Vec::new(),
            general_field_contexts: Vec::new(),
            general_context: true,
            mail_merge_stories: BTreeMap::new(),
            nested_outcomes: Vec::new(),
            missing_merge_fields_as_empty: false,
            current_field_source: None,
            sequence_snapshot: None,
            ref_note_counts: None,
            locked_fields: HashSet::new(),
            sequence_nodes: Vec::new(),
            field_sources: HashMap::new(),
        }
    }

    fn for_mail_merge(document: &'a Document, context: &'a FieldEvaluationContext) -> Self {
        let mut evaluator = Self::new(document, context);
        evaluator.missing_merge_fields_as_empty = true;
        evaluator
    }

    fn refresh_main_bookmark_sequence_text(&mut self, paragraphs: &[&CT_P]) -> Result<()> {
        let Some(snapshot) = &self.sequence_snapshot else {
            return Ok(());
        };
        if self.sequence_nodes.len() != paragraphs.len() {
            return Err(Error::Other(
                "main bookmark sequence source inventory is incomplete".into(),
            ));
        }
        let mut displays = Vec::with_capacity(paragraphs.len());
        for (paragraph, node) in paragraphs.iter().zip(&self.sequence_nodes) {
            let mut fields = Vec::new();
            for run in accepted_toc_runs(paragraph) {
                for content in &run.run.content {
                    if let RunContent::Field(field) = content {
                        collect_preorder_fields(field, false, &mut fields);
                    }
                }
            }
            let mut physical_fields = Vec::new();
            for run in paragraph.source_runs() {
                for content in &run.content {
                    if let RunContent::Field(field) = content {
                        collect_preorder_fields(field, false, &mut physical_fields);
                    }
                }
            }
            let physical_indices = physical_fields
                .iter()
                .enumerate()
                .map(|(index, (field, _))| (std::ptr::from_ref(*field), index as u32))
                .collect::<HashMap<_, _>>();
            let updates = fields
                .iter()
                .map(|(field, _)| {
                    match node.and_then(|node| {
                        snapshot.field_value(oxml_layout::FieldSource {
                            node,
                            index: *physical_indices.get(&std::ptr::from_ref(*field))?,
                        })
                    }) {
                        Some(Ok(value)) => Some(CachedFieldUpdate {
                            typed_runs: None,
                            comment_ranges: Vec::new(),
                            cached_result: value.to_owned(),
                            dirty: false,
                        }),
                        _ => None,
                    }
                })
                .collect::<Vec<_>>();
            let mut projection = (*paragraph).clone();
            let mut consumed = 0;
            apply_updates_to_paragraph(&mut projection, &updates, &mut consumed)?;
            if consumed != updates.len() {
                return Err(Error::Other(
                    "bookmark sequence field traversal is inconsistent".into(),
                ));
            }
            displays.push(
                projection
                    .accepted_bookmark_runs()
                    .into_iter()
                    .map(|run| {
                        let mut displayed = run.clone();
                        for content in &mut displayed.content {
                            if let RunContent::Field(field) = content
                                && field.effective_instruction().name == "SEQ"
                            {
                                *content = RunContent::Text(rdocx_oxml::text::CT_Text::new(
                                    &field.cached_result,
                                ));
                            }
                        }
                        displayed.text()
                    })
                    .collect::<Vec<_>>(),
            );
        }
        for bookmark in self.bookmarks.values_mut() {
            let Some(range) = bookmark.range else {
                continue;
            };
            let mut parts = Vec::new();
            for index in range.start.body_index..=range.end.body_index {
                let Some(runs) = displays.get(index) else {
                    return Err(Error::Other(
                        "bookmark sequence range owner is missing".into(),
                    ));
                };
                let start = if index == range.start.body_index {
                    range.start.run_index
                } else {
                    0
                };
                let end = if index == range.end.body_index {
                    range.end.run_index
                } else {
                    runs.len()
                };
                let Some(slice) = runs.get(start..end) else {
                    return Err(Error::Other("bookmark sequence range is stale".into()));
                };
                parts.push(slice.concat());
            }
            bookmark.text = parts.join("\n");
        }
        Ok(())
    }

    fn evaluate_story(&mut self, story: &str, paragraphs: &[&CT_P]) -> Result<()> {
        if self.numbering_layout.is_none()
            && paragraphs.iter().any(|paragraph| {
                paragraph.runs().into_iter().any(|run| {
                    run.content.iter().any(|content| match content {
                        RunContent::Field(field) => field_needs_numbering_layout(field),
                        _ => false,
                    })
                })
            })
        {
            self.prepare_numbering_layout()?;
        }
        if self.sequence_snapshot.is_some() && self.sequence_nodes.len() != paragraphs.len() {
            return Err(Error::Other(format!(
                "source inventory identified {} of {} paragraphs in {story}",
                self.sequence_nodes.len(),
                paragraphs.len()
            )));
        }
        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            self.general_context = self
                .general_field_contexts
                .get(paragraph_index)
                .copied()
                .unwrap_or(true);
            self.register_paragraph_fields(
                paragraph,
                self.sequence_nodes.get(paragraph_index).copied().flatten(),
            );
            for run in accepted_toc_runs(paragraph) {
                for content in &run.run.content {
                    if let RunContent::Field(field) = content {
                        self.evaluate_field(field, story, paragraphs, paragraph_index);
                    }
                }
            }
        }
        Ok(())
    }

    fn register_paragraph_fields(
        &mut self,
        paragraph: &CT_P,
        node: Option<rdocx_layout::SourceNodeId>,
    ) {
        self.field_sources.clear();
        self.locked_fields.clear();
        let mut fields = Vec::new();
        for run in paragraph.source_runs() {
            for content in &run.content {
                if let RunContent::Field(field) = content {
                    collect_preorder_fields(field, false, &mut fields);
                }
            }
        }
        self.locked_fields.extend(
            fields
                .iter()
                .filter(|(_, locked)| *locked)
                .map(|(field, _)| std::ptr::from_ref(*field) as usize),
        );
        if let Some(node) = node {
            self.field_sources
                .extend(fields.into_iter().enumerate().map(|(index, (field, _))| {
                    (
                        std::ptr::from_ref(field) as usize,
                        oxml_layout::FieldSource {
                            node,
                            index: index as u32,
                        },
                    )
                }));
        }
    }

    fn prepare_numbering_layout(&mut self) -> Result<()> {
        let layout = self.document.layout_deterministic()?;
        self.bind_numbering_layout(layout);
        Ok(())
    }

    fn bind_numbering_layout(&mut self, layout: Arc<rdocx_layout::WordLayoutResult>) {
        self.numbering_nodes.clear();
        for index in 1.. {
            let Some(node) = rdocx_layout::SourceNodeId::new(index) else {
                break;
            };
            let Some(path) = layout.source_node(node) else {
                break;
            };
            self.numbering_nodes.insert(path.clone(), node);
        }
        self.navigation_styles.clear();
        self.navigation_style_pages.clear();
        self.navigation_field_pages.clear();
        let mut aliases: HashMap<&str, Vec<String>> = HashMap::new();
        for style in &self.document.styles.styles {
            let names = aliases.entry(&style.style_id).or_default();
            names.push(style.style_id.to_lowercase());
            if let Some(name) = &style.name {
                names.push(name.to_lowercase());
            }
            names.sort();
            names.dedup();
        }
        for (node, path, _, _, style, _, page) in layout.paragraph_navigation() {
            if path.story != rdocx_layout::WordStory::Document {
                continue;
            }
            for alias in style
                .and_then(|style| aliases.get(style))
                .into_iter()
                .flatten()
            {
                self.navigation_styles
                    .entry(alias.clone())
                    .or_default()
                    .push(node);
                if let Some(page) = page {
                    self.navigation_style_pages
                        .entry(alias.clone())
                        .or_default()
                        .entry(page)
                        .or_default()
                        .push(node);
                }
            }
        }
        for nodes in self.navigation_styles.values_mut() {
            nodes.sort_unstable_by_key(|node| node.get());
            nodes.dedup();
        }
        for placement in layout.field_placements() {
            self.navigation_field_pages
                .entry(placement.source)
                .or_default()
                .insert(placement.physical_page);
        }
        self.numbering_layout = Some(layout);
    }

    fn numbering_source(
        &self,
        node: rdocx_layout::SourceNodeId,
    ) -> Option<rdocx_layout::SourceNodeId> {
        let path = self.sequence_snapshot.as_ref()?.source_node(node)?;
        self.numbering_nodes.get(path).copied()
    }

    fn ensure_numbering_layout_for_field(&mut self, field: &Field) -> Result<()> {
        if self.numbering_layout.is_none() && field_needs_numbering_layout(field) {
            self.prepare_numbering_layout()?;
        }
        Ok(())
    }

    fn evaluate_field(
        &mut self,
        field: &Field,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let instruction = field.effective_instruction();
        let previous_source = self.current_field_source;
        self.current_field_source = self
            .field_sources
            .get(&(std::ptr::from_ref(field) as usize))
            .copied();
        let result_index = self.results.len();
        let visible = self.general_context || matches!(instruction.name.as_str(), "SEQ" | "REF");
        self.visible_results.push(visible);
        self.results.push(FieldEvaluation {
            field_index: result_index,
            instruction: instruction.raw.clone(),
            cached_result: field.cached_result.clone(),
            outcome: keep("field evaluation did not complete"),
        });

        self.nested_outcomes.push(BTreeMap::new());
        self.evaluate_nested_fields(field, &instruction, story, paragraphs, paragraph_index);
        let mut outcome = if !visible {
            keep("ordinary field is outside its established evaluation context")
        } else if self
            .locked_fields
            .contains(&(std::ptr::from_ref(field) as usize))
        {
            keep("locked field retains its stored display")
        } else if instruction.name == "SEQ" && self.sequence_snapshot.is_some() {
            match self
                .field_sources
                .get(&(std::ptr::from_ref(field) as usize))
                .copied()
                .and_then(|source| self.sequence_snapshot.as_ref()?.field_value(source))
            {
                Some(Ok(value)) => FieldOutcome::Resolved(value.to_owned()),
                Some(Err(message)) => keep(message),
                None => keep("SEQ lacks a unique physical accepted source"),
            }
        } else if self
            .locked_fields
            .contains(&(std::ptr::from_ref(field) as usize))
        {
            keep("locked field retains its stored display")
        } else {
            self.evaluate_instruction(&instruction, story, paragraphs, paragraph_index)
        };
        if story == "main"
            && instruction.name == "REF"
            && has_switch(&instruction, "f")
            && matches!(outcome, FieldOutcome::Resolved(_))
        {
            let mut old_references = Vec::new();
            collect_ref_cached_note_references(field, &mut old_references);
            if !old_references.is_empty() {
                let qualification = (|| -> Result<()> {
                    if self.ref_note_counts.is_none() {
                        self.ref_note_counts =
                            Some(ref_note_reference_counts(&self.document.package)?);
                    }
                    let target = text_argument(&instruction, 0)
                        .ok_or_else(|| Error::Other("REF f target is absent".into()))?;
                    let cache = ref_bookmark_runs(self.document, target)?
                        .ok_or_else(|| Error::Other("REF f target is missing".into()))?;
                    let new_references = cache
                        .runs
                        .iter()
                        .flat_map(|run| &run.content)
                        .filter_map(|content| match content {
                            RunContent::FootnoteRef { id, .. } => Some((StoryKind::Footnote, *id)),
                            RunContent::EndnoteRef { id, .. } => Some((StoryKind::Endnote, *id)),
                            _ => None,
                        })
                        .collect::<Vec<_>>();
                    if old_references.len() != new_references.len() {
                        return Err(Error::Other(
                            "cached note graph differs from the current source reference inventory"
                                .into(),
                        ));
                    }
                    for ((kind, id), (source_kind, source_id)) in
                        old_references.iter().zip(&new_references)
                    {
                        if kind != source_kind
                            || id == source_id
                            || self
                                .ref_note_counts
                                .as_ref()
                                .and_then(|counts| counts.get(&(*kind, *id)))
                                != Some(&1)
                        {
                            return Err(Error::Other("cached note reference ownership is ambiguous, shared or borrowed from its source".into()));
                        }
                        self.document.fragment_note_dependency(*kind, *id)?;
                    }
                    Ok(())
                })();
                if let Err(error) = qualification {
                    outcome = keep(&format!(
                        "REF f repeated note copy retains its cache: {error}"
                    ));
                }
            }
        }
        if story == "main"
            && instruction.name == "REF"
            && has_switch(&instruction, "f")
            && matches!(outcome, FieldOutcome::Resolved(_))
        {
            let qualification = text_argument(&instruction, 0)
                .ok_or_else(|| Error::Other("REF f target is absent".into()))
                .and_then(|target| ref_bookmark_runs(self.document, target))
                .and_then(|cache| {
                    cache.ok_or_else(|| Error::Other("REF f target is absent".into()))
                })
                .and_then(|cache| {
                    let replacements = ref_comment_replacement_ids(self.document, field, &cache)?;
                    let ids = ref_comment_reference_ids(&cache)?;
                    if !ids.is_empty() {
                        let mut staged = self.document.clone_for_staging();
                        let mut occupied = self
                            .document
                            .comments
                            .as_ref()
                            .map(|comments| {
                                comments
                                    .comments
                                    .iter()
                                    .map(|comment| comment.id)
                                    .collect::<HashSet<_>>()
                            })
                            .unwrap_or_default();
                        let mut identities =
                            ref_comment_occupied_identity_values(&self.document.package)?;
                        for (id, replacement) in ids.into_iter().zip(replacements) {
                            prepare_ref_comment_copy(
                                &mut staged,
                                self.document,
                                id,
                                replacement,
                                &mut occupied,
                                &mut identities,
                            )?;
                        }
                    }
                    Ok(())
                });
            if let Err(error) = qualification {
                outcome = keep(&format!(
                    "REF f repeated annotation copy retains its cache: {error}"
                ));
            }
        }
        self.nested_outcomes.pop();
        self.results[result_index].outcome = outcome.clone();
        self.current_field_source = previous_source;
        outcome
    }

    fn evaluate_instruction(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        if instruction.name.is_empty() {
            return keep("field instruction has no name");
        }
        if let Some(name) = unsupported_switch(instruction) {
            return keep(&format!(
                "field {} uses unsupported switch \\{name}",
                instruction.name
            ));
        }
        if let Err(diagnostic) = validate_instruction_shape(instruction) {
            return keep(&diagnostic);
        }

        let outcome = match instruction.name.as_str() {
            "PAGE" | "NUMPAGES" | "SECTION" | "SECTIONPAGES" => FieldOutcome::DeferredPagination,
            "PAGEREF" => self.evaluate_pageref(instruction),
            "REF" => self.evaluate_ref(instruction, story, paragraph_index),
            "IF" => self.evaluate_if(instruction, story, paragraphs, paragraph_index),
            "SEQ" => keep("SEQ requires accepted physical source context"),
            "DOCPROPERTY" => self.evaluate_docproperty(instruction),
            "DOCVARIABLE" => self.evaluate_docvariable(instruction),
            "STYLEREF" => self.evaluate_styleref(instruction, paragraphs, paragraph_index),
            "INCLUDETEXT" => self.evaluate_includetext(instruction),
            "DATE" | "TIME" => self.evaluate_date_time(instruction),
            "FILENAME" => self.evaluate_filename(instruction),
            "AUTHOR" => self.evaluate_author(),
            "MERGEFIELD" => self.evaluate_mergefield(instruction),
            "=" => self.evaluate_formula(instruction, story, paragraphs, paragraph_index),
            "TOC" => self.evaluate_toc(instruction, story, paragraphs, paragraph_index),
            "TC" => self.evaluate_tc(instruction, story, paragraphs, paragraph_index),
            "NEXT" | "NEXTIF" | "SKIPIF" | "MERGEREC" | "MERGESEQ" => {
                self.evaluate_mail_merge_control(instruction, story, paragraphs, paragraph_index)
            }
            "DISPLAYBARCODE" | "MERGEBARCODE" => {
                self.evaluate_barcode(instruction, story, paragraphs, paragraph_index)
            }
            name => keep(&format!("field {name} is unsupported")),
        };

        match outcome {
            FieldOutcome::Resolved(value) => {
                match apply_formats(instruction, &value, self.context.now) {
                    Ok(value) => FieldOutcome::Resolved(value),
                    Err(diagnostic) => keep(&diagnostic),
                }
            }
            other => other,
        }
    }

    fn evaluate_nested_fields(
        &mut self,
        field: &Field,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) {
        // Effective instructions own cloned nested operands. Bind those clones
        // through their actual structured source slots, never instruction text.
        let effective_nested = field.effective_nested_fields_in_source_order(instruction);
        let original_nested = field.nested_fields_in_source_order();
        let mut aliases = Vec::new();
        if effective_nested.len() == original_nested.len() {
            for (original, effective) in original_nested.iter().zip(&effective_nested) {
                let mut original_fields = Vec::new();
                let mut effective_fields = Vec::new();
                collect_preorder_fields(original, false, &mut original_fields);
                collect_preorder_fields(effective, false, &mut effective_fields);
                if original_fields.len() == effective_fields.len() {
                    for ((original, _), (effective, _)) in
                        original_fields.into_iter().zip(effective_fields)
                    {
                        let key = std::ptr::from_ref(effective) as usize;
                        aliases.push(key);
                        if self
                            .locked_fields
                            .contains(&(std::ptr::from_ref(original) as usize))
                        {
                            self.locked_fields.insert(key);
                        } else {
                            self.locked_fields.remove(&key);
                        }
                        if let Some(source) = self
                            .field_sources
                            .get(&(std::ptr::from_ref(original) as usize))
                            .copied()
                        {
                            self.field_sources.insert(key, source);
                        } else {
                            self.field_sources.remove(&key);
                        }
                    }
                }
            }
        }
        for nested in effective_nested
            .into_iter()
            .chain(field.cached_fields_in_source_order())
        {
            let key = std::ptr::from_ref(nested) as usize;
            if !self
                .nested_outcomes
                .last()
                .is_some_and(|outcomes| outcomes.contains_key(&key))
            {
                let outcome = self.evaluate_field(nested, story, paragraphs, paragraph_index);
                self.nested_outcomes
                    .last_mut()
                    .expect("field evaluation frame exists")
                    .insert(key, outcome);
            }
        }
        for key in aliases {
            self.field_sources.remove(&key);
            self.locked_fields.remove(&key);
        }
    }

    fn evaluate_ref(
        &self,
        instruction: &FieldInstruction,
        _story: &str,
        _paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(target) = text_argument(instruction, 0) else {
            return keep("REF requires a bookmark name");
        };
        if has_switch(instruction, "f") {
            return match ref_bookmark_runs(self.document, target) {
                Ok(Some(cache)) => FieldOutcome::Resolved(
                    self.bookmarks
                        .get(target)
                        .map(|bookmark| bookmark.text.clone())
                        .unwrap_or_else(|| cache.runs.iter().map(CT_R::text).collect()),
                ),
                Ok(None) => keep(&format!("REF target {target} was not found")),
                Err(error) => keep(&format!("REF f source could not be resolved: {error}")),
            };
        }
        match self.bookmarks.get(target) {
            Some(bookmark) => {
                let position = if has_switch(instruction, "p") {
                    match self
                        .sequence_snapshot
                        .as_ref()
                        .zip(self.current_field_source)
                        .ok_or_else(|| {
                            "REF relative source lacks an accepted physical binding".to_owned()
                        })
                        .and_then(|(snapshot, source)| {
                            snapshot.bookmark_relative_position(target, source)
                        }) {
                        Ok(position) => Some(position.to_owned()),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                } else {
                    None
                };
                let numbering_switch = if has_switch(instruction, "w") {
                    Some("full")
                } else if has_switch(instruction, "r") {
                    Some("relative")
                } else if has_switch(instruction, "n") {
                    Some("level")
                } else {
                    None
                };
                let Some(_) = numbering_switch else {
                    if let Some(position) = position {
                        return FieldOutcome::Resolved(position);
                    }
                    return FieldOutcome::Resolved(bookmark.text.clone());
                };
                let Some(numbering) = self
                    .numbering_layout
                    .as_ref()
                    .and_then(|layout| layout.bookmark_numbering(target))
                else {
                    let mut value = if self
                        .numbering_layout
                        .as_ref()
                        .is_some_and(|layout| layout.bookmark_numbering_suppressed(target))
                    {
                        "0".into()
                    } else {
                        bookmark.text.clone()
                    };
                    if let Some(position) = position {
                        value.push(' ');
                        value.push_str(&position);
                    }
                    return FieldOutcome::Resolved(value);
                };
                let mut value = match numbering.numbered_reference_text(
                    instruction,
                    self.current_field_source
                        .and_then(|source| self.numbering_source(source.node))
                        .and_then(|node| self.numbering_layout.as_ref()?.paragraph_numbering(node)),
                ) {
                    Ok(value) => value,
                    Err(diagnostic) => return keep(&diagnostic),
                };
                if let Some(position) = position {
                    value.push(' ');
                    value.push_str(&position);
                }
                FieldOutcome::Resolved(value)
            }
            None if self.cross_owner_bookmarks.contains(target) => {
                keep(&format!("REF target {target} runs across story owners"))
            }
            None => keep(&format!("REF target {target} was not found")),
        }
    }

    fn evaluate_pageref(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(target) = text_argument(instruction, 0) else {
            return keep("PAGEREF requires a bookmark name");
        };
        if self.bookmarks.contains_key(target) {
            FieldOutcome::DeferredPagination
        } else {
            keep(&format!("PAGEREF target {target} was not found"))
        }
    }

    fn evaluate_if(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        if instruction.arguments.len() != 5 {
            return keep("IF requires two operands, an operator, and two results");
        }
        let arguments = instruction
            .arguments
            .iter()
            .map(|argument| self.resolve_argument(argument, story, paragraphs, paragraph_index))
            .collect::<Vec<_>>();
        let left = match &arguments[0] {
            Ok(value) => value,
            Err(diagnostic) => return keep(diagnostic),
        };
        let operator = match &arguments[1] {
            Ok(value) => value,
            Err(diagnostic) => return keep(diagnostic),
        };
        let right = match &arguments[2] {
            Ok(value) => value,
            Err(diagnostic) => return keep(diagnostic),
        };
        let Some(condition) = compare_if(left, operator, right) else {
            return keep(&format!("IF operator {operator} is unsupported"));
        };
        let selected = if condition { 3 } else { 4 };
        match &arguments[selected] {
            Ok(value) => FieldOutcome::Resolved(value.clone()),
            Err(diagnostic) => keep(diagnostic),
        }
    }

    fn resolve_argument(
        &mut self,
        argument: &FieldArgument,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> std::result::Result<String, String> {
        match argument {
            FieldArgument::Text(value) => Ok(value.clone()),
            FieldArgument::Nested(field) => {
                let key = std::ptr::from_ref(field.as_ref()) as usize;
                let outcome = if let Some(outcome) = self
                    .nested_outcomes
                    .last()
                    .and_then(|outcomes| outcomes.get(&key))
                {
                    outcome.clone()
                } else {
                    let outcome = self.evaluate_field(field, story, paragraphs, paragraph_index);
                    self.nested_outcomes
                        .last_mut()
                        .expect("field evaluation frame exists")
                        .insert(key, outcome.clone());
                    outcome
                };
                match outcome {
                    FieldOutcome::Resolved(value) => Ok(value),
                    FieldOutcome::DeferredPagination => {
                        Err("nested field requires deferred pagination".to_owned())
                    }
                    FieldOutcome::TableOfContents(_)
                    | FieldOutcome::TableOfContentsEntry(_)
                    | FieldOutcome::MailMergeControl(_)
                    | FieldOutcome::Barcode(_) => {
                        Err("nested field produced a non-text result".to_owned())
                    }
                    FieldOutcome::KeepStored { diagnostic } => {
                        Err(format!("nested field was not resolved: {diagnostic}"))
                    }
                }
            }
        }
    }

    fn evaluate_docproperty(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(name) = text_argument(instruction, 0) else {
            return keep("DOCPROPERTY requires a property name");
        };
        if let Some(value) = self.core_property(name) {
            return FieldOutcome::Resolved(value.to_owned());
        }
        let matches = self
            .document
            .custom_properties
            .iter()
            .flat_map(|properties| &properties.properties)
            .filter(|property| {
                property
                    .name
                    .as_deref()
                    .is_some_and(|candidate| candidate.eq_ignore_ascii_case(name))
            })
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return keep(&format!(
                "DOCPROPERTY target {name} was not found or is ambiguous"
            ));
        }
        match custom_property_text(&matches[0].value) {
            Some(value) => FieldOutcome::Resolved(value),
            None => keep(&format!(
                "DOCPROPERTY target {name} has an unsupported value"
            )),
        }
    }

    fn core_property(&self, name: &str) -> Option<&str> {
        let properties = self.document.core_properties.as_ref()?;
        match normalized_name(name).as_str() {
            "title" => properties.title.as_deref(),
            "subject" => properties.subject.as_deref(),
            "author" | "creator" => properties.creator.as_deref(),
            "comments" | "description" => properties.description.as_deref(),
            "keywords" => properties.keywords.as_deref(),
            "lastsavedby" | "lastmodifiedby" => properties.last_modified_by.as_deref(),
            "createtime" | "created" => properties.created.as_deref(),
            "savedate" | "modified" => properties.modified.as_deref(),
            _ => None,
        }
    }

    fn evaluate_docvariable(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(name) = text_argument(instruction, 0) else {
            return keep("DOCVARIABLE requires a variable name");
        };
        let matches = self
            .document
            .settings
            .iter()
            .flat_map(|settings| settings.document_variables())
            .filter(|variable| variable.name.eq_ignore_ascii_case(name))
            .collect::<Vec<_>>();
        if matches.len() == 1 {
            FieldOutcome::Resolved(matches[0].value.clone())
        } else {
            keep(&format!(
                "DOCVARIABLE target {name} was not found or is ambiguous"
            ))
        }
    }

    fn evaluate_styleref(
        &self,
        instruction: &FieldInstruction,
        _paragraphs: &[&CT_P],
        _paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(target) = text_argument(instruction, 0) else {
            return keep("STYLEREF requires a style name");
        };
        let Some(layout) = &self.numbering_layout else {
            return keep("STYLEREF source layout is unavailable");
        };
        if let Some(failure) = layout.navigation_failure() {
            return keep(&format!(
                "STYLEREF navigation source discovery unavailable: {failure}"
            ));
        }
        let Some(field_source) = self.current_field_source else {
            return keep("STYLEREF field lacks an accepted physical binding");
        };
        let Some(caller) = self.numbering_source(field_source.node) else {
            return keep("STYLEREF field source is unavailable in this layout");
        };
        let caller_path = layout.source_node(caller).expect("registered caller");
        let key = target.to_lowercase();
        let candidates = self.navigation_styles.get(&key);
        let source = if caller_path.story == rdocx_layout::WordStory::Document
            && !has_switch(instruction, "l")
        {
            candidates.and_then(|nodes| {
                let boundary = nodes.partition_point(|node| node.get() < caller.get());
                boundary
                    .checked_sub(1)
                    .and_then(|i| nodes.get(i))
                    .or_else(|| nodes.iter().find(|node| **node != caller))
                    .copied()
            })
        } else {
            let source = oxml_layout::FieldSource {
                node: caller,
                index: field_source.index,
            };
            let Some(pages) = self
                .navigation_field_pages
                .get(&source)
                .filter(|pages| pages.len() == 1)
            else {
                return keep("STYLEREF page selection requires one actual placement");
            };
            let page = *pages.first().expect("one actual placement");
            self.navigation_style_pages.get(&key).and_then(|pages| {
                pages
                    .get(&page)
                    .and_then(|nodes| {
                        if has_switch(instruction, "l") {
                            nodes.iter().rev().find(|node| **node != caller)
                        } else {
                            nodes.iter().find(|node| **node != caller)
                        }
                    })
                    .copied()
                    .or_else(|| {
                        pages
                            .range(..page)
                            .next_back()
                            .and_then(|(_, nodes)| nodes.last())
                            .copied()
                    })
                    .or_else(|| {
                        pages
                            .range((std::ops::Bound::Excluded(page), std::ops::Bound::Unbounded))
                            .next()
                            .and_then(|(_, nodes)| nodes.first())
                            .copied()
                    })
            })
        };
        let Some(node) = source else {
            return keep(&format!("STYLEREF target {target} was not found"));
        };
        let Some((path, _, text, _, numbering, placed)) = layout.paragraph_navigation_source(node)
        else {
            return keep("STYLEREF selected source is unavailable");
        };
        if placed.is_none() {
            return keep("STYLEREF selected source is unplaced");
        }
        let numbered = ["n", "r", "w"]
            .iter()
            .any(|name| has_switch(instruction, name));
        let mut value = if numbered {
            match numbering {
                Some(numbering) => match numbering
                    .numbered_reference_text(instruction, layout.paragraph_numbering(caller))
                {
                    Ok(value) => value,
                    Err(diagnostic) => return keep(&diagnostic),
                },
                None if layout.paragraph_numbering_suppressed(node) => "0".into(),
                None => return keep("STYLEREF source numbering is unavailable"),
            }
        } else {
            text.to_owned()
        };
        if has_switch(instruction, "p") {
            if path.story != caller_path.story {
                return keep("STYLEREF relative position lacks a comparable physical story");
            }
            value.push_str(if path.children < caller_path.children {
                " above"
            } else {
                " below"
            });
        }
        FieldOutcome::Resolved(value)
    }

    fn evaluate_includetext(&self, instruction: &FieldInstruction) -> FieldOutcome {
        if has_switch(instruction, "c") {
            return keep("INCLUDETEXT converter selection is unsupported");
        }
        let Some(source) = text_argument(instruction, 0) else {
            return keep("INCLUDETEXT requires a source name");
        };
        let key = text_argument(instruction, 1)
            .map(|bookmark| format!("{source}#{bookmark}"))
            .unwrap_or_else(|| source.to_owned());
        match self.context.included_text.get(&key) {
            Some(value) => FieldOutcome::Resolved(value.clone()),
            None => keep(&format!("INCLUDETEXT input {key} was not supplied")),
        }
    }

    fn evaluate_date_time(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(now) = self.context.now else {
            return keep(&format!(
                "{} requires an explicit date and time",
                instruction.name
            ));
        };
        if !valid_date_time(now) {
            return keep("field date and time is invalid");
        }
        let default_picture = if instruction.name == "DATE" {
            "M/d/yyyy"
        } else {
            "h:mm:ss AM/PM"
        };
        let picture = switch_text(instruction, "@").unwrap_or(default_picture);
        match format_date_time(now, picture) {
            Ok(value) => FieldOutcome::Resolved(value),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_filename(&self, instruction: &FieldInstruction) -> FieldOutcome {
        if has_switch(instruction, "p") {
            return self
                .context
                .file_path
                .clone()
                .map(FieldOutcome::Resolved)
                .unwrap_or_else(|| keep("FILENAME path was not supplied"));
        }
        self.context
            .file_name
            .clone()
            .or_else(|| {
                self.context
                    .file_path
                    .as_deref()
                    .and_then(lexical_file_name)
                    .map(str::to_owned)
            })
            .map(FieldOutcome::Resolved)
            .unwrap_or_else(|| keep("FILENAME name was not supplied"))
    }

    fn evaluate_author(&self) -> FieldOutcome {
        self.document
            .core_properties
            .as_ref()
            .and_then(|properties| properties.creator.clone())
            .map(FieldOutcome::Resolved)
            .unwrap_or_else(|| keep("AUTHOR metadata was not found"))
    }

    fn evaluate_mergefield(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let value = text_argument(instruction, 0)
            .and_then(|name| self.context.merge_fields.get(name))
            .map(String::as_str);
        match resolve_mergefield_text(instruction, value, self.missing_merge_fields_as_empty) {
            Ok(value) => FieldOutcome::Resolved(value),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_formula(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        if instruction.arguments.is_empty() {
            return keep("formula requires an expression");
        }
        let mut expression = String::new();
        for argument in &instruction.arguments {
            let value = match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                Ok(value) => value,
                Err(diagnostic) => return keep(&diagnostic),
            };
            if !expression.is_empty() {
                expression.push(' ');
            }
            expression.push_str(&value);
            if expression.len() > MAX_FORMULA_BYTES {
                return keep("formula exceeds the 4096-byte limit");
            }
        }
        match FormulaParser::new(&expression).and_then(FormulaParser::parse) {
            Ok(value) => FieldOutcome::Resolved(format_formula_value(value)),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_mail_merge_control(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        match instruction.name.as_str() {
            "MERGESEQ" if self.context.merge_sequence_number == Some(0) => {
                return keep("MERGESEQ merge sequence number must be one-based");
            }
            "NEXT" | "NEXTIF" | "SKIPIF" | "MERGEREC"
                if self.context.merge_record_number == Some(0) =>
            {
                return keep(&format!(
                    "{} merge record number must be one-based",
                    instruction.name
                ));
            }
            _ => {}
        }
        let condition = match instruction.name.as_str() {
            "NEXTIF" | "SKIPIF" => {
                let values = instruction
                    .arguments
                    .iter()
                    .map(|argument| {
                        self.resolve_argument(argument, story, paragraphs, paragraph_index)
                    })
                    .collect::<std::result::Result<Vec<_>, _>>();
                let values = match values {
                    Ok(values) => values,
                    Err(diagnostic) => return keep(&diagnostic),
                };
                let Some(condition) = compare_if(&values[0], &values[1], &values[2]) else {
                    return keep(&format!(
                        "{} operator {} is unsupported",
                        instruction.name, values[1]
                    ));
                };
                condition
            }
            _ => false,
        };
        let state =
            self.mail_merge_stories
                .entry(story.to_owned())
                .or_insert(MailMergeStoryState {
                    record_number: self.context.merge_record_number,
                    sequence_number: self.context.merge_sequence_number,
                });
        match instruction.name.as_str() {
            "NEXT" => {
                let Some(current_record) = state.record_number else {
                    return keep("NEXT requires an explicit merge record number");
                };
                let Some(record_number) = current_record.checked_add(1) else {
                    return keep("NEXT record number overflowed");
                };
                state.record_number = Some(record_number);
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecord { record_number })
            }
            "NEXTIF" => {
                let Some(mut record_number) = state.record_number else {
                    return keep("NEXTIF requires an explicit merge record number");
                };
                if condition {
                    let Some(next_record_number) = record_number.checked_add(1) else {
                        return keep("NEXTIF record number overflowed");
                    };
                    record_number = next_record_number;
                    state.record_number = Some(record_number);
                }
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecordIf {
                    condition,
                    record_number,
                })
            }
            "SKIPIF" => match state.record_number {
                Some(record_number) => {
                    FieldOutcome::MailMergeControl(MailMergeControl::SkipRecordIf {
                        condition,
                        record_number,
                    })
                }
                None => keep("SKIPIF requires an explicit merge record number"),
            },
            "MERGEREC" => state
                .record_number
                .map(MailMergeControl::RecordNumber)
                .map(FieldOutcome::MailMergeControl)
                .unwrap_or_else(|| keep("MERGEREC requires an explicit merge record number")),
            "MERGESEQ" => state
                .sequence_number
                .map(MailMergeControl::SequenceNumber)
                .map(FieldOutcome::MailMergeControl)
                .unwrap_or_else(|| keep("MERGESEQ requires an explicit merge sequence number")),
            _ => unreachable!(),
        }
    }

    fn evaluate_barcode(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(source) = instruction.arguments.first() else {
            return keep(&format!("{} requires a value", instruction.name));
        };
        let value = match source {
            FieldArgument::Text(source) if instruction.name == "MERGEBARCODE" => {
                match self.context.merge_fields.get(source) {
                    Some(value) => value.clone(),
                    None => return keep(&format!("MERGEBARCODE input {source} was not supplied")),
                }
            }
            _ => match self.resolve_argument(source, story, paragraphs, paragraph_index) {
                Ok(value) => value,
                Err(diagnostic) => return keep(&diagnostic),
            },
        };
        let Some(kind) = instruction.arguments.get(1) else {
            return keep(&format!("{} requires a barcode type", instruction.name));
        };
        let kind = match self.resolve_argument(kind, story, paragraphs, paragraph_index) {
            Ok(value) => value,
            Err(diagnostic) => return keep(&diagnostic),
        };
        let mut switches = Vec::with_capacity(instruction.switches.len());
        for field_switch in &instruction.switches {
            let argument = match &field_switch.argument {
                Some(argument) => {
                    match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                        Ok(value) => Some(value),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                }
                None => None,
            };
            switches.push((field_switch.name.clone(), argument));
        }
        match parse_barcode(&instruction.name, &switches, value, &kind) {
            Ok(barcode) => FieldOutcome::Barcode(barcode),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_toc(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let mut toc = TocField {
            heading_levels: None,
            custom_styles: Vec::new(),
            entries: TocEntrySelection::None,
            sequence_identifier: None,
            bookmark: None,
            hyperlink: false,
            use_outline_levels: false,
            omit_page_number_levels: None,
            page_number_separator: None,
            entry_page_separator: None,
        };
        let mut has_explicit_source = false;
        for field_switch in &instruction.switches {
            let argument = match &field_switch.argument {
                Some(argument) => {
                    match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                        Ok(value) => Some(value),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                }
                None => None,
            };
            match field_switch.name.as_str() {
                "h" if argument.is_none() => toc.hyperlink = true,
                "z" if argument.is_none() => {}
                "u" if argument.is_none() => {
                    toc.use_outline_levels = true;
                    has_explicit_source = true;
                }
                "o" => {
                    has_explicit_source = true;
                    toc.heading_levels = match argument {
                        Some(value) => match parse_level_range(&value, "TOC heading") {
                            Ok(levels) => Some(levels),
                            Err(diagnostic) => return keep(&diagnostic),
                        },
                        None => Some((1, 9)),
                    };
                }
                "n" => {
                    toc.omit_page_number_levels = match argument {
                        Some(value) => match parse_level_range(&value, "TOC omitted page-number") {
                            Ok(levels) => Some(levels),
                            Err(diagnostic) => return keep(&diagnostic),
                        },
                        None => Some((1, 9)),
                    };
                }
                "t" => {
                    let Some(value) = argument else {
                        return keep("field TOC switch \\t requires a text argument");
                    };
                    has_explicit_source = true;
                    toc.custom_styles = match parse_custom_styles(&value) {
                        Ok(styles) => styles,
                        Err(diagnostic) => return keep(&diagnostic),
                    };
                }
                "f" => {
                    has_explicit_source = true;
                    toc.entries = argument.map_or(TocEntrySelection::All, |value| {
                        TocEntrySelection::Identifier(value)
                    });
                }
                "b" | "p" | "s" | "d" => {
                    let Some(value) = argument else {
                        return keep(&format!(
                            "field TOC switch \\{} requires a text argument",
                            field_switch.name
                        ));
                    };
                    match field_switch.name.as_str() {
                        "b" => toc.bookmark = Some(value),
                        "p" if value.chars().count() == 1 => {
                            toc.page_number_separator = Some(value)
                        }
                        "p" => {
                            return keep(
                                "TOC page-number separator must contain exactly one character",
                            );
                        }
                        "s" if !value.is_empty() => toc.sequence_identifier = Some(value),
                        "s" => return keep("TOC sequence identifier must not be empty"),
                        "d" => toc.entry_page_separator = Some(value),
                        _ => unreachable!(),
                    }
                }
                name if argument.is_some() => {
                    return keep(&format!(
                        "field TOC switch \\{name} does not take an argument"
                    ));
                }
                name => return keep(&format!("field TOC uses unsupported switch \\{name}")),
            }
        }
        if !has_explicit_source {
            toc.heading_levels = Some((1, 9));
        }
        FieldOutcome::TableOfContents(toc)
    }

    fn evaluate_tc(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(entry) = instruction.arguments.first() else {
            return keep("TC requires entry text");
        };
        let entry = match self.resolve_argument(entry, story, paragraphs, paragraph_index) {
            Ok(entry) => entry,
            Err(diagnostic) => return keep(&diagnostic),
        };
        if entry.is_empty() {
            return keep("TC entry text must not be empty");
        }
        let mut tc = TcField {
            entry,
            level: 1,
            table_identifier: None,
            omit_page_number: false,
        };
        for field_switch in &instruction.switches {
            let argument = match &field_switch.argument {
                Some(argument) => {
                    match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                        Ok(value) => Some(value),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                }
                None => None,
            };
            match field_switch.name.as_str() {
                "n" if argument.is_none() => tc.omit_page_number = true,
                "f" => match argument {
                    Some(value) if !value.is_empty() => tc.table_identifier = Some(value),
                    Some(_) => return keep("TC table identifier must not be empty"),
                    None => return keep("field TC switch \\f requires a text argument"),
                },
                "l" => {
                    let Some(value) = argument else {
                        return keep("field TC switch \\l requires a text argument");
                    };
                    tc.level = match parse_toc_level(&value, "TC") {
                        Ok(level) => level,
                        Err(diagnostic) => return keep(&diagnostic),
                    };
                }
                name if argument.is_some() => {
                    return keep(&format!(
                        "field TC switch \\{name} does not take an argument"
                    ));
                }
                name => return keep(&format!("field TC uses unsupported switch \\{name}")),
            }
        }
        FieldOutcome::TableOfContentsEntry(tc)
    }
}

fn referenced_header_footer_parts(
    document: &Document,
    is_header: bool,
    include_control_sections: bool,
) -> Vec<(String, Vec<u8>)> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Vec::new();
    };
    let relationship_type = if is_header {
        rel_types::HEADER
    } else {
        rel_types::FOOTER
    };
    // Pagination inventories modeled block-control section endings. General
    // field evaluation retains its existing direct-body story discovery.
    let sections = if include_control_sections {
        rdocx_layout::engine::document_sections(&document.document)
    } else {
        document
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                _ => None,
            })
            .chain(document.document.body.sect_pr.as_ref())
            .collect()
    };
    let mut seen = HashSet::new();
    let mut parts = Vec::new();
    for section in sections {
        let references = section_header_footer_references(section, is_header);
        for reference in references {
            let Some(relationship) = relationships.get_by_id(&reference.rel_id) else {
                continue;
            };
            if relationship.rel_type != relationship_type
                || !crate::document::relationship_is_internal(relationship)
            {
                continue;
            }
            let part_name =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            if seen.insert(part_name.clone())
                && let Some(xml) = document.package.get_part(&part_name)
            {
                parts.push((part_name, xml.to_vec()));
            }
        }
    }
    parts
}

fn section_header_footer_references(
    section: &CT_SectPr,
    is_header: bool,
) -> &[rdocx_oxml::header_footer::HdrFtrRef] {
    if is_header {
        &section.header_refs
    } else {
        &section.footer_refs
    }
}

fn relationship_parts(document: &Document, relationship_type: &str) -> Vec<(String, Vec<u8>)> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Vec::new();
    };
    let mut seen = HashSet::new();
    relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == relationship_type)
        .filter(|relationship| crate::document::relationship_is_internal(relationship))
        .filter_map(|relationship| {
            let part_name =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            if !seen.insert(part_name.clone()) {
                return None;
            }
            Some((
                part_name.clone(),
                document.package.get_part(&part_name)?.to_vec(),
            ))
        })
        .collect()
}

/// All immutable field occurrences keyed by paragraph path and preorder index.
type PlacedPageFields =
    HashMap<(rdocx_layout::WordSourcePath, u32), Vec<rdocx_layout::WordFieldPlacement>>;

fn placed_page_fields(layout: &rdocx_layout::WordLayoutResult) -> PlacedPageFields {
    let mut placed = PlacedPageFields::new();
    for &placement in layout.field_placements() {
        if let Some(path) = layout.source_node(placement.source.node) {
            placed
                .entry((path.clone(), placement.source.index))
                .or_default()
                .push(placement);
        }
    }
    placed
}

fn collect_preorder_fields<'a>(
    field: &'a Field,
    locked: bool,
    fields: &mut Vec<(&'a Field, bool)>,
) {
    let locked = locked || field.locked() == Some(true);
    fields.push((field, locked));
    for nested in field.all_nested_fields_in_source_order() {
        collect_preorder_fields(nested, locked, fields);
    }
}

#[allow(clippy::too_many_arguments)]
fn push_page_field_updates(
    paragraphs: &[&CT_P],
    paths: &[Vec<rdocx_layout::WordSourcePath>],
    placed: &PlacedPageFields,
    layout: &rdocx_layout::WordLayoutResult,
    document: &CT_Document,
    updates: &mut Vec<Option<CachedFieldUpdate>>,
    report: &mut LayoutBackedFieldUpdateReport,
) {
    for (paragraph, paths) in paragraphs.iter().zip(paths) {
        let mut physical_fields = Vec::new();
        for run in paragraph.source_runs() {
            for content in &run.content {
                if let RunContent::Field(field) = content {
                    collect_preorder_fields(field, false, &mut physical_fields);
                }
            }
        }
        let physical_indices = physical_fields
            .iter()
            .enumerate()
            .map(|(index, (field, _))| (std::ptr::from_ref(*field), index as u32))
            .collect::<HashMap<_, _>>();
        let mut fields = Vec::new();
        for field in accepted_toc_runs(paragraph)
            .into_iter()
            .flat_map(|run| &run.run.content)
            .filter_map(|content| match content {
                RunContent::Field(field) => Some(field),
                _ => None,
            })
        {
            collect_preorder_fields(field, false, &mut fields);
        }
        for (field, locked) in fields {
            let Some(&index) = physical_indices.get(&std::ptr::from_ref(field)) else {
                report.diagnostics.push("field cache retained because its accepted owner has no physical source identity".into());
                updates.push(None);
                continue;
            };
            let instruction = field.effective_instruction();
            let name = instruction.name.as_str();
            if !matches!(
                name,
                "PAGE" | "NUMPAGES" | "SECTION" | "SECTIONPAGES" | "PAGEREF"
            ) {
                updates.push(None);
                continue;
            }
            let retained = |report: &mut LayoutBackedFieldUpdateReport, reason: &str| {
                report
                    .diagnostics
                    .push(format!("{name} field {index} retained cache: {reason}"))
            };
            if locked {
                retained(report, "locked owner");
                updates.push(None);
                continue;
            }
            let placement = paths
                .iter()
                .filter_map(|path| placed.get(&(path.clone(), index)))
                .flatten()
                .min_by_key(|placement| placement.physical_page);
            if placement.is_none() && name != "PAGEREF" {
                retained(report, "source did not reach pagination");
                updates.push(None);
                continue;
            }
            if matches!(name, "PAGE" | "NUMPAGES")
                && paths.iter().any(|path| {
                    matches!(
                        path.story,
                        rdocx_layout::WordStory::Header { .. }
                            | rdocx_layout::WordStory::Footer { .. }
                    )
                })
            {
                retained(
                    report,
                    "Word preserves dynamic header and footer caches on update and save",
                );
                updates.push(None);
                continue;
            }
            if let Some(switch) = unsupported_switch(&instruction) {
                retained(report, &format!("unsupported switch \\{switch}"));
                updates.push(None);
                continue;
            }
            if let Err(reason) = validate_instruction_shape(&instruction) {
                retained(report, &reason);
                updates.push(None);
                continue;
            }
            let value = match name {
                "PAGE" => placement.map(|placement| placement.displayed_page),
                "NUMPAGES" => Some(layout.layout.pages.len()),
                "SECTION" => placement.map(|placement| placement.section_index + 1),
                "SECTIONPAGES" => placement.map(|placement| {
                    layout
                        .page_sections()
                        .iter()
                        .filter(|record| record.section_index == placement.section_index)
                        .map(|record| record.physical_page)
                        .collect::<HashSet<_>>()
                        .len()
                }),
                "PAGEREF" => {
                    text_argument(&instruction, 0).and_then(|name| layout.bookmark_page(name))
                }
                _ => None,
            };
            let Some(value) = value else {
                retained(report, "bookmark target is missing, ambiguous or unplaced");
                updates.push(None);
                continue;
            };
            let cached_result = page_field_result(
                &instruction,
                value,
                if name == "PAGEREF" {
                    text_argument(&instruction, 0)
                        .and_then(|target| layout.bookmark_page_section(target))
                        .and_then(|target| section_page_format(document, target.section_index))
                } else if name == "PAGE" {
                    placement.and_then(|placement| {
                        section_page_format(document, placement.section_index)
                    })
                } else {
                    Some("decimal".into())
                },
            );
            match cached_result {
                Ok(cached_result) => {
                    match name {
                        "PAGE" => report.page_fields += 1,
                        "NUMPAGES" => report.num_pages_fields += 1,
                        "SECTION" => report.section_fields += 1,
                        "SECTIONPAGES" => report.section_pages_fields += 1,
                        "PAGEREF" => report.page_reference_fields += 1,
                        _ => {}
                    }
                    updates.push(Some(CachedFieldUpdate {
                        typed_runs: None,
                        comment_ranges: Vec::new(),
                        cached_result,
                        dirty: false,
                    }));
                }
                Err(reason) => {
                    retained(report, &reason);
                    updates.push(None);
                }
            }
        }
    }
}

fn page_field_result(
    instruction: &FieldInstruction,
    value: usize,
    section_format: Option<String>,
) -> std::result::Result<String, String> {
    let decimal = value.to_string();
    let section_format =
        section_format.ok_or_else(|| "unsupported section page number format".to_owned())?;
    let explicit_numeric = switch_text(instruction, "#").is_some()
        || instruction.switches.iter().any(|switch| {
            switch.name == "*"
                && switch
                    .argument
                    .as_ref()
                    .and_then(argument_text)
                    .is_some_and(|value| {
                        !matches!(
                            value.to_ascii_lowercase().as_str(),
                            "mergeformat" | "charformat"
                        )
                    })
        });
    if explicit_numeric {
        return apply_formats(instruction, &decimal, None);
    }
    let value = if section_format == "decimal" {
        decimal
    } else {
        let format = match section_format.as_str() {
            "upperRoman" => "ROMAN",
            "lowerRoman" => "roman",
            "upperLetter" => "ALPHABETIC",
            "lowerLetter" => "alphabetic",
            _ => return Err("unsupported section page number format".into()),
        };
        let mut base = instruction.clone();
        base.switches = vec![field_option_switch("*", Some(format.into()))];
        rdocx_layout::engine::format_numeric_field_general(&base, &decimal)?
    };
    apply_formats(instruction, &value, None)
}

fn section_page_format(document: &CT_Document, index: usize) -> Option<String> {
    let sections = rdocx_layout::engine::document_sections(document);
    let section = sections.get(index).copied();
    let Some(raw) = section
        .and_then(|section| section.page_number.as_ref())
        .and_then(|number| number.raw_xml.as_deref())
    else {
        return Some("decimal".into());
    };
    let mut reader = quick_xml::Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let mut format = "decimal".to_owned();
                for attribute in element.attributes() {
                    let attribute = attribute.ok()?;
                    match attribute.key.local_name().as_ref() {
                        b"fmt" => {
                            format = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    element.decoder(),
                                )
                                .ok()?
                                .into_owned()
                        }
                        b"chapStyle" => return None,
                        _ => {}
                    }
                }
                return Some(format);
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
        buffer.clear();
    }
}

#[derive(Clone, Copy)]
enum PackageStoryKind {
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum LegacyBlockOwner {
    Root,
    Table,
    Row,
    Cell,
}

#[derive(Clone, Copy)]
enum LegacyStoryElementKind {
    HeaderFooterRoot,
    FootnotesRoot,
    EndnotesRoot,
    CommentsRoot,
    NormalComment,
    NormalNote,
    Table,
    Row,
    Cell,
    ContentControl(LegacyBlockOwner),
    ContentControlContent(LegacyBlockOwner),
    Other,
}

struct LegacyStoryParagraph {
    start: usize,
    end: usize,
    general_context: bool,
    original: CT_P,
    paragraph: CT_P,
}

fn parsed_physical_field_paragraph(xml: &[u8]) -> Result<CT_P> {
    let mut body = CT_Body::default();
    body.content
        .push(BodyContent::Paragraph(Box::new(CT_P::from_xml_fragment(
            xml,
        )?)));
    prepare_physical_story_projection(&mut body, &mut [])?;
    let BodyContent::Paragraph(paragraph) = body.content.remove(0) else {
        unreachable!()
    };
    Ok(*paragraph)
}

fn legacy_story_paragraphs(
    xml: &[u8],
    story_kind: PackageStoryKind,
) -> Result<Vec<LegacyStoryParagraph>> {
    validate_story_document_declarations_and_doctype(xml)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<LegacyStoryElementKind>::new();
    let mut bindings = BTreeMap::<String, Vec<u8>>::new();
    let mut binding_scopes = Vec::new();
    let mut paragraphs = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let general_context = |depth: usize| match story_kind {
        PackageStoryKind::Header | PackageStoryKind::Footer => depth == 1,
        PackageStoryKind::Footnotes | PackageStoryKind::Endnotes => depth == 2,
        PackageStoryKind::Comments => false,
    };
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package story XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let word = namespace_is_word(&namespace);
                let name = element.name();
                let local = local_name(name.as_ref());
                if stack.is_empty() {
                    if root_seen || !package_story_root_matches(story_kind, word, local) {
                        return Err(Error::Other(
                            "package story must contain exactly one relationship-appropriate root"
                                .to_owned(),
                        ));
                    }
                    root_seen = true;
                }
                let kind = legacy_story_element_kind(
                    story_kind,
                    &stack,
                    word,
                    local,
                    &element,
                    reader.resolver(),
                );
                if matches!(kind, Some(LegacyStoryElementKind::Other)) && word && local == b"p" {
                    reader
                        .read_to_end_into(element.name(), &mut Vec::new())
                        .map_err(|error| {
                            Error::Other(format!("invalid package story paragraph: {error}"))
                        })?;
                } else if kind.is_none() && word && local == b"p" {
                    let start_after = reader.buffer_position() as usize;
                    reader
                        .read_to_end_into(element.name(), &mut Vec::new())
                        .map_err(|error| {
                            Error::Other(format!("invalid package story paragraph: {error}"))
                        })?;
                    let end = reader.buffer_position() as usize;
                    let fragment = paragraph_fragment_with_bindings(
                        &xml[before..end],
                        start_after - before,
                        &bindings,
                    )?;
                    let paragraph = parsed_physical_field_paragraph(&fragment)?;
                    paragraphs.push(LegacyStoryParagraph {
                        start: before,
                        end,
                        general_context: general_context(stack.len()),
                        original: paragraph.clone(),
                        paragraph,
                    });
                } else {
                    let mut local_bindings = bindings.clone();
                    apply_namespace_declarations(&element, &mut local_bindings)?;
                    binding_scopes.push(std::mem::replace(&mut bindings, local_bindings));
                    stack.push(kind.unwrap_or(LegacyStoryElementKind::Other));
                }
            }
            Event::End(_) => {
                if stack.pop().is_none() {
                    return Err(Error::Other(
                        "package story XML has an unmatched end element".to_owned(),
                    ));
                }
                bindings = binding_scopes.pop().unwrap_or_default();
                if stack.is_empty() {
                    root_closed = true;
                }
            }
            // A self-closing paragraph is a story paragraph too, as the
            // layout source inventory counts it.
            Event::Empty(element) if !stack.is_empty() => {
                let word = namespace_is_word(&namespace);
                let name = element.name();
                let local = local_name(name.as_ref());
                if word
                    && local == b"p"
                    && legacy_story_element_kind(
                        story_kind,
                        &stack,
                        word,
                        local,
                        &element,
                        reader.resolver(),
                    )
                    .is_none()
                {
                    let end = reader.buffer_position() as usize;
                    let fragment = paragraph_fragment_with_bindings(
                        &xml[before..end],
                        end - before,
                        &bindings,
                    )?;
                    let paragraph = parsed_physical_field_paragraph(&fragment)?;
                    paragraphs.push(LegacyStoryParagraph {
                        start: before,
                        end,
                        general_context: general_context(stack.len()),
                        original: paragraph.clone(),
                        paragraph,
                    });
                }
            }
            Event::Empty(element) if stack.is_empty() => {
                let word = namespace_is_word(&namespace);
                let name = element.name();
                let local = local_name(name.as_ref());
                if root_seen || !package_story_root_matches(story_kind, word, local) {
                    return Err(Error::Other(
                        "package story must contain exactly one relationship-appropriate root"
                            .to_owned(),
                    ));
                }
                root_seen = true;
                root_closed = true;
            }
            Event::Text(text)
                if stack.is_empty() && text.iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(Error::Other(
                    "package story has non-whitespace text outside its root".to_owned(),
                ));
            }
            Event::CData(_) if stack.is_empty() => {
                return Err(Error::Other(
                    "package story has character data outside its root".to_owned(),
                ));
            }
            Event::GeneralRef(_) if stack.is_empty() => {
                return Err(Error::Other(
                    "package story has a character reference outside its root".to_owned(),
                ));
            }
            Event::Eof => {
                if !stack.is_empty() {
                    return Err(Error::Other(
                        "package story XML has an unclosed element".to_owned(),
                    ));
                }
                if !root_seen || !root_closed {
                    return Err(Error::Other(
                        "package story must contain exactly one relationship-appropriate root"
                            .to_owned(),
                    ));
                }
                return Ok(paragraphs);
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn validate_story_document_declarations_and_doctype(xml: &[u8]) -> Result<()> {
    validate_strict_xml_1_0(xml).map_err(package_story_lexical_error)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut declaration_allowed = true;
    let mut declaration_seen = false;
    let mut depth = 0usize;
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package story XML: {error}")))?;
        let event = event.into_owned();
        match event {
            Event::Decl(_) => {
                if !declaration_allowed || declaration_seen {
                    return Err(Error::Other(
                        "misplaced or duplicate package story XML declaration".to_owned(),
                    ));
                }
                declaration_seen = true;
                declaration_allowed = false;
            }
            Event::DocType(_) => {
                return Err(Error::Other(
                    "package story XML cannot contain a document type".to_owned(),
                ));
            }
            Event::Start(_) => {
                declaration_allowed = false;
                depth += 1;
            }
            Event::Empty(_) => {
                declaration_allowed = false;
            }
            Event::End(_) => {
                declaration_allowed = false;
                depth = depth.saturating_sub(1);
            }
            Event::GeneralRef(_) => {
                declaration_allowed = false;
            }
            Event::Eof => return Ok(()),
            _ => declaration_allowed = false,
        }
        buffer.clear();
    }
}

fn package_story_lexical_error(error: XmlLexicalError) -> Error {
    let message = match error {
        XmlLexicalError::InvalidUtf8 => "invalid package story XML: input is not UTF-8".to_owned(),
        XmlLexicalError::InvalidDeclaration(reason) if reason == "must begin with version" => {
            "package story XML declaration must begin with version".to_owned()
        }
        XmlLexicalError::InvalidDeclaration(reason)
            if reason == "attributes are invalid, duplicated, or out of order" =>
        {
            "invalid package story XML declaration".to_owned()
        }
        XmlLexicalError::InvalidDeclaration(reason) => {
            format!("invalid package story XML declaration: {reason}")
        }
        XmlLexicalError::ForbiddenLiteralCharacter => {
            "package story XML contains a forbidden literal XML 1.0 character".to_owned()
        }
        XmlLexicalError::InvalidName(reason) if reason.starts_with("qualified name ") => {
            format!("invalid package story XML {reason}")
        }
        XmlLexicalError::InvalidName(reason) if reason.starts_with("name ") => {
            format!("invalid package story XML {reason}")
        }
        XmlLexicalError::InvalidName(reason) => format!("invalid package story XML: {reason}"),
        XmlLexicalError::InvalidNamespace(reason)
            if reason == "element uses the reserved xmlns prefix" =>
        {
            "package story XML element uses the reserved xmlns prefix".to_owned()
        }
        XmlLexicalError::InvalidNamespace(reason) if reason == "invalid namespace declaration" => {
            "package story XML contains an invalid namespace declaration".to_owned()
        }
        XmlLexicalError::InvalidNamespace(_) => {
            "package story XML uses an unbound namespace prefix".to_owned()
        }
        XmlLexicalError::DuplicateExpandedAttribute => {
            "package story XML element has duplicate expanded-name attributes".to_owned()
        }
        XmlLexicalError::InvalidReference(reason)
            if reason == "attribute contains a literal less-than sign" =>
        {
            "package story XML attribute contains a literal less-than sign".to_owned()
        }
        XmlLexicalError::InvalidReference(reason)
            if reason == "character reference is not legal in XML 1.0" =>
        {
            "package story character reference is not legal in XML 1.0".to_owned()
        }
        XmlLexicalError::InvalidReference(reason)
            if reason.starts_with("undeclared entity reference ") =>
        {
            reason.replacen(
                "undeclared entity reference ",
                "undeclared package story XML entity reference ",
                1,
            )
        }
        XmlLexicalError::InvalidReference(reason) => {
            format!("invalid package story XML: {reason}")
        }
        XmlLexicalError::InvalidProcessingInstruction(reason)
            if reason == "reserved XML target" =>
        {
            "reserved package story XML processing instruction".to_owned()
        }
        XmlLexicalError::InvalidProcessingInstruction(reason) if reason.starts_with("name ") => {
            format!("invalid package story XML {reason}")
        }
        XmlLexicalError::InvalidProcessingInstruction(reason) => {
            format!("invalid package story XML processing instruction: {reason}")
        }
        XmlLexicalError::InvalidComment(reason) => format!("invalid package story XML: {reason}"),
    };
    Error::Other(message)
}

fn legacy_story_element_kind(
    story_kind: PackageStoryKind,
    stack: &[LegacyStoryElementKind],
    word: bool,
    local: &[u8],
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Option<LegacyStoryElementKind> {
    if stack.is_empty() {
        return Some(match story_kind {
            PackageStoryKind::Header if word && local == b"hdr" => {
                LegacyStoryElementKind::HeaderFooterRoot
            }
            PackageStoryKind::Footer if word && local == b"ftr" => {
                LegacyStoryElementKind::HeaderFooterRoot
            }
            PackageStoryKind::Footnotes if word && local == b"footnotes" => {
                LegacyStoryElementKind::FootnotesRoot
            }
            PackageStoryKind::Endnotes if word && local == b"endnotes" => {
                LegacyStoryElementKind::EndnotesRoot
            }
            PackageStoryKind::Comments if word && local == b"comments" => {
                LegacyStoryElementKind::CommentsRoot
            }
            _ => LegacyStoryElementKind::Other,
        });
    }
    if word
        && matches!(
            (story_kind, stack.last()),
            (
                PackageStoryKind::Footnotes,
                Some(LegacyStoryElementKind::FootnotesRoot)
            ) | (
                PackageStoryKind::Endnotes,
                Some(LegacyStoryElementKind::EndnotesRoot)
            )
        )
        && matches!(local, b"footnote" | b"endnote")
        && note_is_normal(element, resolver)
    {
        return Some(LegacyStoryElementKind::NormalNote);
    }
    if matches!(story_kind, PackageStoryKind::Comments)
        && word
        && local == b"comment"
        && matches!(stack.last(), Some(LegacyStoryElementKind::CommentsRoot))
    {
        return Some(LegacyStoryElementKind::NormalComment);
    }
    if let Some(LegacyStoryElementKind::ContentControl(owner)) = stack.last()
        && word
        && local == b"sdtContent"
    {
        return Some(LegacyStoryElementKind::ContentControlContent(*owner));
    }
    let owner = match stack.last()? {
        LegacyStoryElementKind::HeaderFooterRoot
        | LegacyStoryElementKind::NormalNote
        | LegacyStoryElementKind::NormalComment => LegacyBlockOwner::Root,
        LegacyStoryElementKind::Table => LegacyBlockOwner::Table,
        LegacyStoryElementKind::Row => LegacyBlockOwner::Row,
        LegacyStoryElementKind::Cell => LegacyBlockOwner::Cell,
        LegacyStoryElementKind::ContentControlContent(owner) => *owner,
        _ => return Some(LegacyStoryElementKind::Other),
    };
    if !word {
        return Some(LegacyStoryElementKind::Other);
    }
    match local {
        b"p" if matches!(owner, LegacyBlockOwner::Root | LegacyBlockOwner::Cell) => None,
        b"tbl" if matches!(owner, LegacyBlockOwner::Root | LegacyBlockOwner::Cell) => {
            Some(LegacyStoryElementKind::Table)
        }
        b"tr" if owner == LegacyBlockOwner::Table => Some(LegacyStoryElementKind::Row),
        b"tc" if owner == LegacyBlockOwner::Row => Some(LegacyStoryElementKind::Cell),
        b"sdt" => Some(LegacyStoryElementKind::ContentControl(owner)),
        _ => Some(LegacyStoryElementKind::Other),
    }
}

fn package_story_root_matches(story_kind: PackageStoryKind, word: bool, local: &[u8]) -> bool {
    word && match story_kind {
        PackageStoryKind::Header => local == b"hdr",
        PackageStoryKind::Footer => local == b"ftr",
        PackageStoryKind::Footnotes => local == b"footnotes",
        PackageStoryKind::Endnotes => local == b"endnotes",
        PackageStoryKind::Comments => local == b"comments",
    }
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    bindings: &mut BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Other(format!("invalid namespace declaration: {error}")))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            bindings.insert(
                String::from_utf8_lossy(key).into_owned(),
                attribute.value.as_ref().to_vec(),
            );
        }
    }
    Ok(())
}

fn paragraph_fragment_with_bindings(
    raw: &[u8],
    start_len: usize,
    bindings: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<u8>> {
    // Declarations go before the `>` of a start tag, or the `/>` of an
    // empty paragraph.
    let close = if raw[..start_len].ends_with(b"/>") {
        2
    } else {
        1
    };
    let insert_at = start_len
        .checked_sub(close)
        .ok_or_else(|| Error::Other("package story paragraph start tag is missing".to_owned()))?;
    let mut output = raw[..insert_at].to_vec();
    for (key, value) in bindings {
        if start_tag_has_raw_attribute(&raw[..start_len], key.as_bytes()) {
            continue;
        }
        output.push(b' ');
        output.extend_from_slice(key.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(value);
        output.push(b'"');
    }
    output.extend_from_slice(&raw[insert_at..]);
    Ok(output)
}

fn start_tag_has_raw_attribute(start: &[u8], name: &[u8]) -> bool {
    let mut reader = quick_xml::Reader::from_reader(start);
    let mut buffer = Vec::new();
    matches!(
        reader.read_event_into(&mut buffer),
        Ok(Event::Start(element) | Event::Empty(element))
            if element
                .attributes()
                .flatten()
                .any(|attribute| attribute.key.as_ref() == name)
    )
}

fn patch_legacy_story_field_sources(
    xml: &[u8],
    paragraphs: &[LegacyStoryParagraph],
    preserve_original_cache_sources: bool,
) -> Result<Vec<u8>> {
    let mut edits = Vec::new();
    for paragraph in paragraphs {
        let mut search_start = 0usize;
        let replacements = if preserve_original_cache_sources {
            paragraph_field_source_replacements(&paragraph.original, &paragraph.paragraph)?
        } else {
            staged_paragraph_field_source_replacements(&paragraph.paragraph)?
        };
        for (source, replacement) in replacements {
            let Some(start) = find_typed_field_source(
                xml,
                paragraph.start,
                paragraph.end,
                &source,
                search_start,
            )?
            else {
                return Err(Error::Other(
                    "package story field source was not found at its typed paragraph boundary"
                        .to_owned(),
                ));
            };
            let end = start + source.len();
            edits.push(FieldSourceEdit {
                start: paragraph.start + start,
                end: paragraph.start + end,
                replacement,
            });
            search_start = end;
        }
    }
    let mut updated = xml.to_vec();
    for edit in edits.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

struct FieldSourceEdit {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

fn paragraph_field_source_replacements(
    original: &CT_P,
    staged: &CT_P,
) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    fn instruction_context(paragraph: &CT_P) -> Vec<(String, bool)> {
        fn field_context(field: &Field, locked: bool, output: &mut Vec<(String, bool)>) {
            let locked = locked || field.locked() == Some(true);
            output.push((field.effective_instruction_text(), locked));
            for child in field.nested_fields_in_source_order() {
                field_context(child, locked, output);
            }
        }
        let mut output = Vec::new();
        for run in accepted_toc_runs(paragraph) {
            for content in &run.run.content {
                if let RunContent::Field(field) = content {
                    field_context(field, false, &mut output);
                }
            }
        }
        output
    }
    if instruction_context(original) != instruction_context(staged) {
        return Err(Error::Other(
            "field cache publication changed physical instruction or lock ownership".into(),
        ));
    }
    let original = staged_paragraph_field_source_replacements(original)?;
    let staged = staged_paragraph_field_source_replacements(staged)?;
    if original.len() != staged.len() {
        return Err(Error::Other(
            "field cache publication changed physical source span cardinality".into(),
        ));
    }
    Ok(original
        .into_iter()
        .zip(staged)
        .map(|((source, _), (_, replacement))| (source, replacement))
        .collect())
}

fn staged_paragraph_field_source_replacements(paragraph: &CT_P) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut replacements = Vec::new();
    let mut next_run = 0;
    for boundary in 0..=paragraph.runs.len() {
        for (_, owner) in accepted_paragraph_owners(paragraph, boundary) {
            match owner {
                AcceptedParagraphOwner::Control(control) => {
                    append_control_field_source_replacements(control, &mut replacements)?;
                }
                AcceptedParagraphOwner::Revision(revision) => {
                    if matches!(
                        revision.kind(),
                        RevisionKind::Insertion | RevisionKind::MoveTo
                    ) && let Some(content) = revision.content_paragraph()
                    {
                        replacements.extend(staged_paragraph_field_source_replacements(content)?);
                    }
                }
            }
        }
        if boundary < next_run || boundary == paragraph.runs.len() {
            continue;
        }
        next_run =
            boundary + paragraph.field_source_replacements_at(boundary, &mut replacements)?;
    }
    Ok(replacements)
}

fn append_control_field_source_replacements(
    control: &CT_Sdt,
    output: &mut Vec<(Vec<u8>, Vec<u8>)>,
) -> Result<()> {
    output.extend(
        control
            .inline_field_source_replacements()
            .into_iter()
            .map(|(source, replacement)| (source.to_vec(), replacement.to_vec())),
    );
    for boundary in 0..=control.content.len() {
        for (_, revision) in control.revisions().iter().filter(|(at, _)| *at == boundary) {
            if matches!(
                revision.kind(),
                RevisionKind::Insertion | RevisionKind::MoveTo
            ) && let Some(paragraph) = revision.content_paragraph()
            {
                output.extend(staged_paragraph_field_source_replacements(paragraph)?);
            }
        }
        let Some(content) = control.content.get(boundary) else {
            continue;
        };
        match content {
            SdtContent::ContentControl(control) => {
                append_control_field_source_replacements(control, output)?
            }
            SdtContent::Paragraph(paragraph) => {
                output.extend(staged_paragraph_field_source_replacements(paragraph)?);
            }
            SdtContent::Table(_)
            | SdtContent::Row(_)
            | SdtContent::Cell(_)
            | SdtContent::Run(_)
            | SdtContent::RawXml(_) => {}
        }
    }
    Ok(())
}

fn note_is_normal(element: &BytesStart<'_>, resolver: &NamespaceResolver) -> bool {
    let mut id = None;
    let mut note_type = None;
    for attribute in element.attributes() {
        let Ok(attribute) = attribute else {
            return false;
        };
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !namespace_is_word(&namespace) {
            continue;
        }
        if local.as_ref() == b"id" {
            let Ok(raw) = std::str::from_utf8(&attribute.value) else {
                return false;
            };
            let Ok(value) = quick_xml::escape::unescape(raw) else {
                return false;
            };
            let Ok(value) = value.parse::<i32>() else {
                return false;
            };
            if id.replace(value).is_some() {
                return false;
            }
        } else if local.as_ref() == b"type" {
            let Ok(raw) = std::str::from_utf8(&attribute.value) else {
                return false;
            };
            let Ok(value) = quick_xml::escape::unescape(raw) else {
                return false;
            };
            if note_type.replace(value.into_owned()).is_some() {
                return false;
            }
        }
    }
    match note_type.as_deref() {
        Some("separator" | "continuationSeparator" | "continuationNotice") => false,
        Some("normal") => id.is_some(),
        Some(_) => false,
        None => id.is_some_and(|id| id > 0),
    }
}

pub(crate) fn find_typed_field_source(
    xml: &[u8],
    paragraph_start: usize,
    paragraph_end: usize,
    source: &[u8],
    search_start: usize,
) -> Result<Option<usize>> {
    let paragraph_xml = &xml[paragraph_start..paragraph_end];
    let mut candidate_start = search_start;
    while let Some(relative) = find_bytes(&paragraph_xml[candidate_start..], source) {
        let candidate = candidate_start + relative;
        if field_source_has_typed_ancestors(xml, paragraph_start + candidate, paragraph_start)? {
            return Ok(Some(candidate));
        }
        candidate_start = candidate + source.len().max(1);
    }
    Ok(None)
}

fn field_source_has_typed_ancestors(
    xml: &[u8],
    source_start: usize,
    paragraph_start: usize,
) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut ancestors = Vec::<(usize, bool, Vec<u8>)>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package paragraph XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        match event {
            Event::Start(element) => {
                if before == source_start {
                    let name = element.name();
                    let local = local_name(name.as_ref());
                    return Ok(word
                        && matches!(local, b"fldSimple" | b"r")
                        && typed_field_ancestors(&ancestors, paragraph_start));
                }
                ancestors.push((before, word, local_name(element.name().as_ref()).to_vec()));
            }
            Event::Empty(element) => {
                if before == source_start {
                    let name = element.name();
                    let local = local_name(name.as_ref());
                    return Ok(word
                        && matches!(local, b"fldSimple" | b"r")
                        && typed_field_ancestors(&ancestors, paragraph_start));
                }
            }
            Event::End(_) => {
                ancestors.pop();
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn typed_field_ancestors(ancestors: &[(usize, bool, Vec<u8>)], paragraph_start: usize) -> bool {
    let Some(paragraph_index) = ancestors
        .iter()
        .position(|(start, _, _)| *start == paragraph_start)
    else {
        return false;
    };
    let (_, paragraph_word, paragraph) = &ancestors[paragraph_index];
    *paragraph_word
        && paragraph.as_slice() == b"p"
        && ancestors
            .iter()
            .skip(paragraph_index + 1)
            .all(|(_, word, local)| {
                *word
                    && matches!(
                        local.as_slice(),
                        b"hyperlink"
                            | b"sdt"
                            | b"sdtContent"
                            | b"ins"
                            | b"del"
                            | b"moveFrom"
                            | b"moveTo"
                    )
            })
}

fn namespace_is_word(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == W_NS.as_bytes())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        None
    } else {
        haystack
            .windows(needle.len())
            .position(|window| window == needle)
    }
}

fn apply_updates_to_body(
    body: &mut CT_Body,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for content in &mut body.content {
        match content {
            BodyContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index)?;
            }
            BodyContent::Table(table) => apply_updates_to_table(table, updates, update_index)?,
            BodyContent::ContentControl(control) => {
                apply_updates_to_block_control(control, updates, update_index)?;
            }
            BodyContent::RawXml(_) => {}
        }
    }
    Ok(())
}

fn apply_updates_to_table(
    table: &mut CT_Tbl,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter_mut()
            .filter(|(position, _, _)| *position == boundary)
        {
            apply_updates_to_block_control(control, updates, update_index)?;
        }
        if let Some(row) = table.rows.get_mut(boundary) {
            apply_updates_to_row(row, updates, update_index)?;
        }
    }
    Ok(())
}

fn apply_updates_to_row(
    row: &mut CT_Row,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter_mut()
            .filter(|(position, _, _)| *position == boundary)
        {
            apply_updates_to_block_control(control, updates, update_index)?;
        }
        if let Some(cell) = row.cells.get_mut(boundary) {
            apply_updates_to_cell(cell, updates, update_index)?;
        }
    }
    Ok(())
}

fn apply_updates_to_cell(
    cell: &mut CT_Tc,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index)?;
            }
            CellContent::Table(table) => apply_updates_to_table(table, updates, update_index)?,
            CellContent::ContentControl(control) => {
                apply_updates_to_block_control(control, updates, update_index)?;
            }
        }
    }
    Ok(())
}

fn apply_updates_to_block_control(
    control: &mut CT_Sdt,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index)?;
            }
            SdtContent::Table(table) => apply_updates_to_table(table, updates, update_index)?,
            SdtContent::Row(row) => apply_updates_to_row(row, updates, update_index)?,
            SdtContent::Cell(cell) => apply_updates_to_cell(cell, updates, update_index)?,
            SdtContent::ContentControl(control) => {
                apply_updates_to_block_control(control, updates, update_index)?;
            }
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
    Ok(())
}

fn apply_updates_to_paragraph(
    paragraph: &mut CT_P,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for path in paragraph.accepted_run_paths() {
        let Some(source) = paragraph.accepted_run(&path) else {
            return Err(Error::Other(
                "accepted field update run path is stale".into(),
            ));
        };
        if !source
            .content
            .iter()
            .any(|content| matches!(content, RunContent::Field(_)))
        {
            continue;
        }
        let mut replacement = source.clone();
        apply_updates_to_run(&mut replacement, updates, update_index)?;
        if !paragraph.replace_accepted_run(&path, replacement)? {
            return Err(Error::Other(
                "accepted field update owner disappeared".into(),
            ));
        }
    }
    Ok(())
}

fn apply_updates_to_run(
    run: &mut rdocx_oxml::text::CT_R,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    for content in &mut run.content {
        if let RunContent::Field(field) = content {
            apply_updates_to_field(field, updates, update_index)?;
        }
    }
    Ok(())
}

fn apply_updates_to_field(
    field: &mut Field,
    updates: &[Option<CachedFieldUpdate>],
    update_index: &mut usize,
) -> Result<()> {
    let Some(update) = updates.get(*update_index) else {
        return Ok(());
    };
    if let Some(update) = update {
        field.cached_result.clone_from(&update.cached_result);
        field.dirty = Some(update.dirty);
    }
    *update_index += 1;

    let nested_pointers = field
        .nested_fields_in_source_order()
        .into_iter()
        .map(|nested| std::ptr::from_ref(nested) as usize)
        .collect::<Vec<_>>();
    for pointer in nested_pointers {
        if let Some(nested) = nested_field_mut(field, pointer) {
            apply_updates_to_field(nested, updates, update_index)?;
        }
    }
    for index in 0..field.cached_fields_in_source_order().len() {
        if let Some(nested) = field.cached_field_mut(index) {
            apply_updates_to_field(nested, updates, update_index)?;
        }
    }
    if let Some(runs) = update
        .as_ref()
        .and_then(|update| update.typed_runs.as_ref())
    {
        field.set_cached_runs_with_comment_ranges(
            runs.clone(),
            update
                .as_ref()
                .expect("typed cache update")
                .comment_ranges
                .clone(),
        )?;
    }
    if update.is_none() && !field.cached_fields_in_source_order().is_empty() {
        field.refresh_cached_field_projection();
    }
    Ok(())
}

fn nested_field_mut(field: &mut Field, pointer: usize) -> Option<&mut Field> {
    for argument in &mut field.instruction.arguments {
        if let FieldArgument::Nested(nested) = argument
            && std::ptr::from_ref(nested.as_ref()) as usize == pointer
        {
            return Some(nested.as_mut());
        }
    }
    for field_switch in &mut field.instruction.switches {
        if let Some(FieldArgument::Nested(nested)) = &mut field_switch.argument
            && std::ptr::from_ref(nested.as_ref()) as usize == pointer
        {
            return Some(nested.as_mut());
        }
    }
    None
}

fn collect_body_paragraphs<'a>(body: &'a CT_Body, output: &mut Vec<&'a CT_P>) {
    for content in &body.content {
        match content {
            BodyContent::Paragraph(paragraph) => output.push(paragraph),
            BodyContent::Table(table) => collect_table_paragraphs(table, output),
            BodyContent::ContentControl(control) => {
                collect_control_paragraphs(control, BlockControlOwner::Body, output)
            }
            BodyContent::RawXml(_) => {}
        }
    }
}

fn collect_table_paragraphs<'a>(table: &'a CT_Tbl, output: &mut Vec<&'a CT_P>) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(position, _, _)| *position == boundary)
        {
            collect_control_paragraphs(control, BlockControlOwner::Table, output);
        }
        if let Some(row) = table.rows.get(boundary) {
            collect_row_paragraphs(row, output);
        }
    }
}

fn collect_row_paragraphs<'a>(row: &'a CT_Row, output: &mut Vec<&'a CT_P>) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(position, _, _)| *position == boundary)
        {
            collect_control_paragraphs(control, BlockControlOwner::Row, output);
        }
        if let Some(cell) = row.cells.get(boundary) {
            collect_cell_paragraphs(cell, output);
        }
    }
}

fn collect_cell_paragraphs<'a>(cell: &'a CT_Tc, output: &mut Vec<&'a CT_P>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => output.push(paragraph),
            CellContent::Table(table) => collect_table_paragraphs(table, output),
            CellContent::ContentControl(control) => {
                collect_control_paragraphs(control, BlockControlOwner::Cell, output)
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BlockControlOwner {
    Body,
    Table,
    Row,
    Cell,
}

fn collect_control_paragraphs<'a>(
    control: &'a CT_Sdt,
    owner: BlockControlOwner,
    output: &mut Vec<&'a CT_P>,
) {
    for content in &control.content {
        match (owner, content) {
            (
                BlockControlOwner::Body | BlockControlOwner::Cell,
                SdtContent::Paragraph(paragraph),
            ) => output.push(paragraph),
            (BlockControlOwner::Body | BlockControlOwner::Cell, SdtContent::Table(table)) => {
                collect_table_paragraphs(table, output)
            }
            (BlockControlOwner::Table, SdtContent::Row(row)) => collect_row_paragraphs(row, output),
            (BlockControlOwner::Row, SdtContent::Cell(cell)) => {
                collect_cell_paragraphs(cell, output)
            }
            (_, SdtContent::ContentControl(control)) => {
                collect_control_paragraphs(control, owner, output)
            }
            _ => {}
        }
    }
}

fn keep(diagnostic: &str) -> FieldOutcome {
    FieldOutcome::KeepStored {
        diagnostic: format!("{diagnostic}, stored display retained"),
    }
}

const MAX_FORMULA_BYTES: usize = 4096;
const MAX_FORMULA_TOKENS: usize = 512;
const MAX_FORMULA_DEPTH: usize = 32;

fn parse_toc_level(value: &str, field: &str) -> std::result::Result<u8, String> {
    value
        .trim()
        .parse::<u8>()
        .ok()
        .filter(|level| (1..=9).contains(level))
        .ok_or_else(|| format!("{field} level must be from 1 through 9"))
}

fn parse_level_range(value: &str, field: &str) -> std::result::Result<(u8, u8), String> {
    let Some((start, end)) = value.split_once('-') else {
        let level = parse_toc_level(value, field)?;
        return Ok((level, level));
    };
    let start = parse_toc_level(start, field)?;
    let end = parse_toc_level(end, field)?;
    if start > end {
        return Err(format!("{field} range starts after it ends"));
    }
    Ok((start, end))
}

fn parse_custom_styles(value: &str) -> std::result::Result<Vec<(String, u8)>, String> {
    let mut parts = value.split(',').collect::<Vec<_>>();
    if parts.last().is_some_and(|part| part.trim().is_empty()) {
        parts.pop();
    }
    if parts.len() % 2 != 0 || parts.is_empty() {
        return Err("TOC custom styles require style and level pairs".to_owned());
    }
    let mut styles = Vec::with_capacity(parts.len() / 2);
    for pair in parts.chunks_exact(2) {
        let name = pair[0].trim();
        if name.is_empty() {
            return Err("TOC custom style name must not be empty".to_owned());
        }
        styles.push((name.to_owned(), parse_toc_level(pair[1], "TOC style")?));
    }
    Ok(styles)
}

fn parse_barcode(
    field: &str,
    switches: &[(String, Option<String>)],
    value: String,
    kind: &str,
) -> std::result::Result<BarcodeField, String> {
    let kind = parse_barcode_kind(kind).ok_or_else(|| {
        format!(
            "{field} barcode type {} is unsupported",
            kind.to_ascii_uppercase()
        )
    })?;
    if value.is_empty() || value.chars().count() > 1024 {
        return Err(format!(
            "{field} value must contain from 1 through 1024 characters"
        ));
    }
    validate_barcode_value(field, kind, &value)?;
    let mut barcode = BarcodeField {
        value,
        kind,
        height: None,
        scale: None,
        error_correction: None,
        point_of_sale_style: None,
        case_style: None,
        fix_check_digit: false,
        rotation: None,
        foreground_color: None,
        background_color: None,
        display_text: false,
        add_start_stop: false,
    };
    for (name, argument) in switches {
        match name.as_str() {
            "t" if argument.is_none() => barcode.display_text = true,
            "x" if argument.is_none() => barcode.fix_check_digit = true,
            "d" => {
                if argument.is_some() {
                    return Err(format!(
                        "field {field} switch \\d does not take an argument"
                    ));
                }
                if !matches!(kind, BarcodeKind::Nw7 | BarcodeKind::Code39) {
                    return Err(format!("{field} switch \\d requires NW7 or CODE39"));
                }
                barcode.add_start_stop = true;
            }
            "h" | "s" | "q" | "p" | "c" | "r" | "f" | "b" => {
                let Some(value) = argument.as_deref() else {
                    return Err(format!(
                        "field {field} switch \\{name} requires a text argument"
                    ));
                };
                match name.as_str() {
                    "h" => barcode.height = Some(parse_unsigned_integer(value, field, "height")?),
                    "s" => {
                        barcode.scale = Some(
                            u16::try_from(parse_bounded_integer(value, 10, 1000, field, "scale")?)
                                .expect("bounded barcode scale fits u16"),
                        )
                    }
                    "q" => {
                        if kind != BarcodeKind::Qr {
                            return Err(format!("{field} switch \\q requires QR"));
                        }
                        barcode.error_correction = Some(
                            u8::try_from(parse_bounded_integer(
                                value,
                                0,
                                3,
                                field,
                                "error correction",
                            )?)
                            .expect("bounded error correction fits u8"),
                        );
                    }
                    "p" => {
                        if !matches!(
                            kind,
                            BarcodeKind::Upca
                                | BarcodeKind::Upce
                                | BarcodeKind::Ean13
                                | BarcodeKind::Ean8
                        ) {
                            return Err(format!(
                                "{field} switch \\p requires UPCA, UPCE, EAN13, or EAN8"
                            ));
                        }
                        barcode.point_of_sale_style =
                            Some(parse_point_of_sale_style(value).ok_or_else(|| {
                                format!(
                                    "{field} point-of-sale style must be STD, SUP2, SUP5, or CASE"
                                )
                            })?);
                    }
                    "c" => {
                        if !matches!(kind, BarcodeKind::Case | BarcodeKind::Itf14) {
                            return Err(format!("{field} switch \\c requires ITF14"));
                        }
                        barcode.case_style = Some(parse_case_style(value).ok_or_else(|| {
                            format!("{field} case style must be STD, EXT, or ADD")
                        })?);
                    }
                    "r" => {
                        let rotation = parse_bounded_integer(value, 0, 3, field, "rotation")?;
                        barcode.rotation =
                            Some(u8::try_from(rotation).expect("bounded barcode rotation fits u8"));
                    }
                    "f" => barcode.foreground_color = Some(parse_barcode_color(value, field)?),
                    "b" => barcode.background_color = Some(parse_barcode_color(value, field)?),
                    _ => unreachable!(),
                }
            }
            name if argument.is_some() => {
                return Err(format!(
                    "field {field} switch \\{name} does not take an argument"
                ));
            }
            _ => return Err(format!("field {field} uses unsupported switch \\{name}")),
        }
    }
    Ok(barcode)
}

fn parse_barcode_kind(value: &str) -> Option<BarcodeKind> {
    match value.to_ascii_uppercase().as_str() {
        "UPCA" => Some(BarcodeKind::Upca),
        "UPCE" => Some(BarcodeKind::Upce),
        "JAN13" => Some(BarcodeKind::Jan13),
        "JAN8" => Some(BarcodeKind::Jan8),
        "EAN13" => Some(BarcodeKind::Ean13),
        "EAN8" => Some(BarcodeKind::Ean8),
        "CASE" => Some(BarcodeKind::Case),
        "ITF14" => Some(BarcodeKind::Itf14),
        "NW7" => Some(BarcodeKind::Nw7),
        "CODE39" => Some(BarcodeKind::Code39),
        "CODE128" => Some(BarcodeKind::Code128),
        "JPPOST" => Some(BarcodeKind::JpPost),
        "QR" => Some(BarcodeKind::Qr),
        _ => None,
    }
}

fn barcode_kind_name(kind: BarcodeKind) -> &'static str {
    match kind {
        BarcodeKind::Upca => "UPCA",
        BarcodeKind::Upce => "UPCE",
        BarcodeKind::Jan13 => "JAN13",
        BarcodeKind::Jan8 => "JAN8",
        BarcodeKind::Ean13 => "EAN13",
        BarcodeKind::Ean8 => "EAN8",
        BarcodeKind::Case => "CASE",
        BarcodeKind::Itf14 => "ITF14",
        BarcodeKind::Nw7 => "NW7",
        BarcodeKind::Code39 => "CODE39",
        BarcodeKind::Code128 => "CODE128",
        BarcodeKind::JpPost => "JPPOST",
        BarcodeKind::Qr => "QR",
    }
}

fn parse_point_of_sale_style(value: &str) -> Option<BarcodePointOfSaleStyle> {
    match value.to_ascii_uppercase().as_str() {
        "STD" => Some(BarcodePointOfSaleStyle::Standard),
        "SUP2" => Some(BarcodePointOfSaleStyle::SupplementalTwoDigit),
        "SUP5" => Some(BarcodePointOfSaleStyle::SupplementalFiveDigit),
        "CASE" => Some(BarcodePointOfSaleStyle::Case),
        _ => None,
    }
}

fn parse_case_style(value: &str) -> Option<BarcodeCaseStyle> {
    match value.to_ascii_uppercase().as_str() {
        "STD" => Some(BarcodeCaseStyle::Standard),
        "EXT" => Some(BarcodeCaseStyle::Extended),
        "ADD" => Some(BarcodeCaseStyle::Add),
        _ => None,
    }
}

fn parse_unsigned_integer(
    value: &str,
    field: &str,
    label: &str,
) -> std::result::Result<u32, String> {
    value
        .parse::<u32>()
        .map_err(|_| format!("{field} {label} must be a nonnegative integer"))
}

fn parse_bounded_integer(
    value: &str,
    minimum: u32,
    maximum: u32,
    field: &str,
    label: &str,
) -> std::result::Result<u32, String> {
    value
        .parse::<u32>()
        .ok()
        .filter(|value| (minimum..=maximum).contains(value))
        .ok_or_else(|| format!("{field} {label} must be from {minimum} through {maximum}"))
}

fn validate_barcode_value(
    field: &str,
    kind: BarcodeKind,
    value: &str,
) -> std::result::Result<(), String> {
    let digit_range = match kind {
        BarcodeKind::Ean8 | BarcodeKind::Jan8 => Some(7..=8),
        BarcodeKind::Ean13 | BarcodeKind::Jan13 => Some(12..=13),
        BarcodeKind::Upca => Some(11..=12),
        BarcodeKind::Upce => Some(6..=8),
        BarcodeKind::Case | BarcodeKind::Itf14 => Some(13..=14),
        _ => None,
    };
    if let Some(range) = digit_range
        && (!value.bytes().all(|byte| byte.is_ascii_digit()) || !range.contains(&value.len()))
    {
        return Err(format!(
            "{field} {} value has an invalid digit count or character",
            barcode_kind_name(kind)
        ));
    }
    if kind == BarcodeKind::Code39
        && !value.bytes().all(|byte| {
            byte.is_ascii_uppercase()
                || byte.is_ascii_digit()
                || matches!(byte, b' ' | b'-' | b'.' | b'$' | b'/' | b'+' | b'%')
        })
    {
        return Err(format!(
            "{field} CODE39 value contains an unsupported character"
        ));
    }
    Ok(())
}

fn parse_barcode_color(value: &str, field: &str) -> std::result::Result<u32, String> {
    let parsed = if let Some(hexadecimal) = value
        .strip_prefix("0x")
        .or_else(|| value.strip_prefix("0X"))
    {
        u32::from_str_radix(hexadecimal, 16).ok()
    } else {
        value.parse::<u32>().ok()
    };
    parsed
        .filter(|value| *value <= 0xFF_FFFF)
        .ok_or_else(|| format!("{field} barcode colour must be from 0 through 0xFFFFFF"))
}

struct FormulaParser {
    characters: Vec<char>,
    index: usize,
    token_count: usize,
    depth: usize,
}

impl FormulaParser {
    fn new(input: &str) -> std::result::Result<Self, String> {
        if input.len() > MAX_FORMULA_BYTES {
            return Err("formula exceeds the 4096-byte limit".to_owned());
        }
        if !input.is_ascii() {
            return Err("formula contains unsupported non-ASCII syntax".to_owned());
        }
        Ok(Self {
            characters: input.chars().collect(),
            index: 0,
            token_count: 0,
            depth: 0,
        })
    }

    fn parse(mut self) -> std::result::Result<f64, String> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if self.index != self.characters.len() {
            return Err("formula contains unsupported or trailing syntax".to_owned());
        }
        if !value.is_finite() {
            return Err("formula result is outside the finite numeric range".to_owned());
        }
        Ok(value)
    }

    fn parse_comparison(&mut self) -> std::result::Result<f64, String> {
        let left = self.parse_additive()?;
        self.skip_whitespace();
        let operator = ["<=", ">=", "<>", "=", "<", ">"]
            .into_iter()
            .find(|operator| self.remaining().starts_with(operator));
        let Some(operator) = operator else {
            return Ok(left);
        };
        self.index += operator.len();
        self.count_token()?;
        let right = self.parse_additive()?;
        Ok(
            if match operator {
                "=" => left == right,
                "<>" => left != right,
                "<" => left < right,
                "<=" => left <= right,
                ">" => left > right,
                ">=" => left >= right,
                _ => unreachable!(),
            } {
                1.0
            } else {
                0.0
            },
        )
    }

    fn parse_additive(&mut self) -> std::result::Result<f64, String> {
        let mut value = self.parse_multiplicative()?;
        loop {
            self.skip_whitespace();
            let operator = self.peek();
            if !matches!(operator, Some('+') | Some('-')) {
                return Ok(value);
            }
            self.index += 1;
            self.count_token()?;
            let right = self.parse_multiplicative()?;
            value = if operator == Some('+') {
                value + right
            } else {
                value - right
            };
            self.ensure_finite(value)?;
        }
    }

    fn parse_multiplicative(&mut self) -> std::result::Result<f64, String> {
        let mut value = self.parse_power()?;
        loop {
            self.skip_whitespace();
            let operator = self.peek();
            if !matches!(operator, Some('*') | Some('/')) {
                return Ok(value);
            }
            self.index += 1;
            self.count_token()?;
            let right = self.parse_power()?;
            if operator == Some('/') && right == 0.0 {
                return Err("formula divides by zero".to_owned());
            }
            value = match operator {
                Some('*') => value * right,
                Some('/') => value / right,
                _ => unreachable!(),
            };
            self.ensure_finite(value)?;
        }
    }

    fn parse_power(&mut self) -> std::result::Result<f64, String> {
        let value = self.parse_unary()?;
        self.skip_whitespace();
        if self.peek() != Some('^') {
            return Ok(value);
        }
        self.index += 1;
        self.count_token()?;
        let exponent = self.parse_power()?;
        let result = value.powf(exponent);
        self.ensure_finite(result)?;
        Ok(result)
    }

    fn parse_unary(&mut self) -> std::result::Result<f64, String> {
        self.skip_whitespace();
        match self.peek() {
            Some('+') => {
                self.index += 1;
                self.count_token()?;
                self.parse_unary()
            }
            Some('-') => {
                self.index += 1;
                self.count_token()?;
                Ok(-self.parse_unary()?)
            }
            _ => self.parse_percentage(),
        }
    }

    fn parse_percentage(&mut self) -> std::result::Result<f64, String> {
        let mut value = self.parse_primary()?;
        loop {
            self.skip_whitespace();
            if self.peek() != Some('%') {
                return Ok(value);
            }
            self.index += 1;
            self.count_token()?;
            value /= 100.0;
            self.ensure_finite(value)?;
        }
    }

    fn parse_primary(&mut self) -> std::result::Result<f64, String> {
        self.skip_whitespace();
        if self.peek() == Some('(') {
            self.index += 1;
            self.count_token()?;
            self.depth += 1;
            if self.depth > MAX_FORMULA_DEPTH {
                return Err("formula exceeds the 32-level nesting limit".to_owned());
            }
            let value = self.parse_comparison()?;
            self.skip_whitespace();
            if self.peek() != Some(')') {
                return Err("formula has an unclosed parenthesis".to_owned());
            }
            self.index += 1;
            self.count_token()?;
            self.depth -= 1;
            return Ok(value);
        }
        let start = self.index;
        while self
            .peek()
            .is_some_and(|character| character.is_ascii_digit() || matches!(character, '.' | ','))
        {
            self.index += 1;
        }
        if start == self.index {
            return Err(
                if self
                    .peek()
                    .is_some_and(|character| character.is_ascii_alphabetic())
                {
                    "formula functions are unsupported".to_owned()
                } else {
                    "formula requires a numeric operand".to_owned()
                },
            );
        }
        self.count_token()?;
        let number = self.characters[start..self.index]
            .iter()
            .filter(|character| **character != ',')
            .collect::<String>();
        number
            .parse::<f64>()
            .ok()
            .filter(|value| value.is_finite())
            .ok_or_else(|| "formula contains an invalid or out-of-range number".to_owned())
    }

    fn remaining(&self) -> String {
        self.characters[self.index..].iter().collect()
    }

    fn peek(&self) -> Option<char> {
        self.characters.get(self.index).copied()
    }

    fn skip_whitespace(&mut self) {
        while self.peek().is_some_and(char::is_whitespace) {
            self.index += 1;
        }
    }

    fn count_token(&mut self) -> std::result::Result<(), String> {
        self.token_count += 1;
        if self.token_count > MAX_FORMULA_TOKENS {
            Err("formula exceeds the 512-token limit".to_owned())
        } else {
            Ok(())
        }
    }

    fn ensure_finite(&self, value: f64) -> std::result::Result<(), String> {
        if value.is_finite() {
            Ok(())
        } else {
            Err("formula result is outside the finite numeric range".to_owned())
        }
    }
}

fn format_formula_value(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        return format!("{value:.0}");
    }
    format!("{value:.14e}")
        .parse::<f64>()
        .unwrap_or(value)
        .to_string()
}

fn text_argument(instruction: &FieldInstruction, index: usize) -> Option<&str> {
    instruction.arguments.get(index).and_then(argument_text)
}

fn argument_text(argument: &FieldArgument) -> Option<&str> {
    match argument {
        FieldArgument::Text(value) => Some(value),
        FieldArgument::Nested(_) => None,
    }
}

fn has_switch(instruction: &FieldInstruction, name: &str) -> bool {
    instruction
        .switches
        .iter()
        .any(|switch| switch.name == name)
}

fn switch_text<'a>(instruction: &'a FieldInstruction, name: &str) -> Option<&'a str> {
    instruction.switches.iter().find_map(|switch| {
        (switch.name == name)
            .then_some(switch.argument.as_ref())
            .flatten()
            .and_then(argument_text)
    })
}

fn unsupported_switch(instruction: &FieldInstruction) -> Option<&str> {
    let allowed: &[&str] = match instruction.name.as_str() {
        "PAGE" | "NUMPAGES" | "SECTION" | "SECTIONPAGES" => &["*", "#"],
        "REF" => &["h", "n", "r", "t", "w", "p", "f", "d", "*", "#"],
        "PAGEREF" => &["h", "p", "*", "#"],
        "IF" => &["*", "#"],
        "SEQ" => &["n", "c", "h", "r", "s", "*", "#"],
        "DOCPROPERTY" | "DOCVARIABLE" => &["*", "#", "@"],
        "STYLEREF" => &["l", "n", "r", "t", "w", "p", "*", "#"],
        "INCLUDETEXT" => &["!", "c", "*", "#"],
        "DATE" | "TIME" => &["@", "*"],
        "FILENAME" => &["p", "*"],
        "AUTHOR" => &["*"],
        "MERGEFIELD" => &["b", "f", "m", "v", "*", "#", "@"],
        "=" => &["*", "#"],
        "NEXT" | "NEXTIF" | "SKIPIF" | "MERGEREC" | "MERGESEQ" => &[],
        _ => return None,
    };
    instruction
        .switches
        .iter()
        .find(|switch| !allowed.contains(&switch.name.as_str()))
        .map(|switch| switch.name.as_str())
}

fn validate_instruction_shape(instruction: &FieldInstruction) -> std::result::Result<(), String> {
    if !instruction.quotes_are_balanced() {
        return Err(format!("field {} has unclosed quoting", instruction.name));
    }
    let argument_range = match instruction.name.as_str() {
        "PAGE" | "NUMPAGES" | "SECTION" | "SECTIONPAGES" | "DATE" | "TIME" | "FILENAME"
        | "AUTHOR" => 0..=0,
        "REF" | "PAGEREF" | "SEQ" | "DOCPROPERTY" | "DOCVARIABLE" | "STYLEREF" | "MERGEFIELD" => {
            1..=1
        }
        "IF" => 5..=5,
        "INCLUDETEXT" => 1..=2,
        "=" => 1..=usize::MAX,
        "TOC" => 0..=0,
        "TC" => 1..=1,
        "DISPLAYBARCODE" | "MERGEBARCODE" => 2..=2,
        "NEXT" | "MERGEREC" | "MERGESEQ" => 0..=0,
        "NEXTIF" | "SKIPIF" => 3..=3,
        _ => return Ok(()),
    };
    if !argument_range.contains(&instruction.arguments.len()) {
        let expected = if argument_range.start() == argument_range.end() {
            argument_range.start().to_string()
        } else {
            format!(
                "{} through {}",
                argument_range.start(),
                argument_range.end()
            )
        };
        return Err(format!(
            "field {} requires {expected} positional operands",
            instruction.name
        ));
    }

    if matches!(
        instruction.name.as_str(),
        "TOC" | "TC" | "DISPLAYBARCODE" | "MERGEBARCODE"
    ) {
        return Ok(());
    }

    for switch in &instruction.switches {
        let requires_text = matches!(switch.name.as_str(), "*" | "#" | "@")
            || instruction.name == "REF" && switch.name == "d"
            || instruction.name == "SEQ" && matches!(switch.name.as_str(), "r" | "s")
            || instruction.name == "MERGEFIELD" && matches!(switch.name.as_str(), "b" | "f")
            || instruction.name == "INCLUDETEXT" && switch.name == "c";
        match (&switch.argument, requires_text) {
            (Some(FieldArgument::Text(_)), true) | (None, false) => {}
            (_, true) => {
                return Err(format!(
                    "field {} switch \\{} requires a text argument",
                    instruction.name, switch.name
                ));
            }
            (Some(_), false) => {
                return Err(format!(
                    "field {} switch \\{} does not take an argument",
                    instruction.name, switch.name
                ));
            }
        }
    }
    Ok(())
}

fn compare_if(left: &str, operator: &str, right: &str) -> Option<bool> {
    let numeric = left
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
        .zip(right.parse::<f64>().ok().filter(|value| value.is_finite()));
    let ordering = numeric
        .map(|(left, right)| left.total_cmp(&right))
        .unwrap_or_else(|| left.to_lowercase().cmp(&right.to_lowercase()));
    match operator {
        "=" => Some(if right.contains(['*', '?']) {
            wildcard_matches(right, left)
        } else {
            ordering.is_eq()
        }),
        "<>" => Some(if right.contains(['*', '?']) {
            !wildcard_matches(right, left)
        } else {
            !ordering.is_eq()
        }),
        "<" => Some(ordering.is_lt()),
        "<=" => Some(!ordering.is_gt()),
        ">" => Some(ordering.is_gt()),
        ">=" => Some(!ordering.is_lt()),
        _ => None,
    }
}

fn wildcard_matches(pattern: &str, value: &str) -> bool {
    let mut expression = String::from("(?is)^");
    for character in pattern.chars() {
        match character {
            '*' => expression.push_str(".*"),
            '?' => expression.push('.'),
            other => expression.push_str(&regex::escape(&other.to_string())),
        }
    }
    expression.push('$');
    regex::Regex::new(&expression).is_ok_and(|regex| regex.is_match(value))
}

fn custom_property_text(value: &CustomPropertyValue) -> Option<String> {
    match value {
        CustomPropertyValue::Lpstr(value)
        | CustomPropertyValue::Lpwstr(value)
        | CustomPropertyValue::FileTime(value) => Some(value.clone()),
        CustomPropertyValue::I4(value) => Some(value.to_string()),
        CustomPropertyValue::R8(value) if value.is_finite() => Some(value.to_string()),
        CustomPropertyValue::Bool(value) => Some(value.to_string()),
        CustomPropertyValue::Empty => Some(String::new()),
        CustomPropertyValue::R8(_) | CustomPropertyValue::Raw(_) => None,
    }
}

fn normalized_name(value: &str) -> String {
    value
        .chars()
        .filter(|character| character.is_ascii_alphanumeric())
        .flat_map(char::to_lowercase)
        .collect()
}

fn lexical_file_name(path: &str) -> Option<&str> {
    path.rsplit(['/', '\\'])
        .find(|component| !component.is_empty())
}

fn apply_formats(
    instruction: &FieldInstruction,
    value: &str,
    date_time: Option<FieldDateTime>,
) -> std::result::Result<String, String> {
    let mut output = value.to_owned();
    if let Some(picture) = switch_text(instruction, "#") {
        output = rdocx_layout::engine::format_numeric_field_picture(value, picture)?;
    }
    if let Some(picture) = switch_text(instruction, "@") {
        let date_time = if matches!(instruction.name.as_str(), "DATE" | "TIME") {
            date_time.ok_or_else(|| {
                "date-time formatting requires an explicit date and time".to_owned()
            })?
        } else {
            parse_field_date_time(value).ok_or_else(|| {
                "date-time field value is not a supported civil date and time".to_owned()
            })?
        };
        output = format_date_time(date_time, picture)?;
    }
    rdocx_layout::engine::format_numeric_field_general(instruction, &output)
}

fn valid_date_time(value: FieldDateTime) -> bool {
    value.month >= 1
        && value.month <= 12
        && value.day >= 1
        && value.day <= days_in_month(value.year, value.month)
        && value.hour < 24
        && value.minute < 60
        && value.second < 60
}

fn parse_field_date_time(value: &str) -> Option<FieldDateTime> {
    let value = value.trim().strip_suffix('Z').unwrap_or(value.trim());
    let (date, time) = value
        .split_once('T')
        .or_else(|| value.split_once(' '))
        .map_or((value, None), |(date, time)| (date, Some(time)));
    let mut date_parts = date.split('-');
    let year = date_parts.next()?.parse().ok()?;
    let month = date_parts.next()?.parse().ok()?;
    let day = date_parts.next()?.parse().ok()?;
    if date_parts.next().is_some() {
        return None;
    }
    let (hour, minute, second) = if let Some(time) = time {
        let mut time_parts = time.split(':');
        let hour = time_parts.next()?.parse().ok()?;
        let minute = time_parts.next()?.parse().ok()?;
        let second = time_parts.next()?.split('.').next()?.parse().ok()?;
        if time_parts.next().is_some() {
            return None;
        }
        (hour, minute, second)
    } else {
        (0, 0, 0)
    };
    let parsed = FieldDateTime {
        year,
        month,
        day,
        hour,
        minute,
        second,
    };
    valid_date_time(parsed).then_some(parsed)
}

fn days_in_month(year: i32, month: u8) -> u8 {
    match month {
        4 | 6 | 9 | 11 => 30,
        2 if year % 400 == 0 || year % 4 == 0 && year % 100 != 0 => 29,
        2 => 28,
        _ => 31,
    }
}

fn format_date_time(value: FieldDateTime, picture: &str) -> std::result::Result<String, String> {
    if !valid_date_time(value) {
        return Err("field date and time is invalid".to_owned());
    }
    let months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    let weekdays = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let characters = picture.chars().collect::<Vec<_>>();
    let mut output = String::new();
    let mut index = 0;
    while index < characters.len() {
        if characters[index] == '"' {
            index += 1;
            while index < characters.len() && characters[index] != '"' {
                output.push(characters[index]);
                index += 1;
            }
            if index == characters.len() {
                return Err("date-time picture has unclosed quoting".to_owned());
            }
            index += 1;
            continue;
        }
        if characters[index] == '\\' {
            index += 1;
            let Some(character) = characters.get(index) else {
                return Err("date-time picture ends with an escape".to_owned());
            };
            output.push(*character);
            index += 1;
            continue;
        }
        let rest = characters[index..].iter().collect::<String>();
        if rest.to_ascii_uppercase().starts_with("AM/PM") {
            output.push_str(if value.hour < 12 { "AM" } else { "PM" });
            index += 5;
            continue;
        }
        let token = characters[index];
        if !matches!(token, 'y' | 'M' | 'd' | 'H' | 'h' | 'm' | 's') {
            output.push(token);
            index += 1;
            continue;
        }
        let mut count = 1;
        while index + count < characters.len() && characters[index + count] == token {
            count += 1;
        }
        match token {
            'y' if count == 2 => output.push_str(&format!("{:02}", value.year.rem_euclid(100))),
            'y' => output.push_str(&format!("{:04}", value.year)),
            'M' if count == 1 => output.push_str(&value.month.to_string()),
            'M' if count == 2 => output.push_str(&format!("{:02}", value.month)),
            'M' if count == 3 => output.push_str(&months[value.month as usize - 1][..3]),
            'M' => output.push_str(months[value.month as usize - 1]),
            'd' if count == 1 => output.push_str(&value.day.to_string()),
            'd' if count == 2 => output.push_str(&format!("{:02}", value.day)),
            'd' if count == 3 => output.push_str(&weekdays[weekday(value)][..3]),
            'd' => output.push_str(weekdays[weekday(value)]),
            'H' if count == 1 => output.push_str(&value.hour.to_string()),
            'H' => output.push_str(&format!("{:02}", value.hour)),
            'h' if count == 1 => output.push_str(&(value.hour % 12).max(1).to_string()),
            'h' => output.push_str(&format!("{:02}", (value.hour % 12).max(1))),
            'm' if count == 1 => output.push_str(&value.minute.to_string()),
            'm' => output.push_str(&format!("{:02}", value.minute)),
            's' if count == 1 => output.push_str(&value.second.to_string()),
            's' => output.push_str(&format!("{:02}", value.second)),
            _ => unreachable!(),
        }
        index += count;
    }
    Ok(output)
}

fn weekday(value: FieldDateTime) -> usize {
    let mut year = i64::from(value.year);
    let mut month = i64::from(value.month);
    if month < 3 {
        month += 12;
        year -= 1;
    }
    let year_of_century = year.rem_euclid(100);
    let century = year.div_euclid(100);
    let h = (i64::from(value.day)
        + (13 * (month + 1)) / 5
        + year_of_century
        + year_of_century / 4
        + century / 4
        + 5 * century)
        .rem_euclid(7);
    ((h + 6) % 7) as usize
}

#[cfg(test)]
mod tests {
    #[test]
    fn f283_navigation_disjoint_generated_owners_have_bounded_visits() {
        const OWNERS: usize = 64;
        let mut body = String::new();
        for index in 0..OWNERS {
            body.push_str(&format!(r#"<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>CACHED {index}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>REAL {index}</w:t></w:r></w:p>"#));
        }
        let mut document = Document::new();
        document.document = CT_Document::from_xml(
            format!(r#"<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>"#)
                .as_bytes(),
        )
        .unwrap();
        NAVIGATION_OWNER_VISITS.set(0);
        let outline = document.numbered_document_outline().unwrap();
        let visits = NAVIGATION_OWNER_VISITS.get();
        assert_eq!(outline.len(), OWNERS);
        assert!(
            outline
                .iter()
                .enumerate()
                .all(|(index, node)| node.text == format!("REAL {index}"))
        );
        assert!(
            visits <= 32 * OWNERS,
            "{visits} owner comparisons for {OWNERS} disjoint owners"
        );
    }

    #[test]
    fn f283_navigation_same_paragraph_generated_owners_have_bounded_visits() {
        const OWNERS: usize = 64;
        let owner = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>CACHED</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
        let mut document = Document::new();
        document.document = CT_Document::from_xml(format!(r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>{}</w:p><w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>REAL</w:t></w:r></w:p></w:body></w:document>"#, owner.repeat(OWNERS)).as_bytes()).unwrap();
        NAVIGATION_OWNER_VISITS.set(0);
        let outline = document.numbered_document_outline().unwrap();
        let visits = NAVIGATION_OWNER_VISITS.get();
        assert_eq!(outline.len(), 1);
        assert_eq!(outline[0].text, "REAL");
        assert!(
            visits <= 48 * OWNERS,
            "{visits} owner comparisons for {OWNERS} same-paragraph owners"
        );
    }

    #[test]
    fn fx222_evaluator_reuses_one_owned_paragraph_inventory() {
        let mut document = Document::new();
        document
            .add_paragraph("")
            .add_run("")
            .add_field("DATE", "old")
            .unwrap();
        let mut table = document.add_table(32, 8);
        for row in 0..32 {
            for column in 0..8 {
                table
                    .cell(row, column)
                    .unwrap()
                    .set_text(&format!("cell {row} {column}"));
            }
        }
        let document = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        crate::document::STORY_SOURCE_BUILDS.set(0);
        let context = FieldEvaluationContext::default();
        let evaluator = Evaluator::new(&document, &context);
        assert!(evaluator.bookmarks.is_empty());
        assert_eq!(crate::document::STORY_SOURCE_BUILDS.get(), 1);
    }

    use rdocx_oxml::document::BodyContent;
    use rdocx_oxml::properties::CT_PPr;
    use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc};
    use rdocx_oxml::text::{CT_P, CT_R, Field, FieldSwitch, RunContent};

    use super::*;

    #[test]
    fn fragment_style_numbering_reuse_retains_opaque_reference_facts() {
        let word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let build = |num_id, abstract_id, keep: &str| {
            let mut document = Document::new();
            document.add_list_definition(&[crate::ListLevel::decimal()]);
            let numbering = format!(
                r#"<w:numbering xmlns:w="{word}"><w:abstractNum w:abstractNumId="{abstract_id}"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num></w:numbering>"#
            );
            document.numbering =
                Some(rdocx_oxml::numbering::CT_Numbering::from_xml(numbering.as_bytes()).unwrap());
            let styles = format!(
                r#"<w:styles xmlns:w="{word}" xmlns:f="urn:fragment"><w:style w:type="paragraph" w:styleId="ImportedNumberStyle" f:owner="same"><w:name w:val="Imported"/><w:pPr f:keep="same"><w:numPr><w:numId w:val="{num_id}" f:keep="{keep}"><f:payload value="kept&#32;exact"><w:numId w:val="9"/></f:payload></w:numId></w:numPr></w:pPr></w:style></w:styles>"#
            );
            document.styles = rdocx_oxml::styles::CT_Styles::from_xml(styles.as_bytes()).unwrap();
            document
        };
        let destination = build(1, 12, "A");
        let mut source = build(9, 7, "A");
        assert_eq!(
            destination.equivalent_numbering_dependency_id(&source, 9),
            Some(1)
        );
        assert_eq!(
            fragment_style_numbering_projection(&source.styles.styles[0], Some(1)).unwrap(),
            fragment_style_numbering_projection(&destination.styles.styles[0], None).unwrap(),
        );
        let before = source.styles.to_xml().unwrap();
        assert!(
            !remap_fragment_style_collisions(
                &destination,
                &mut source,
                &mut Vec::new(),
                true,
                true,
                false,
            )
            .unwrap()
        );
        assert_eq!(source.styles.to_xml().unwrap(), before);
        assert!(source.styles.get_by_id("ImportedNumberStyle").is_some());
        let mut changed = build(9, 7, "B");
        assert!(
            remap_fragment_style_collisions(
                &destination,
                &mut changed,
                &mut Vec::new(),
                true,
                true,
                false,
            )
            .unwrap()
        );
        assert!(
            changed
                .styles
                .get_by_id("ImportedNumberStyleMerge1")
                .is_some()
        );
        let xml = String::from_utf8(changed.styles.to_xml().unwrap()).unwrap();
        assert!(xml.contains("f:keep=\"B\""), "{xml}");
        assert!(xml.contains("kept&#32;exact"), "{xml}");
    }

    #[test]
    fn pagination_cache_failure_keeps_every_receiver_field_and_part() {
        let mut document = Document::new();
        for name in ["PAGE", "NUMPAGES"] {
            document
                .add_paragraph("")
                .add_run("")
                .add_field(name, "OLD")
                .unwrap();
        }
        document.set_header("unchanged physical story");
        let before = document.to_bytes().unwrap();
        let updates = [
            Some(CachedFieldUpdate {
                typed_runs: None,
                comment_ranges: Vec::new(),
                cached_result: "1".into(),
                dirty: false,
            }),
            Some(CachedFieldUpdate {
                typed_runs: None,
                comment_ranges: Vec::new(),
                cached_result: "\0".into(),
                dirty: false,
            }),
        ];
        assert!(
            document
                .apply_cached_field_updates(&updates, false)
                .is_err()
        );
        assert_eq!(document.to_bytes().unwrap(), before);
    }

    #[test]
    fn toc_number_stop_saturates_for_a_producer_marker_beyond_the_twip_range() {
        let marker = "W".repeat(3_000_000);
        let units = 3_000_000u64 * 1827;
        let expected_steps = 1 + units.div_ceil(2048);
        assert_eq!(
            toc_number_tab_stop(&marker, 1),
            (expected_steps * 240) as i32
        );

        let marker = "W".repeat(12_000_000);
        assert_eq!(toc_number_tab_stop(&marker, 1), i32::MAX / 240 * 240);
    }

    #[test]
    fn toc_number_stops_match_the_ones_word_writes() {
        // Word 16 for Mac, TOC fields updated in documents whose fonts and
        // styles vary: the stop depends only on the number and the level.
        for (marker, level, stop) in [
            ("1", 1, 480),
            ("1.", 1, 480),
            ("(1)", 1, 720),
            ("10.", 1, 720),
            ("1.1.", 1, 720),
            ("iiii1", 1, 720),
            ("1.1.1.", 1, 960),
            ("WW1", 1, 960),
            ("1.1.1.1.", 1, 1200),
            ("WWW1", 1, 1200),
            ("WWWWW1", 1, 1440),
            ("Chapter I:", 1, 1440),
            ("Section 1.2", 1, 1440),
            ("1.1.1.1.1.", 1, 1440),
            ("iiiiiiiiiiiiiiii1", 1, 1440),
            ("WWWWWWW1", 1, 1920),
            ("1.1", 2, 960),
            ("1.1.1", 3, 1440),
        ] {
            assert_eq!(toc_number_tab_stop(marker, level), stop, "{marker}");
        }
    }

    fn document_with_fields(fields: &[(&str, &str)]) -> Document {
        let mut document = Document::new();
        let mut paragraph = CT_P::new();
        for (instruction, cached_result) in fields {
            paragraph.runs.push(CT_R {
                properties: None,
                content: vec![RunContent::Field(Field::new(instruction, cached_result))],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            });
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(paragraph)));
        document
    }

    fn document_with_parsed_paragraph(xml: &str) -> Document {
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let paragraph = loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Event::Eof => panic!("missing paragraph"),
                _ => {}
            }
            buffer.clear();
        };
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(paragraph)));
        document
    }

    #[test]
    fn fragment_descendant_allocation_ignores_relationship_traversal_order() {
        let copy = |reverse: bool| {
            let mut source = Document::new().package;
            source.set_part("/word/media/root.bin", b"root".to_vec());
            source.set_part("/word/media/a.bin", b"a".to_vec());
            source.set_part("/word/media/a-merge-1.bin", b"a-merge-1".to_vec());
            for part in [
                "/word/media/root.bin",
                "/word/media/a.bin",
                "/word/media/a-merge-1.bin",
            ] {
                source
                    .content_types
                    .add_override(part, "application/octet-stream");
            }
            let relationships = source.get_or_create_part_rels("/word/media/root.bin");
            let items = if reverse {
                [("rId2", "a-merge-1.bin"), ("rId1", "a.bin")]
            } else {
                [("rId1", "a.bin"), ("rId2", "a-merge-1.bin")]
            };
            for (id, target) in items {
                relationships.add_with_id(id, "urn:rdocx:test:child", target);
            }

            let mut seed = Document::new();
            seed.package
                .set_part("/word/media/a.bin", b"occupied".to_vec());
            seed.package
                .content_types
                .add_override("/word/media/a.bin", "application/octet-stream");
            let mut seed_bytes = std::io::Cursor::new(Vec::new());
            seed.package.write_to(&mut seed_bytes).unwrap();
            let mut destination = Document::from_bytes(seed_bytes.get_ref()).unwrap();

            let mut closure = HashSet::new();
            discover_fragment_part_closure(&source, "/word/media/root.bin", &mut closure).unwrap();
            let mut closure = closure.into_iter().collect::<Vec<_>>();
            closure.sort();
            let mut part_map = BTreeMap::new();
            for source_part in closure {
                let destination_part = destination
                    .identifiers
                    .reserve_fragment_part_name(&source_part)
                    .unwrap();
                part_map.insert(source_part, destination_part);
            }
            copy_fragment_part(
                &source,
                &mut destination,
                "/word/media/root.bin",
                &part_map,
                &mut HashSet::new(),
            )
            .unwrap();
            let mut output = std::io::Cursor::new(Vec::new());
            destination.package.write_to(&mut output).unwrap();
            (part_map, output.into_inner())
        };

        let forward = copy(false);
        let reverse = copy(true);
        assert_eq!(forward, reverse);
        assert_eq!(
            forward.0.get("/word/media/a.bin").unwrap(),
            "/word/media/a-merge-2.bin"
        );
        assert_eq!(
            forward.0.get("/word/media/a-merge-1.bin").unwrap(),
            "/word/media/a-merge-1.bin"
        );
    }

    #[test]
    fn a_typed_field_is_reported_without_mutating_its_cache() {
        let document = document_with_fields(&[("AUTHOR", "stored")]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 1, "the typed field must be evaluated");
    }

    #[test]
    fn formula_fields_use_bounded_precedence_and_stable_failures() {
        let document = document_with_fields(&[
            ("= 2 + 3 * 4", "precedence"),
            ("= (2 + 3) * 4", "parentheses"),
            ("= 2 ^ 3 ^ 2", "power"),
            ("= 1,000 + .5", "grouping"),
            ("= 50%", "percentage"),
            ("= (50 + 50)%", "grouped percentage"),
            (r#"= 7 / 2 \# "0.00""#, "picture"),
            ("= 1 / 0", "zero"),
            ("= SUM(1, 2)", "function"),
            ("= (1 + 2", "malformed"),
            ("= 1e309", "bounds"),
        ]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results
                .iter()
                .take(7)
                .map(|result| result.outcome.clone())
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("14".to_owned()),
                FieldOutcome::Resolved("20".to_owned()),
                FieldOutcome::Resolved("512".to_owned()),
                FieldOutcome::Resolved("1000.5".to_owned()),
                FieldOutcome::Resolved("0.5".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("3.50".to_owned()),
            ]
        );
        for result in &results[7..] {
            assert!(
                matches!(result.outcome, FieldOutcome::KeepStored { .. }),
                "{} unexpectedly resolved as {:?}",
                result.instruction,
                result.outcome
            );
        }
        assert_eq!(results[7].outcome, keep("formula divides by zero"));
        assert_eq!(
            results[8].outcome,
            keep("formula functions are unsupported")
        );
        assert_eq!(
            results[9].outcome,
            keep("formula has an unclosed parenthesis")
        );
        let invalid_percentage = document_with_fields(&[("= 5 % 2", "stored")]);
        assert_eq!(
            invalid_percentage
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            keep("formula contains unsupported or trailing syntax")
        );
        let normalized_decimal = document_with_fields(&[("= 0.1 + 0.2", "stored")]);
        assert_eq!(
            normalized_decimal
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("0.3".to_owned())
        );
        let compact = format!(
            "= {}",
            std::iter::repeat_n("1", 129).collect::<Vec<_>>().join("+")
        );
        let spaced = format!(
            "= {}",
            std::iter::repeat_n("1", 129)
                .collect::<Vec<_>>()
                .join(" + ")
        );
        for instruction in [compact, spaced] {
            let equivalent = document_with_fields(&[(&instruction, "stored")]);
            assert_eq!(
                equivalent
                    .evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::Resolved("129".to_owned())
            );
        }
        for instruction in [
            format!("= {}", "1".repeat(MAX_FORMULA_BYTES + 1)),
            format!(
                "= {}1{}",
                "(".repeat(MAX_FORMULA_DEPTH + 1),
                ")".repeat(MAX_FORMULA_DEPTH + 1)
            ),
            format!("= {}1", "1+".repeat(MAX_FORMULA_TOKENS)),
        ] {
            let bounded = document_with_fields(&[(&instruction, "bounded fallback")]);
            let result = bounded
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap();
            assert!(matches!(result[0].outcome, FieldOutcome::KeepStored { .. }));
            assert_eq!(result[0].cached_result, "bounded fallback");
        }

        let mut nested = document_with_fields(&[("= 1 + 2", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut nested.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(formula) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        formula.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Amount", "nested")));
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Amount".to_owned(), "3".to_owned())]),
            ..Default::default()
        };
        let results = nested.evaluate_fields(&context).unwrap();
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("5".to_owned()));
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("3".to_owned()));
    }

    #[test]
    fn mail_merge_control_state_is_story_and_record_scoped() {
        let document = document_with_fields(&[
            ("NEXT", "next"),
            ("MERGEREC", "record"),
            (r#"NEXTIF "A" = "B""#, "conditional next"),
            (r#"SKIPIF "A" = "A""#, "conditional skip"),
            ("MERGESEQ", "sequence"),
        ]);
        let context = FieldEvaluationContext {
            merge_record_number: Some(4),
            merge_sequence_number: Some(2),
            ..Default::default()
        };
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(
            results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecord { record_number: 5 }),
                FieldOutcome::MailMergeControl(MailMergeControl::RecordNumber(5)),
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecordIf {
                    condition: false,
                    record_number: 5,
                }),
                FieldOutcome::MailMergeControl(MailMergeControl::SkipRecordIf {
                    condition: true,
                    record_number: 5,
                }),
                FieldOutcome::MailMergeControl(MailMergeControl::SequenceNumber(2)),
            ]
        );

        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            unreachable!()
        };
        let paragraphs = [paragraph.as_ref()];
        let mut evaluator = Evaluator::new(&document, &context);
        evaluator.evaluate_story("main", &paragraphs).unwrap();
        evaluator.evaluate_story("header:one", &paragraphs).unwrap();
        assert_eq!(
            evaluator.results[0].outcome, evaluator.results[5].outcome,
            "each story must start from the explicit record context"
        );

        let unavailable = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(
            unavailable
                .iter()
                .all(|result| matches!(result.outcome, FieldOutcome::KeepStored { .. }))
        );

        let record_only = FieldEvaluationContext {
            merge_record_number: Some(9),
            ..Default::default()
        };
        let record = document_with_fields(&[("MERGEREC", "stored")]);
        assert_eq!(
            record.evaluate_fields(&record_only).unwrap()[0].outcome,
            FieldOutcome::MailMergeControl(MailMergeControl::RecordNumber(9))
        );

        for (instruction, context, diagnostic) in [
            (
                "MERGEREC",
                FieldEvaluationContext {
                    merge_record_number: Some(0),
                    ..Default::default()
                },
                "MERGEREC merge record number must be one-based",
            ),
            (
                "MERGESEQ",
                FieldEvaluationContext {
                    merge_sequence_number: Some(0),
                    ..Default::default()
                },
                "MERGESEQ merge sequence number must be one-based",
            ),
        ] {
            let invalid = document_with_fields(&[(instruction, "stored")]);
            assert_eq!(
                invalid.evaluate_fields(&context).unwrap()[0].outcome,
                keep(diagnostic)
            );
        }
    }

    #[test]
    fn toc_tc_and_barcode_fields_preserve_non_text_results() {
        let document = document_with_fields(&[
            (
                r#"TOC \o "1-3" \t "Heading 1,1,Appendix,2" \f C \b Main \h \u \n "2-3" \p " " \d ":""#,
                "stored toc",
            ),
            (r#"TC "Entry" \l 2 \f C \n"#, "stored tc"),
            (
                r#"DISPLAYBARCODE "0123456789012" EAN13 \h 100 \s 200 \f 0xFF0000 \b 0xFFFFFF \t"#,
                "stored barcode",
            ),
        ]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results[0].outcome,
            FieldOutcome::TableOfContents(TocField {
                heading_levels: Some((1, 3)),
                custom_styles: vec![("Heading 1".to_owned(), 1), ("Appendix".to_owned(), 2)],
                entries: TocEntrySelection::Identifier("C".to_owned()),
                sequence_identifier: None,
                bookmark: Some("Main".to_owned()),
                hyperlink: true,
                use_outline_levels: true,
                omit_page_number_levels: Some((2, 3)),
                page_number_separator: Some(" ".to_owned()),
                entry_page_separator: Some(":".to_owned()),
            })
        );
        assert_eq!(
            results[1].outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: "Entry".to_owned(),
                level: 2,
                table_identifier: Some("C".to_owned()),
                omit_page_number: true,
            })
        );
        assert_eq!(
            results[2].outcome,
            FieldOutcome::Barcode(BarcodeField {
                value: "0123456789012".to_owned(),
                kind: BarcodeKind::Ean13,
                height: Some(100),
                scale: Some(200),
                error_correction: None,
                point_of_sale_style: None,
                case_style: None,
                fix_check_digit: false,
                rotation: None,
                foreground_color: Some(0xFF0000),
                background_color: Some(0xFFFFFF),
                display_text: true,
                add_start_stop: false,
            })
        );

        let unsupported = document_with_fields(&[
            (r#"TOC \o "9-1""#, "toc"),
            (r#"TC "Entry" \l 10"#, "tc"),
            ("DISPLAYBARCODE value UNKNOWN", "barcode"),
            ("DISPLAYBARCODE 123 EAN13", "barcode digits"),
        ]);
        let outcomes = unsupported
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            outcomes
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                keep("TOC heading range starts after it ends"),
                keep("TC level must be from 1 through 9"),
                keep("DISPLAYBARCODE barcode type UNKNOWN is unsupported"),
                keep("DISPLAYBARCODE EAN13 value has an invalid digit count or character"),
            ]
        );

        for (instruction, heading_levels, entries, use_outline_levels) in [
            ("TOC", Some((1, 9)), TocEntrySelection::None, false),
            (r"TOC \o", Some((1, 9)), TocEntrySelection::None, false),
            (r"TOC \f", None, TocEntrySelection::All, false),
            (
                r"TOC \f C",
                None,
                TocEntrySelection::Identifier("C".to_owned()),
                false,
            ),
            (r"TOC \u", None, TocEntrySelection::None, true),
        ] {
            let toc = document_with_fields(&[(instruction, "stored toc")]);
            assert_eq!(
                toc.evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::TableOfContents(TocField {
                    heading_levels,
                    custom_styles: Vec::new(),
                    entries,
                    sequence_identifier: None,
                    bookmark: None,
                    hyperlink: false,
                    use_outline_levels,
                    omit_page_number_levels: None,
                    page_number_separator: None,
                    entry_page_separator: None,
                })
            );
        }

        let normalized_toc =
            document_with_fields(&[(r#"TOC \t "Heading 1,1, Appendix,2" \p ":""#, "stored toc")]);
        assert_eq!(
            normalized_toc
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                heading_levels: None,
                custom_styles: vec![("Heading 1".to_owned(), 1), ("Appendix".to_owned(), 2)],
                entries: TocEntrySelection::None,
                sequence_identifier: None,
                bookmark: None,
                hyperlink: false,
                use_outline_levels: false,
                omit_page_number_levels: None,
                page_number_separator: Some(":".to_owned()),
                entry_page_separator: None,
            })
        );
        let trailing_separator =
            document_with_fields(&[(r#"TOC \t "Heading 1,1,Appendix,2,""#, "stored toc")]);
        assert!(matches!(
            trailing_separator
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                ref custom_styles,
                ..
            }) if custom_styles == &[("Heading 1".to_owned(), 1), ("Appendix".to_owned(), 2)]
        ));
        for malformed in [
            r#"TOC \t "Heading 1,,Appendix,2""#,
            r#"TOC \t "Heading 1,1,,2""#,
        ] {
            let malformed = document_with_fields(&[(malformed, "stored toc")]);
            assert!(matches!(
                malformed
                    .evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::KeepStored { .. }
            ));
        }
        let decorated_default = document_with_fields(&[(r"TOC \h", "stored toc")]);
        assert!(matches!(
            decorated_default
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                heading_levels: Some((1, 9)),
                hyperlink: true,
                ..
            })
        ));
        let invalid_toc = document_with_fields(&[(r#"TOC \p "ab""#, "stored toc")]);
        assert_eq!(
            invalid_toc
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            keep("TOC page-number separator must contain exactly one character")
        );
        let sequenced_toc =
            document_with_fields(&[(r#"TOC \o "1-3" \s chapter \d ":""#, "stored toc")]);
        assert!(matches!(
            sequenced_toc
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                sequence_identifier: Some(ref identifier),
                entry_page_separator: Some(ref separator),
                ..
            }) if identifier == "chapter" && separator == ":"
        ));

        let barcode_grammar = document_with_fields(&[
            (r"DISPLAYBARCODE 0123456789012 EAN13 \x", "fix"),
            (r"DISPLAYBARCODE 0123456789012 EAN13 \p STD", "pos"),
            (r"DISPLAYBARCODE 1234567890123 ITF14 \c EXT", "case"),
            (r"DISPLAYBARCODE value QR \q 3", "correction"),
            (r"DISPLAYBARCODE 0123456789012 EAN13 \p OTHER", "bad pos"),
            (r"DISPLAYBARCODE value QR \q H", "bad correction"),
        ]);
        let outcomes = barcode_grammar
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        let FieldOutcome::Barcode(fix) = &outcomes[0].outcome else {
            panic!("expected check-digit barcode")
        };
        assert!(fix.fix_check_digit);
        let FieldOutcome::Barcode(pos) = &outcomes[1].outcome else {
            panic!("expected point-of-sale barcode")
        };
        assert_eq!(
            pos.point_of_sale_style,
            Some(BarcodePointOfSaleStyle::Standard)
        );
        let FieldOutcome::Barcode(case) = &outcomes[2].outcome else {
            panic!("expected ITF14 case barcode")
        };
        assert_eq!(case.case_style, Some(BarcodeCaseStyle::Extended));
        let FieldOutcome::Barcode(correction) = &outcomes[3].outcome else {
            panic!("expected corrected QR barcode")
        };
        assert_eq!(correction.error_correction, Some(3));
        assert_eq!(
            outcomes[4].outcome,
            keep("DISPLAYBARCODE point-of-sale style must be STD, SUP2, SUP5, or CASE")
        );
        assert_eq!(
            outcomes[5].outcome,
            keep("DISPLAYBARCODE error correction must be from 0 through 3")
        );

        let case_alias = document_with_fields(&[
            (r"DISPLAYBARCODE 1234567890123 CASE \c EXT", "case"),
            (r"DISPLAYBARCODE not-digits CASE", "invalid case"),
        ]);
        let outcomes = case_alias
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(matches!(
            outcomes[0].outcome,
            FieldOutcome::Barcode(BarcodeField {
                kind: BarcodeKind::Case,
                case_style: Some(BarcodeCaseStyle::Extended),
                ..
            })
        ));
        assert_eq!(
            outcomes[1].outcome,
            keep("DISPLAYBARCODE CASE value has an invalid digit count or character")
        );

        let extra_operands = document_with_fields(&[
            (r#"TC "Entry" unexpected"#, "tc"),
            (r"DISPLAYBARCODE value QR unexpected \t", "barcode"),
        ]);
        let outcomes = extra_operands
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            outcomes[0].outcome,
            keep("field TC requires 1 positional operands")
        );
        assert_eq!(
            outcomes[1].outcome,
            keep("field DISPLAYBARCODE requires 2 positional operands")
        );

        let escaped = document_with_fields(&[
            (r#"TC "A\"B""#, "stored quote"),
            (r#"TC "\Entry""#, "stored slash"),
        ]);
        assert_eq!(
            escaped
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: "A\"B".to_owned(),
                level: 1,
                table_identifier: None,
                omit_page_number: false,
            })
        );
        assert_eq!(
            escaped
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[1]
                .outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: r"\Entry".to_owned(),
                level: 1,
                table_identifier: None,
                omit_page_number: false,
            })
        );

        let long_value = "x".repeat(1025);
        let barcode_bounds = document_with_fields(&[
            (&format!("DISPLAYBARCODE {long_value} QR"), "value"),
            (r"DISPLAYBARCODE value QR \h 4294967296", "height"),
            (r"DISPLAYBARCODE value QR \s 9", "scale low"),
            (r"DISPLAYBARCODE value QR \s 1001", "scale high"),
            (r"DISPLAYBARCODE value QR \r 4", "rotation"),
            (r"DISPLAYBARCODE value QR \f 0x1000000", "foreground"),
            (r"DISPLAYBARCODE value QR \b 16777216", "background"),
        ]);
        let outcomes = barcode_bounds
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            outcomes
                .into_iter()
                .map(|outcome| outcome.outcome)
                .collect::<Vec<_>>(),
            [
                keep("DISPLAYBARCODE value must contain from 1 through 1024 characters"),
                keep("DISPLAYBARCODE height must be a nonnegative integer"),
                keep("DISPLAYBARCODE scale must be from 10 through 1000"),
                keep("DISPLAYBARCODE scale must be from 10 through 1000"),
                keep("DISPLAYBARCODE rotation must be from 0 through 3"),
                keep("DISPLAYBARCODE barcode colour must be from 0 through 0xFFFFFF"),
                keep("DISPLAYBARCODE barcode colour must be from 0 through 0xFFFFFF"),
            ]
        );
        let maximum_value = "x".repeat(1024);
        let boundary_instruction = format!(
            "DISPLAYBARCODE {maximum_value} QR \\h 4294967295 \\s 10 \\r 3 \\f 0 \\b 0xFFFFFF"
        );
        let boundaries = document_with_fields(&[(&boundary_instruction, "stored")]);
        assert!(matches!(
            boundaries
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Barcode(BarcodeField {
                height: Some(u32::MAX),
                scale: Some(10),
                rotation: Some(3),
                foreground_color: Some(0),
                background_color: Some(0xFF_FFFF),
                ..
            })
        ));

        let mut nested = document_with_fields(&[("DISPLAYBARCODE placeholder QR", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut nested.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(barcode) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        barcode.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Code", "nested")));
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Code".to_owned(), "nested value".to_owned())]),
            ..Default::default()
        };
        assert!(matches!(
            nested.evaluate_fields(&context).unwrap()[0].outcome,
            FieldOutcome::Barcode(BarcodeField { ref value, .. }) if value == "nested value"
        ));

        let mut nested_tc = document_with_fields(&[("TC placeholder", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut nested_tc.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(tc) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        tc.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Entry", "nested")));
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Entry".to_owned(), "Nested entry".to_owned())]),
            ..Default::default()
        };
        assert!(matches!(
            nested_tc.evaluate_fields(&context).unwrap()[0].outcome,
            FieldOutcome::TableOfContentsEntry(TcField { ref entry, .. })
                if entry == "Nested entry"
        ));

        let mut nested_switches = document_with_fields(&[
            (r"TOC \b placeholder", "toc"),
            (r#"TC "Entry" \f placeholder"#, "tc"),
            (r"DISPLAYBARCODE value QR \s 100", "barcode"),
        ]);
        let BodyContent::Paragraph(paragraph) = &mut nested_switches.document.body.content[0]
        else {
            unreachable!()
        };
        for (index, instruction) in ["MERGEFIELD Scope", "MERGEFIELD Kind", "MERGEFIELD Scale"]
            .into_iter()
            .enumerate()
        {
            let RunContent::Field(field) = &mut paragraph.runs[index].content[0] else {
                unreachable!()
            };
            field.instruction.switches[0].argument = Some(FieldArgument::Nested(Box::new(
                Field::new(instruction, "nested"),
            )));
        }
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([
                ("Scope".to_owned(), "Main".to_owned()),
                ("Kind".to_owned(), "C".to_owned()),
                ("Scale".to_owned(), "250".to_owned()),
            ]),
            ..Default::default()
        };
        let outcomes = nested_switches.evaluate_fields(&context).unwrap();
        assert!(matches!(
            outcomes[0].outcome,
            FieldOutcome::TableOfContents(TocField {
                bookmark: Some(ref bookmark),
                ..
            }) if bookmark == "Main"
        ));
        assert!(matches!(
            outcomes[2].outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                table_identifier: Some(ref identifier),
                ..
            }) if identifier == "C"
        ));
        assert!(matches!(
            outcomes[4].outcome,
            FieldOutcome::Barcode(BarcodeField {
                scale: Some(250),
                ..
            })
        ));
    }

    #[test]
    fn raw_only_instruction_edits_use_the_serialized_instruction() {
        let mut document = document_with_fields(&[("DATE", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        field.instruction.raw = "AUTHOR".to_owned();
        document.set_author("Ada Lovelace");

        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results[0].instruction, "AUTHOR");
        assert_eq!(
            results[0].outcome,
            FieldOutcome::Resolved("Ada Lovelace".to_owned())
        );
    }

    #[test]
    fn same_run_nested_raw_instruction_edit_updates_saves_and_reopens() {
        let word_namespace = rdocx_oxml::namespace::W_NS;
        let xml = format!(
            concat!(
                r#"<w:p xmlns:w="{0}"><q:r xmlns:q="{0}" xmlns:x="urn:producer">"#,
                r#"<q:fldChar q:fldCharType="begin"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText>MERGEFIELD Old</q:instrText><x:nestedInstruction/><q:fldChar q:fldCharType="separate"/><q:t>stored nested</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"</q:r></w:p>"#,
            ),
            word_namespace,
        );
        let mut document = document_with_parsed_paragraph(&xml);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        let Some(FieldArgument::Nested(nested)) = outer.instruction.arguments.first_mut() else {
            panic!("expected nested field")
        };
        nested.instruction.raw = "MERGEFIELD New".to_owned();

        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("New".to_owned(), "new value".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        assert_eq!(document.update_fields(&context).unwrap(), 2);
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains("MERGEFIELD New"), "{saved_xml}");
        assert!(!saved_xml.contains("MERGEFIELD Old"), "{saved_xml}");
        assert_eq!(saved_xml.matches(">new value<").count(), 1, "{saved_xml}");
        assert!(saved_xml.contains(r#"w:dirty="0""#), "{saved_xml}");
        assert!(saved_xml.contains("<x:nestedInstruction/>"), "{saved_xml}");

        let reopened = Document::from_bytes(&saved).unwrap();
        let evaluations = reopened.evaluate_fields(&context).unwrap();
        assert_eq!(evaluations.len(), 2);
        assert_eq!(evaluations[0].cached_result, "yes");
        assert_eq!(evaluations[1].instruction, "MERGEFIELD New");
        assert_eq!(evaluations[1].cached_result, "new value");
        assert_eq!(
            evaluations[1].outcome,
            FieldOutcome::Resolved("new value".to_owned())
        );
    }

    #[test]
    fn raw_only_nested_edit_suppresses_its_stale_nested_operand() {
        let word_namespace = rdocx_oxml::namespace::W_NS;
        let xml = format!(
            concat!(
                r#"<w:p xmlns:w="{0}"><q:r xmlns:q="{0}" xmlns:x="urn:producer">"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText><x:middleBefore/>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText>MERGEFIELD Stale</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored grandchild</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<x:middleAfter/><q:instrText xml:space="preserve"> = &quot;stale&quot; &quot;old yes&quot; &quot;old no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored middle</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"</q:r></w:p>"#,
            ),
            word_namespace,
        );
        let mut document = document_with_parsed_paragraph(&xml);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        let Some(FieldArgument::Nested(middle)) = outer.instruction.arguments.first_mut() else {
            panic!("expected middle field")
        };
        assert!(matches!(
            middle.instruction.arguments.first(),
            Some(FieldArgument::Nested(_))
        ));
        middle.instruction.raw = "MERGEFIELD New".to_owned();

        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("New".to_owned(), "new value".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        let before_save = document.evaluate_fields(&context).unwrap();
        assert_eq!(before_save.len(), 2);
        assert_eq!(before_save[0].field_index, 0);
        assert_eq!(before_save[1].field_index, 1);
        assert_eq!(before_save[1].instruction, "MERGEFIELD New");
        assert_eq!(
            before_save[1].outcome,
            FieldOutcome::Resolved("new value".to_owned())
        );

        assert_eq!(document.update_fields(&context).unwrap(), 2);
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains("MERGEFIELD New"), "{saved_xml}");
        assert!(!saved_xml.contains("MERGEFIELD Stale"), "{saved_xml}");
        assert!(!saved_xml.contains("stored grandchild"), "{saved_xml}");
        assert!(!saved_xml.contains("stored middle"), "{saved_xml}");
        assert!(saved_xml.contains("<x:middleBefore/>"), "{saved_xml}");
        assert!(saved_xml.contains("<x:middleAfter/>"), "{saved_xml}");
        assert!(
            saved_xml.find("<x:middleBefore/>").unwrap()
                < saved_xml.find("<x:middleAfter/>").unwrap(),
            "{saved_xml}"
        );
        assert_eq!(saved_xml.matches(r#"w:dirty="0""#).count(), 2);
        assert_eq!(saved_xml.matches(">new value<").count(), 1, "{saved_xml}");

        let reopened = Document::from_bytes(&saved).unwrap();
        let evaluations = reopened.evaluate_fields(&context).unwrap();
        assert_eq!(evaluations.len(), 2);
        assert_eq!(evaluations[0].field_index, 0);
        assert_eq!(evaluations[0].cached_result, "yes");
        assert_eq!(evaluations[1].field_index, 1);
        assert_eq!(evaluations[1].instruction, "MERGEFIELD New");
        assert_eq!(evaluations[1].cached_result, "new value");
        assert_eq!(
            evaluations[1].outcome,
            FieldOutcome::Resolved("new value".to_owned())
        );
    }

    #[test]
    fn multi_run_raw_only_nested_edit_preserves_every_run_scaffold() {
        let word_namespace = rdocx_oxml::namespace::W_NS;
        let xml = format!(
            concat!(
                r#"<w:p xmlns:w="{0}">"#,
                r#"<q:r xmlns:q="{0}" data-run="outer-start"><q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:a="urn:start" data-run="start"><q:rPr><q:i/></q:rPr><q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText><a:prefix/><q:fldChar q:fldCharType="begin" q:dirty="on"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:m="urn:middle" data-run="middle"><q:rPr><q:u q:val="single"/></q:rPr><q:instrText>MERGEFIELD Stale</q:instrText><m:inside/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:z="urn:end" data-run="end"><q:rPr><q:b/></q:rPr><q:fldChar q:fldCharType="separate"/><q:t>stored grandchild</q:t><q:fldChar q:fldCharType="end"/><z:suffix/><q:instrText xml:space="preserve"> = &quot;stale&quot; &quot;old yes&quot; &quot;old no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored middle</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" data-run="outer-end"><q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
                r#"</w:p>"#,
            ),
            word_namespace,
        );
        let mut document = document_with_parsed_paragraph(&xml);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        let Some(FieldArgument::Nested(middle)) = outer.instruction.arguments.first_mut() else {
            panic!("expected middle field")
        };
        assert!(matches!(
            middle.instruction.arguments.first(),
            Some(FieldArgument::Nested(_))
        ));
        middle.instruction.raw = "MERGEFIELD New".to_owned();

        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("New".to_owned(), "new value".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        assert_eq!(document.update_fields(&context).unwrap(), 2);
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains("MERGEFIELD New"), "{saved_xml}");
        assert!(!saved_xml.contains("MERGEFIELD Stale"), "{saved_xml}");
        assert!(!saved_xml.contains("stored grandchild"), "{saved_xml}");
        for preserved in [
            r#"data-run="start""#,
            r#"data-run="middle""#,
            r#"data-run="end""#,
            r#"xmlns:a="urn:start""#,
            r#"xmlns:m="urn:middle""#,
            r#"xmlns:z="urn:end""#,
            "<q:i/>",
            r#"<q:u q:val="single"/>"#,
            "<q:b/>",
            "<a:prefix/>",
            "<m:inside/>",
            "<z:suffix/>",
        ] {
            assert!(
                saved_xml.contains(preserved),
                "missing {preserved}: {saved_xml}"
            );
        }
        assert!(
            saved_xml.find("<a:prefix/>").unwrap() < saved_xml.find("<m:inside/>").unwrap()
                && saved_xml.find("<m:inside/>").unwrap() < saved_xml.find("<z:suffix/>").unwrap(),
            "{saved_xml}"
        );

        let reopened = Document::from_bytes(&saved).unwrap();
        let evaluations = reopened.evaluate_fields(&context).unwrap();
        assert_eq!(evaluations.len(), 2);
        assert_eq!(evaluations[0].field_index, 0);
        assert_eq!(evaluations[0].cached_result, "yes");
        assert_eq!(evaluations[1].field_index, 1);
        assert_eq!(evaluations[1].instruction, "MERGEFIELD New");
        assert_eq!(evaluations[1].cached_result, "new value");
    }

    #[test]
    fn nested_if_and_comparison_operators_evaluate_recursively() {
        let mut document = document_with_fields(&[(r#"IF left = right yes no"#, "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Score", "stored score")));
        outer.instruction.arguments[1] = FieldArgument::Text(">=".to_owned());
        outer.instruction.arguments[2] = FieldArgument::Text("2".to_owned());
        outer.instruction.arguments[4] =
            FieldArgument::Nested(Box::new(Field::new("UNKNOWN", "stored branch")));
        let mut context = FieldEvaluationContext::default();
        context
            .merge_fields
            .insert("Score".to_owned(), "10".to_owned());
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(results[0].field_index, 0);
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("yes".to_owned()));
        assert_eq!(results[1].field_index, 1);
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("10".to_owned()));
        assert_eq!(results[2].field_index, 2);
        assert!(matches!(
            results[2].outcome,
            FieldOutcome::KeepStored { .. }
        ));

        for (operator, expected) in [
            ("=", "yes"),
            ("<>", "no"),
            ("<", "no"),
            ("<=", "yes"),
            (">", "no"),
            (">=", "yes"),
        ] {
            let document = document_with_fields(&[(
                &format!(r#"IF "2" {operator} "2" "yes" "no""#),
                "stored",
            )]);
            assert_eq!(
                document
                    .evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::Resolved(expected.to_owned())
            );
        }
        let wildcard =
            document_with_fields(&[(r#"IF "Alphabet" = "A?pha*" "yes" "no""#, "stored")]);
        assert_eq!(
            wildcard
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("yes".to_owned())
        );

        let mut unresolved = document_with_fields(&[(r#"IF left = right yes no"#, "outer")]);
        let BodyContent::Paragraph(paragraph) = &mut unresolved.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("DATE", "stored date")));
        let outcomes = unresolved
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(matches!(
            outcomes[0].outcome,
            FieldOutcome::KeepStored { .. }
        ));
        assert!(matches!(
            outcomes[1].outcome,
            FieldOutcome::KeepStored { .. }
        ));
    }

    #[test]
    fn ref_reads_sequence_snapshot_in_the_same_explicit_update() {
        let mut document = document_with_fields(&[
            (r"SEQ Figure \r 7", "OLD"),
            ("SEQ Figure", "STALE-CAPTION"),
            ("REF CaptionNumber", "STALE-REF"),
        ]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        for run in &mut paragraph.runs {
            if let RunContent::Field(field) = &mut run.content[0] {
                *field = Field::from_raw(
                    &field.instruction.raw,
                    rdocx_oxml::text::FieldForm::Complex,
                    vec![CT_R::new(&field.cached_result)],
                )
                .unwrap();
            }
        }
        document
            .add_bookmark(
                "CaptionNumber",
                crate::RunRange {
                    start: crate::RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                    end: crate::RunPosition {
                        body_index: 0,
                        run_index: 2,
                    },
                },
            )
            .unwrap();
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results
                .iter()
                .map(|field| field.outcome.clone())
                .collect::<Vec<_>>(),
            vec![
                FieldOutcome::Resolved("7".into()),
                FieldOutcome::Resolved("8".into()),
                FieldOutcome::Resolved("8".into())
            ]
        );
        assert_eq!(document.bookmarks()[0].text(), "STALE-CAPTION");
        document
            .update_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(document.bookmarks()[0].text(), "8");
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let result = reopened
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(result[2].cached_result, "8");
    }

    #[test]
    fn nested_if_reuses_the_eager_effective_instruction_outcome() {
        let mut document = document_with_fields(&[(r#"IF left = 1 yes no"#, "outer")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("SEQ Figure", "stored sequence")));

        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 2);
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("yes".to_owned()));
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("1".to_owned()));
    }

    #[test]
    fn nested_outcome_frames_do_not_leak_between_outer_fields() {
        let mut document = document_with_fields(&[
            (r#"IF left = 1 first no"#, "first outer"),
            (r#"IF left = 2 second no"#, "second outer"),
        ]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        for run in &mut paragraph.runs {
            let RunContent::Field(outer) = &mut run.content[0] else {
                unreachable!()
            };
            outer.instruction.arguments[0] =
                FieldArgument::Nested(Box::new(Field::new("SEQ Figure", "stored sequence")));
        }

        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 4);
        assert_eq!(
            results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("first".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("second".to_owned()),
                FieldOutcome::Resolved("2".to_owned()),
            ]
        );
    }

    #[test]
    fn every_nested_field_is_reported_before_outer_fallback() {
        let mut document = document_with_fields(&[("UNKNOWN", "outer")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer
            .instruction
            .arguments
            .push(FieldArgument::Nested(Box::new(Field::new(
                "MERGEFIELD Name",
                "stored name",
            ))));
        outer.instruction.switches.push(FieldSwitch {
            name: "*".to_owned(),
            argument: Some(FieldArgument::Nested(Box::new(Field::new(
                "AUTHOR",
                "stored author",
            )))),
        });
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Name".to_owned(), "Ada".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(results.len(), 3);
        assert!(matches!(
            results[0].outcome,
            FieldOutcome::KeepStored { .. }
        ));
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("Ada".to_owned()));
        assert!(matches!(
            results[2].outcome,
            FieldOutcome::KeepStored { .. }
        ));
    }

    #[test]
    fn malformed_arity_and_extreme_inputs_keep_stored_without_panicking() {
        let document = document_with_fields(&[
            (r"DATE \@", "date"),
            (r"SEQ Figure \r", "reset"),
            ("PAGE extra", "page"),
            (r#"MERGEFIELD "Name"#, "merge"),
            (r##"MERGEFIELD Amount \# "$0.00"##, "picture"),
            (r"SEQ Figure \r 9223372036854775807", "max"),
            ("SEQ Figure", "overflow"),
        ]);
        let mut context = FieldEvaluationContext::default();
        context
            .merge_fields
            .insert("Name".to_owned(), "Ada".to_owned());
        context
            .merge_fields
            .insert("Amount".to_owned(), "12".to_owned());
        let results = document.evaluate_fields(&context).unwrap();
        assert!(
            results[..5]
                .iter()
                .all(|result| matches!(result.outcome, FieldOutcome::KeepStored { .. }))
        );
        assert_eq!(
            results[5].outcome,
            FieldOutcome::Resolved(i64::MAX.to_string())
        );
        assert_eq!(results[6].outcome, keep("SEQ value overflowed"));

        let date = document_with_fields(&[(r#"DATE \@ "dddd""#, "date")]);
        let context = FieldEvaluationContext {
            now: Some(FieldDateTime {
                year: i32::MIN,
                month: 1,
                day: 1,
                hour: 0,
                minute: 0,
                second: 0,
            }),
            ..FieldEvaluationContext::default()
        };
        assert!(matches!(
            date.evaluate_fields(&context).unwrap()[0].outcome,
            FieldOutcome::Resolved(_)
        ));
    }

    #[test]
    fn sequence_reset_switches_require_physical_source_context() {
        let document = document_with_fields(&[
            ("SEQ Figure", "0"),
            (r"SEQ Figure \c", "0"),
            (r"SEQ Figure \r 5", "0"),
            (r"SEQ Figure \h", "0"),
            (r"SEQ Figure \n \* ROMAN", "0"),
            ("SEQ Table", "0"),
        ]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("5".to_owned()),
                FieldOutcome::Resolved(String::new()),
                FieldOutcome::Resolved("VII".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
            ]
        );

        let story_document = document_with_fields(&[("SEQ Shared", "0")]);
        let BodyContent::Paragraph(paragraph) = &story_document.document.body.content[0] else {
            unreachable!()
        };
        let paragraphs = [paragraph.as_ref()];
        let story_context = FieldEvaluationContext::default();
        let mut evaluator = Evaluator::new(&story_document, &story_context);
        evaluator.evaluate_story("header:one", &paragraphs).unwrap();
        evaluator.evaluate_story("footer:one", &paragraphs).unwrap();
        assert!(
            evaluator
                .results
                .iter()
                .all(|result| matches!(result.outcome, FieldOutcome::KeepStored { .. }))
        );

        let mut heading_restart = Document::new();
        for (style_id, instruction) in [
            (Some("Heading1"), None),
            (None, Some(r"SEQ Figure \s 1")),
            (None, Some("SEQ Figure")),
            (Some("Heading1"), None),
            (None, Some(r"SEQ Figure \s 1")),
        ] {
            let mut paragraph = CT_P::new();
            paragraph.properties = style_id.map(|style_id| CT_PPr {
                style_id: Some(style_id.to_owned()),
                outline_lvl: Some(0),
                ..Default::default()
            });
            if let Some(instruction) = instruction {
                paragraph.runs.push(CT_R {
                    properties: None,
                    content: vec![RunContent::Field(Field::new(instruction, "stored"))],
                    extra_xml: Vec::new(),
                    extra_xml_positions: Vec::new(),
                    alt_drawings: Vec::new(),
                });
            }
            heading_restart
                .document
                .body
                .content
                .push(BodyContent::Paragraph(Box::new(paragraph)));
        }
        assert_eq!(
            heading_restart
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("2".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
            ]
        );
    }

    #[test]
    fn missing_context_and_unsupported_fields_keep_their_cached_display() {
        let document = document_with_fields(&[("DATE", "stored"), ("UNKNOWN", "stored")]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 2, "fallback outcomes must still be reported");
    }

    #[test]
    fn document_properties_variables_and_author_use_package_values() {
        let document = document_with_fields(&[("AUTHOR", "stored")]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 1, "package-backed fields must be reported");
    }

    #[test]
    fn styleref_searches_the_approved_direction_and_scope() {
        let mut document = Document::new();
        document
            .set_style(style::StyleBuilder::paragraph("Heading1", "Heading 1"))
            .unwrap();
        let definition = document
            .add_numbering_definition(&[crate::ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        let mut source = CT_P::new();
        source.properties = Some(CT_PPr {
            style_id: Some("Heading1".to_owned()),
            num_id: Some(instance),
            num_ilvl: Some(0),
            ..Default::default()
        });
        source.runs.push(CT_R::new("numbered heading"));
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(source)));
        let mut field = CT_P::new();
        field.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r#"STYLEREF "Heading 1" \n"#,
                "stored",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(field)));
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 1, "style fields must be reported");
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("1".to_owned()));

        let BodyContent::Paragraph(source) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        source.properties.as_mut().unwrap().num_id = Some(0);
        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("0".to_owned())
        );
    }

    #[test]
    fn ref_numbering_switches_use_the_authoritative_layout_counter() {
        let mut document = Document::new();
        let definition = document
            .add_numbering_definition(&[crate::ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        assert!(
            document
                .add_paragraph("numbered target")
                .set_numbering(instance, 0)
        );
        document
            .add_bookmark(
                "target",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        let mut paragraph = CT_P::new();
        for instruction in [
            r"REF target \n",
            r"REF target \r",
            r"REF target \w \t",
            r"REF target \n \p",
        ] {
            paragraph.runs.push(CT_R {
                properties: None,
                content: vec![RunContent::Field(Field::new(instruction, "stored"))],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            });
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(paragraph)));

        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("1 above".to_owned()),
            ]
        );
    }

    #[test]
    fn fields_without_numbering_switches_do_not_build_numbering_layout() {
        let document = Document::new();
        let context = FieldEvaluationContext::default();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new("AUTHOR", "stored"))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        let paragraphs = [&paragraph];
        let mut evaluator = Evaluator::new(&document, &context);

        evaluator.evaluate_story("main", &paragraphs).unwrap();

        assert!(evaluator.numbering_layout.is_none());

        let mut nested = Field::new("TOC", "stored");
        nested
            .instruction
            .arguments
            .push(FieldArgument::Nested(Box::new(Field::new(
                r"REF target \n",
                "stored",
            ))));
        evaluator
            .ensure_numbering_layout_for_field(&nested)
            .unwrap();
        assert!(evaluator.numbering_layout.is_some());
    }

    #[test]
    fn numbered_ref_uses_flattened_bookmark_paths_inside_and_after_a_table() {
        let mut document = Document::new();
        let definition = document
            .add_numbering_definition(&[crate::ListLevel::decimal()])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();

        let table_xml = format!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{instance}"/></w:numPr></w:pPr><w:bookmarkStart w:id="7" w:name="target_in_table"/><w:r><w:t>table item</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>"#
        );
        let mut reader = quick_xml::Reader::from_str(&table_xml);
        reader.config_mut().trim_text(true);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element)) if matches_local_name(element.name().as_ref(), b"p") => {
                    break;
                }
                Ok(Event::Eof) => panic!("table bookmark paragraph was missing"),
                Ok(_) => {}
                Err(error) => panic!("table bookmark paragraph failed to parse: {error}"),
            }
            buffer.clear();
        }
        let table_paragraph = CT_P::from_xml(&mut reader).unwrap();
        let mut cell = CT_Tc::new();
        *cell.paragraphs_mut()[0] = table_paragraph;
        let mut row = CT_Row::new();
        row.cells.push(cell);
        let mut table = CT_Tbl::new();
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(Box::new(table)));
        assert!(
            document
                .add_paragraph("numbered target")
                .set_numbering(instance, 0)
        );
        document
            .add_bookmark(
                "target_after_table",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 1,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 1,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        let mut field_paragraph = CT_P::new();
        field_paragraph.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r"REF target_in_table \n",
                "stored",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        field_paragraph.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r"REF target_after_table \n",
                "stored",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(field_paragraph)));

        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .into_iter()
                .map(|field| field.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("2".to_owned()),
            ]
        );
    }

    #[test]
    fn ref_switches_match_the_word_16_112_numbering_record() {
        let mut document = Document::new();
        let definition = document
            .add_numbering_definition(&[
                crate::ListLevel::decimal().level_text("Section %1."),
                crate::ListLevel::decimal().level_text("%1.%2."),
                crate::ListLevel::decimal().level_text("Section %1.%2.%3."),
            ])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        for (text, level) in [("Body", 0), ("Table", 1), ("Deep", 2)] {
            assert!(document.add_paragraph(text).set_numbering(instance, level));
        }
        document
            .add_bookmark(
                "deep",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 2,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 2,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        let mut paragraph = CT_P::new();
        for instruction in [
            r"REF deep \n",
            r"REF deep \n \t",
            r"REF deep \r",
            r"REF deep \w",
            r"REF deep \w \t",
            r"REF deep \n \p",
            r"REF deep \p",
        ] {
            paragraph.runs.push(CT_R {
                properties: None,
                content: vec![RunContent::Field(Field::new(instruction, "stored"))],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            });
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(paragraph)));

        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("Section 1.1.1".to_owned()),
                FieldOutcome::Resolved("1.1.1".to_owned()),
                FieldOutcome::Resolved("Section 1.1.1".to_owned()),
                FieldOutcome::Resolved("Section 1.1.1".to_owned()),
                FieldOutcome::Resolved("1.1.1".to_owned()),
                FieldOutcome::Resolved("Section 1.1.1 above".to_owned()),
                FieldOutcome::Resolved("above".to_owned()),
            ]
        );
    }

    #[test]
    fn ref_relative_number_keeps_level_text_and_position_fallback() {
        let mut document = Document::new();
        let definition = document
            .add_numbering_definition(&[
                crate::ListLevel::decimal().level_text("Section %1."),
                crate::ListLevel::decimal().level_text("%1.%2."),
                crate::ListLevel::decimal().level_text("Section %1.%2.%3."),
            ])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        for (text, level) in [("Root", 0), ("Branch", 1), ("Target", 2)] {
            assert!(document.add_paragraph(text).set_numbering(instance, level));
        }
        document
            .add_bookmark(
                "numbered_target",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 2,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 2,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        let mut relative = CT_P::new();
        relative.properties = Some(CT_PPr {
            num_id: Some(instance),
            num_ilvl: Some(2),
            ..CT_PPr::default()
        });
        relative.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r"REF numbered_target \r",
                "stored",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(relative)));
        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("Section 1.1.1".to_owned())
        );

        let mut plain = Document::new();
        plain.add_paragraph("Plain target");
        plain
            .add_bookmark(
                "plain_target",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
            )
            .unwrap();
        let mut fallback = CT_P::new();
        fallback.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r"REF plain_target \n \p",
                "stored",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        plain
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(fallback)));
        assert_eq!(
            plain
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("Plain target above".to_owned())
        );
    }

    #[test]
    fn ref_n_r_and_w_apply_level_relative_and_full_context_rules() {
        let mut document = Document::new();
        let definition = document
            .add_numbering_definition(&[
                crate::ListLevel::decimal().level_text("%1."),
                crate::ListLevel::decimal().level_text("%1.%2."),
                crate::ListLevel::decimal().level_text("%3."),
            ])
            .unwrap();
        let instance = document.add_numbering_instance(definition, &[]).unwrap();
        for index in 1..=4 {
            assert!(
                document
                    .add_paragraph(&format!("Root {index}"))
                    .set_numbering(instance, 0)
            );
        }
        for index in 1..=3 {
            assert!(
                document
                    .add_paragraph(&format!("Branch {index}"))
                    .set_numbering(instance, 1)
            );
        }
        let mut field_paragraph = CT_P::new();
        field_paragraph.properties = Some(CT_PPr {
            num_id: Some(instance),
            num_ilvl: Some(2),
            ..CT_PPr::default()
        });
        for instruction in [
            r"REF later_target \n",
            r"REF later_target \r",
            r"REF later_target \w",
        ] {
            field_paragraph.runs.push(CT_R {
                properties: None,
                content: vec![RunContent::Field(Field::new(instruction, "stored"))],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            });
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(Box::new(field_paragraph)));
        for index in 4..=5 {
            assert!(
                document
                    .add_paragraph(&format!("Branch {index}"))
                    .set_numbering(instance, 1)
            );
        }
        assert!(
            document
                .add_paragraph("Target precursor")
                .set_numbering(instance, 2)
        );
        assert!(document.add_paragraph("Target").set_numbering(instance, 2));
        document
            .add_bookmark(
                "later_target",
                crate::comments::RunRange {
                    start: crate::comments::RunPosition {
                        body_index: 11,
                        run_index: 0,
                    },
                    end: crate::comments::RunPosition {
                        body_index: 11,
                        run_index: 1,
                    },
                },
            )
            .unwrap();

        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .into_iter()
                .map(|field| field.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("2".to_owned()),
                FieldOutcome::Resolved("4.5.2".to_owned()),
                FieldOutcome::Resolved("4.5.2".to_owned()),
            ]
        );
    }

    #[test]
    fn date_time_filename_mergefield_and_includetext_use_only_explicit_context() {
        let document = document_with_fields(&[("FILENAME", "stored")]);
        let context = FieldEvaluationContext {
            file_name: Some("report.docx".to_owned()),
            ..FieldEvaluationContext::default()
        };
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(
            results[0].outcome,
            FieldOutcome::Resolved("report.docx".to_owned())
        );
    }

    #[test]
    fn formatting_switches_match_the_pinned_word_matrix() {
        let document = document_with_fields(&[(r#"MERGEFIELD Name \* Upper"#, "stored")]);
        let mut context = FieldEvaluationContext::default();
        context
            .merge_fields
            .insert("Name".to_owned(), "field value".to_owned());
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(
            results[0].outcome,
            FieldOutcome::Resolved("FIELD VALUE".to_owned())
        );
        for (value, format, expected) in [
            ("FiELD", "Lower", "field"),
            ("field VALUE", "FirstCap", "Field VALUE"),
            ("field value", "Caps", "Field Value"),
            ("27", "Arabic", "27"),
            ("same", "MERGEFORMAT", "same"),
            ("same", "Charformat", "same"),
            ("27", "alphabetic", "aa"),
            ("27", "ALPHABETIC", "AA"),
            ("14", "roman", "xiv"),
            ("14", "ROMAN", "XIV"),
            ("22", "Ordinal", "22nd"),
        ] {
            let instruction = FieldInstruction::new(
                "SEQ",
                Vec::new(),
                vec![field_option_switch("*", Some(format.into()))],
            )
            .unwrap();
            assert_eq!(
                rdocx_layout::engine::format_numeric_field_general(&instruction, value).unwrap(),
                expected
            );
        }
        for (value, picture, expected) in [
            ("1234.5", "#,##0.00", "1,234.50"),
            ("-12", "$0.00;($0.00);\"zero\"", "($12.00)"),
            ("0", "$0.00;($0.00);\"zero\"", "zero"),
            ("0", "#", " "),
            ("15", "$###", "$ 15"),
            ("12.5", "$##0.00 'is sales tax'", "$ 12.50 is sales tax"),
            ("15", "#'x'##", " x15"),
            ("123456", "000'-'000", "123-456"),
        ] {
            assert_eq!(
                rdocx_layout::engine::format_numeric_field_picture(value, picture).unwrap(),
                expected
            );
        }
        let now = FieldDateTime {
            year: 2025,
            month: 12,
            day: 14,
            hour: 21,
            minute: 7,
            second: 5,
        };
        assert_eq!(
            format_date_time(now, "dddd, MMMM d, yyyy HH:mm:ss AM/PM").unwrap(),
            "Sunday, December 14, 2025 21:07:05 PM"
        );
        let property = Field::new(r#"DOCPROPERTY SavedAt \@ "MMMM d, yyyy""#, "stored");
        assert_eq!(
            apply_formats(&property.instruction, "2025-12-14T21:07:05Z", None).unwrap(),
            "December 14, 2025"
        );
        let merge = Field::new(r#"MERGEFIELD Date \@ "yyyy-MM-dd""#, "stored");
        assert_eq!(
            apply_formats(&merge.instruction, "2026-01-02", None).unwrap(),
            "2026-01-02"
        );
    }

    #[test]
    fn wildcard_if_matches_multiline_nested_ref_values() {
        let mut document =
            document_with_fields(&[(r#"IF left = "first*second" yes no"#, "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("REF Multi", "stored ref")));

        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            unreachable!()
        };
        let paragraphs = [paragraph.as_ref()];
        let context = FieldEvaluationContext::default();
        let mut evaluator = Evaluator::new(&document, &context);
        evaluator.bookmarks.insert(
            "Multi".to_owned(),
            BookmarkValue {
                text: "first\nsecond".to_owned(),
                range: None,
            },
        );
        evaluator.evaluate_story("main", &paragraphs).unwrap();
        assert_eq!(
            evaluator.results[0].outcome,
            FieldOutcome::Resolved("yes".to_owned())
        );
        assert_eq!(
            evaluator.results[1].outcome,
            FieldOutcome::Resolved("first\nsecond".to_owned())
        );
    }

    #[test]
    fn nested_merge_regions_resolve_lexically_before_named_sources() {
        let record = |value: &str| MailMergeRecord {
            values: BTreeMap::from([("Value".to_owned(), MailMergeValue::Text(value.to_owned()))]),
            ..Default::default()
        };
        let root = MailMergeRecord {
            regions: BTreeMap::from([("Items".to_owned(), vec![record("root")])]),
            ..Default::default()
        };
        let inner = MailMergeRecord {
            regions: BTreeMap::from([("Items".to_owned(), vec![record("inner")])]),
            ..Default::default()
        };
        let sibling = MailMergeRecord {
            regions: BTreeMap::from([("Items".to_owned(), vec![record("sibling")])]),
            ..Default::default()
        };
        let empty = MailMergeRecord {
            regions: BTreeMap::from([("Items".to_owned(), Vec::new())]),
            ..Default::default()
        };
        let no_local_region = MailMergeRecord::default();
        let data = MailMergeData {
            sources: BTreeMap::from([
                ("Items".to_owned(), vec![record("named items")]),
                ("Fallback".to_owned(), vec![record("named fallback")]),
            ]),
            ..Default::default()
        };
        let scopes = [&root, &inner];
        assert_eq!(
            resolve_region_records(&scopes, &data, "Items").unwrap()[0]
                .values
                .get("Value"),
            Some(&MailMergeValue::Text("inner".to_owned()))
        );
        let sibling_scopes = [&root, &sibling];
        assert_eq!(
            resolve_region_records(&sibling_scopes, &data, "Items").unwrap()[0]
                .values
                .get("Value"),
            Some(&MailMergeValue::Text("sibling".to_owned()))
        );
        let fallback_scopes = [&root, &no_local_region];
        assert_eq!(
            resolve_region_records(&fallback_scopes, &data, "Items").unwrap()[0]
                .values
                .get("Value"),
            Some(&MailMergeValue::Text("named items".to_owned()))
        );
        assert_eq!(
            resolve_region_records(&fallback_scopes, &data, "Fallback").unwrap()[0]
                .values
                .get("Value"),
            Some(&MailMergeValue::Text("named fallback".to_owned()))
        );
        assert!(
            resolve_region_records(&[&empty], &data, "Items")
                .unwrap()
                .is_empty()
        );
        assert!(resolve_region_records(&scopes, &data, "Missing").is_err());
        assert_eq!(one_based_record_number(0).unwrap(), 1);
        assert!(one_based_record_number(u32::MAX as usize).is_err());

        let marker = |name: &str| {
            let instruction = format!("MERGEFIELD {name}");
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R {
                properties: None,
                content: vec![RunContent::Field(Field::new(&instruction, "marker"))],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            });
            BodyContent::Paragraph(Box::new(paragraph))
        };
        let crossed = [
            marker("TableStart:Outer"),
            marker("TableStart:Inner"),
            marker("TableEnd:Outer"),
        ];
        assert!(find_body_region_end(&crossed, 1, "Outer").is_err());
        let missing = [marker("TableStart:Outer")];
        assert!(find_body_region_end(&missing, 1, "Outer").is_err());
    }

    #[test]
    fn package_story_forbidden_comment_bodies_fail_closed() {
        let xml = format!(r#"<w:hdr xmlns:w="{W_NS}"><!--producer--invalid--><w:p/></w:hdr>"#);
        assert!(validate_story_document_declarations_and_doctype(xml.as_bytes()).is_err());
    }

    #[test]
    fn package_story_processing_instruction_targets_require_xml_names() {
        for instruction in ["<?XML version=\"1.0\"?>", "<?1producer value?>"] {
            let xml = format!(r#"{instruction}<w:hdr xmlns:w="{W_NS}"><w:p/></w:hdr>"#);
            assert!(
                validate_story_document_declarations_and_doctype(xml.as_bytes()).is_err(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn package_story_xml_names_bindings_and_expanded_attributes_fail_closed() {
        for malformed in [
            r#"<1producer/>"#,
            r#"<producer:item/>"#,
            r#"<w:p 1producer="value"/>"#,
            r#"<w:p producer:value="opaque"/>"#,
            r#"<w:p xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidR="one" q:rsidR="two"/>"#,
        ] {
            let xml = format!(r#"<w:hdr xmlns:w="{W_NS}">{malformed}</w:hdr>"#);
            assert!(
                validate_story_document_declarations_and_doctype(xml.as_bytes()).is_err(),
                "{malformed}"
            );
        }
    }

    #[test]
    fn package_story_nested_references_and_literal_xml_characters_fail_closed() {
        for malformed in [
            "<x:raw>&undefined;</x:raw>".to_owned(),
            "<x:raw>&#xFFFE;</x:raw>".to_owned(),
            "<x:raw>producer\u{1}</x:raw>".to_owned(),
            "<x:raw><![CDATA[producer\u{FFFE}]]></x:raw>".to_owned(),
        ] {
            let xml = format!(
                r#"<w:hdr xmlns:w="{W_NS}" xmlns:x="urn:producer">{malformed}<w:p/></w:hdr>"#
            );
            assert!(
                validate_story_document_declarations_and_doctype(xml.as_bytes()).is_err(),
                "{malformed:?}"
            );
        }
    }

    #[test]
    fn package_story_shared_lexical_failures_keep_the_other_error_surface() {
        for (xml, expected) in [
            (
                &b"<root>\x01</root>"[..],
                "package story XML contains a forbidden literal XML 1.0 character",
            ),
            (
                &b"<producer:item/>"[..],
                "package story XML uses an unbound namespace prefix",
            ),
            (
                &b"<root xmlns:a=\"urn:same\" xmlns:b=\"urn:same\" a:id=\"1\" b:id=\"2\"/>"[..],
                "package story XML element has duplicate expanded-name attributes",
            ),
            (
                &b"<root>&undefined;</root>"[..],
                "undeclared package story XML entity reference &undefined;",
            ),
            (
                &b"<?XML value?><root/>"[..],
                "reserved package story XML processing instruction",
            ),
        ] {
            assert!(
                matches!(
                    validate_story_document_declarations_and_doctype(xml),
                    Err(Error::Other(message)) if message == expected
                ),
                "{xml:?}"
            );
        }
    }

    #[test]
    fn content_fragment_identity_preflight_rejects_ambiguous_bookmark_ownership() {
        for fragment in [
            r#"<x:raw><w:bookmarkStart w:id="4" w:name="one"/><w:bookmarkStart w:id="4" w:name="one"/><w:bookmarkEnd w:id="4"/></x:raw>"#,
            r#"<x:raw><w:bookmarkStart w:id="4" w:name="one"/><w:bookmarkEnd w:id="4"/><w:bookmarkEnd w:id="4"/></x:raw>"#,
            r#"<x:raw><w:bookmarkStart w:id="4" w:name="one"/><w:bookmarkStart w:id="4" w:name="two"/><w:bookmarkEnd w:id="4"/></x:raw>"#,
        ] {
            let wrapped = format!(
                r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body>{fragment}</w:body></w:document>"#
            );
            assert!(
                freshen_content_fragment_identities(&mut Document::new(), wrapped.as_bytes(),)
                    .is_err()
            );
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum DynamicOwnerPolicy {
    Toc,
    GeneratedTables,
    Bibliography,
}

fn generated_table_opcode(name: &str, policy: DynamicOwnerPolicy) -> bool {
    name == "TOC"
        || policy != DynamicOwnerPolicy::Toc && matches!(name, "INDEX" | "TOA")
        || policy == DynamicOwnerPolicy::Bibliography && matches!(name, "CITATION" | "BIBLIOGRAPHY")
}

fn generated_authority_categories(document: &Document) -> Result<Vec<u8>> {
    let snapshot = rdocx_layout::engine::evaluate_sequence_fields(&document.build_layout_input()?)
        .map_err(|error| Error::Other(format!("authority source evaluation failed: {error}")))?;
    let mut categories = BTreeSet::new();
    for index in 1u32.. {
        let Some(node) = rdocx_layout::SourceNodeId::new(index) else {
            break;
        };
        if snapshot.source_node(node).is_none() {
            break;
        }
        for index in 0u32.. {
            let source = oxml_layout::FieldSource { node, index };
            let Some((instruction, _, cached)) = snapshot.source_field_context(source) else {
                break;
            };
            if cached {
                continue;
            }
            let field = Field::new(instruction, "");
            if field.instruction.name == "TA"
                && let Some(category) =
                    switch_text(&field.instruction, "c").and_then(|value| value.parse::<u8>().ok())
            {
                generated_category(category)?;
                categories.insert(category);
            }
        }
    }
    Ok(categories.into_iter().collect())
}

#[derive(Clone)]
struct GeneratedStorySource {
    story: crate::StoryId,
    range: std::ops::Range<usize>,
    wrapper_len: usize,
    xml: Vec<u8>,
    body: CT_Body,
    spans: Vec<DynamicTocSpan>,
}

#[derive(Clone)]
enum GeneratedTableDefinition {
    Index {
        options: IndexOptions,
        sequence: Option<String>,
        cross_reference_separator: String,
    },
    Figures {
        options: TableOfFiguresOptions,
        styles: Vec<(String, u8)>,
        sequence: Option<String>,
        page_prefix_separator: String,
    },
    Authorities(TableOfAuthoritiesOptions),
}

#[derive(Clone)]
struct GeneratedTableOwner {
    story_index: usize,
    span_index: usize,
    definition: GeneratedTableDefinition,
    properties: CT_PPr,
}

#[derive(Clone)]
enum GeneratedSourceKind {
    Index(IndexEntry),
    Authority(AuthorityEntry),
    Caption {
        label: Option<String>,
        runs: Vec<CT_R>,
        tail: Vec<CT_R>,
        style: Option<String>,
    },
}

#[derive(Clone)]
struct GeneratedTableSource {
    location: crate::ContentLocation,
    run: usize,
    kind: GeneratedSourceKind,
    properties: Option<CT_RPr>,
    target: String,
    last_target: Option<String>,
    sequence_values: BTreeMap<String, i64>,
    end_sequence_values: BTreeMap<String, i64>,
}

impl Document {
    /// Rebuild supported INDEX, caption-selected TOC and TOA caches atomically.
    /// Every target is resolved from one immutable layout after provisional insertion.
    /// Producer namespace bindings remain intact. Target insertion that requires
    /// unsafe canonical serialization refuses without changing the document.
    pub fn rebuild_generated_tables(&mut self) -> Result<GeneratedTablesReport> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        let mut report = GeneratedTablesReport::default();
        let original_stories =
            generated_story_inventory(&candidate, DynamicOwnerPolicy::GeneratedTables)?;
        let original_sources = generated_table_sources(&candidate, &original_stories)?;
        generated_normalize_simple_tables(
            &mut candidate,
            SimpleGeneratedOwnerContext::Tables(&original_sources),
        )?;
        candidate.prepare_staged_package()?;
        let stories = generated_story_inventory(&candidate, DynamicOwnerPolicy::GeneratedTables)?;
        let mut owners = generated_table_owners(
            &candidate,
            &stories,
            &original_sources,
            &mut report.diagnostics,
        )?;
        if owners.is_empty() {
            return Ok(report);
        }
        let mut sources = generated_table_sources(&candidate, &stories)?;
        owners.retain(|owner| {
            let supported = generated_collation_supported(&candidate, &owner.definition, &sources);
            if !supported {
                report.diagnostics.push(
                    "generated table retains source keys outside captured en-US ASCII collation"
                        .into(),
                );
            }
            supported
        });
        if owners.is_empty() {
            return Ok(report);
        }
        sources.retain(|source| {
            owners
                .iter()
                .any(|owner| generated_source_matches(&candidate, &owner.definition, source))
        });
        generated_source_targets(&mut candidate, &mut sources, &mut report)?;
        generated_ensure_styles(&mut candidate, &sources);
        candidate.flush_to_package()?;
        let stories = generated_story_inventory(&candidate, DynamicOwnerPolicy::GeneratedTables)?;
        generated_publish_caches(
            &mut candidate,
            &stories,
            &owners,
            &sources,
            None,
            &mut report,
        )?;
        candidate = reopen_staged_document(candidate)?;
        // This is the only pagination call in the generated-table transaction.
        let snapshot = candidate.layout_deterministic()?;
        let stories = generated_story_inventory(&candidate, DynamicOwnerPolicy::GeneratedTables)?;
        generated_publish_caches(
            &mut candidate,
            &stories,
            &owners,
            &sources,
            Some(&snapshot),
            &mut report,
        )?;
        let completed = reopen_staged_document(candidate)?;
        completed.story_ranges()?;
        self.commit_staged_mutation(completed);
        Ok(report)
    }
}

fn generated_story_inventory(
    document: &Document,
    policy: DynamicOwnerPolicy,
) -> Result<Vec<GeneratedStorySource>> {
    let mut result = Vec::new();
    for owner in document.generated_table_story_sources()? {
        let crate::document::GeneratedStoryOwner {
            story,
            range,
            xml: original,
            namespaces: scope,
        } = owner;
        let mut word_prefix = "generatedWord".to_owned();
        while scope.contains_key(&word_prefix) {
            word_prefix.push('_');
        }
        let mut wrapper = format!("<{word_prefix}:document xmlns:{word_prefix}=\"{W_NS}\"");
        for (prefix, namespace) in scope {
            if prefix == "xml" {
                continue;
            }
            if prefix.is_empty() {
                wrapper.push_str(" xmlns=\"");
            } else {
                wrapper.push_str(&format!(" xmlns:{prefix}=\""));
            }
            wrapper.push_str(&xml_escape_attribute(&namespace));
            wrapper.push('"');
        }
        wrapper.push_str(&format!("><{word_prefix}:body>"));
        let wrapper_len = wrapper.len();
        let mut xml = wrapper.into_bytes();
        xml.extend_from_slice(&original[range.clone()]);
        xml.extend_from_slice(format!("</{word_prefix}:body></{word_prefix}:document>").as_bytes());
        let mut parsed = CT_Document::from_xml(&xml)?;
        prepare_physical_story_projection(&mut parsed.body, &mut [])?;
        let spans = scan_dynamic_table_spans(&xml, policy)?;
        result.push(GeneratedStorySource {
            story,
            range,
            wrapper_len,
            xml,
            body: parsed.body,
            spans,
        });
    }
    Ok(result)
}

fn generated_table_owners(
    document: &Document,
    stories: &[GeneratedStorySource],
    original_sources: &[GeneratedTableSource],
    diagnostics: &mut Vec<String>,
) -> Result<Vec<GeneratedTableOwner>> {
    let mut result = Vec::new();
    for (story_index, story) in stories.iter().enumerate() {
        let mut paragraphs = Vec::new();
        collect_body_paragraphs(&story.body, &mut paragraphs);
        for (span_index, span) in story.spans.iter().enumerate() {
            let field = parse_dynamic_toc_field(&story.xml, span)?;
            if span.simple_field.is_some() && field.locked() != Some(true) {
                if generated_table_definition(&field.effective_instruction()).is_ok_and(
                    |definition| {
                        definition.is_some_and(|definition| {
                            !generated_collation_supported(document, &definition, original_sources)
                        })
                    },
                ) {
                    diagnostics.push("generated table retains source keys outside captured en-US ASCII collation".into());
                } else {
                    diagnostics.push(format!(
                        "generated {} retains unsupported simple owner formatting",
                        field.instruction.name
                    ));
                }
                continue;
            }
            if field.locked() == Some(true) {
                diagnostics.push(format!(
                    "generated {} field retains its locked cache",
                    field.instruction.name
                ));
                continue;
            }
            match generated_table_definition(&field.effective_instruction()) {
                Ok(Some(mut definition)) => {
                    if let GeneratedTableDefinition::Index { options, .. } = &mut definition {
                        options.hyperlink = paragraphs
                            .get(span.begin_paragraph)
                            .and_then(|paragraph| paragraph.properties.as_ref())
                            .and_then(|properties| properties.rpr.as_ref())
                            .and_then(|properties| properties.style_id.as_deref())
                            == Some("Hyperlink")
                            || field
                                .cached_display_segments()
                                .iter()
                                .any(|(_, properties)| {
                                    properties.is_some_and(|properties| {
                                        properties.style_id.as_deref() == Some("Hyperlink")
                                    })
                                });
                    }
                    result.push(GeneratedTableOwner {
                        story_index,
                        span_index,
                        definition,
                        properties: paragraphs
                            .get(span.begin_paragraph)
                            .and_then(|paragraph| paragraph.properties.clone())
                            .unwrap_or_default(),
                    });
                }
                Ok(None) => {}
                Err(diagnostic) => diagnostics.push(diagnostic),
            }
        }
    }
    Ok(result)
}

fn generated_table_definition(
    instruction: &FieldInstruction,
) -> std::result::Result<Option<GeneratedTableDefinition>, String> {
    if !instruction.arguments.is_empty() || !instruction.quotes_are_balanced() {
        return Err(format!(
            "generated {} retains malformed instruction",
            instruction.name
        ));
    }
    let allowed: &[&str] = match instruction.name.as_str() {
        "INDEX" => &["z", "f", "h", "e", "l", "g", "k", "s", "r", "*"],
        "TOA" => &["c", "h", "p", "e", "l", "g", "*"],
        "TOC" => &["c", "a", "t", "h", "p", "s", "d", "*"],
        _ => return Ok(None),
    };
    if instruction.name == "TOC"
        && !has_switch(instruction, "c")
        && !has_switch(instruction, "a")
        && !has_switch(instruction, "t")
    {
        return Ok(None);
    }
    let mut seen = HashSet::new();
    for switch in &instruction.switches {
        if !allowed.contains(&switch.name.as_str()) || !seen.insert(&switch.name) {
            return Err(format!(
                "generated {} retains unsupported or duplicate switch \\{}",
                instruction.name, switch.name
            ));
        }
        let flag = match instruction.name.as_str() {
            "INDEX" => switch.name == "r",
            "TOA" => matches!(switch.name.as_str(), "h" | "p"),
            _ => switch.name == "h",
        };
        if flag != switch.argument.is_none()
            || switch
                .argument
                .as_ref()
                .is_some_and(|argument| !matches!(argument, FieldArgument::Text(_)))
        {
            return Err(format!(
                "generated {} retains malformed switch \\{}",
                instruction.name, switch.name
            ));
        }
        if switch.name == "*" && switch_text(instruction, "*") != Some("MERGEFORMAT") {
            return Err(format!(
                "generated {} retains unsupported formatting switch",
                instruction.name
            ));
        }
    }
    let separator =
        |name: &str, default: &str| switch_text(instruction, name).unwrap_or(default).to_owned();
    match instruction.name.as_str() {
        "INDEX" => {
            if switch_text(instruction, "z").is_some_and(|locale| locale != "1033") {
                return Err("INDEX retains unsupported collation locale".into());
            }
            let options = IndexOptions {
                identifier: switch_text(instruction, "f").map(str::to_owned),
                heading_separator: switch_text(instruction, "h").map(str::to_owned),
                entry_page_separator: separator("e", ", "),
                page_separator: separator("l", ", "),
                range_separator: separator("g", "–"),
                run_in: has_switch(instruction, "r"),
                ..Default::default()
            };
            Ok(Some(GeneratedTableDefinition::Index {
                options,
                sequence: switch_text(instruction, "s").map(str::to_owned),
                cross_reference_separator: separator("e", ". "),
            }))
        }
        "TOA" => {
            let category = switch_text(instruction, "c")
                .and_then(|value| value.parse::<u8>().ok())
                .filter(|value| (1..=16).contains(value))
                .ok_or("TOA retains invalid or omitted category")?;
            Ok(Some(GeneratedTableDefinition::Authorities(
                TableOfAuthoritiesOptions {
                    category: Some(category),
                    include_category_headings: has_switch(instruction, "h"),
                    use_passim: has_switch(instruction, "p"),
                    entry_page_separator: separator("e", "\t"),
                    page_separator: separator("l", ", "),
                    range_separator: separator("g", "–"),
                    ..Default::default()
                },
            )))
        }
        _ => {
            if has_switch(instruction, "c") && has_switch(instruction, "a") {
                return Err("TOC retains conflicting caption selectors".into());
            }
            let styles = switch_text(instruction, "t")
                .map(parse_custom_styles)
                .transpose()?
                .unwrap_or_default();
            Ok(Some(GeneratedTableDefinition::Figures {
                options: TableOfFiguresOptions {
                    label: switch_text(instruction, "c")
                        .or_else(|| switch_text(instruction, "a"))
                        .unwrap_or("")
                        .into(),
                    include_label_and_number: !has_switch(instruction, "a"),
                    hyperlink: has_switch(instruction, "h"),
                    entry_page_separator: separator("p", "\t"),
                    ..Default::default()
                },
                styles,
                sequence: switch_text(instruction, "s").map(str::to_owned),
                page_prefix_separator: separator("d", "-"),
            }))
        }
    }
}

fn generated_source_location(
    document: &Document,
    location: &crate::ContentLocation,
) -> Result<crate::ContentLocation> {
    let story = document
        .stories()?
        .into_iter()
        .find(|story| {
            story.kind() == location.story().kind()
                && story.part_name() == location.story().part_name()
                && story.owner_index() == location.story().owner_index()
        })
        .ok_or_else(|| Error::Other("generated source owner disappeared".into()))?;
    Ok(crate::ContentLocation::new(
        story,
        crate::StoryItemKind::Paragraph,
        location.index_path().to_vec(),
    ))
}

fn generated_source_targets(
    document: &mut Document,
    sources: &mut [GeneratedTableSource],
    report: &mut GeneratedTablesReport,
) -> Result<()> {
    let mut names = caption_bookmark_names(document)?;
    let original_ranges = document.story_ranges()?;
    let mut targets = HashMap::<(StoryKind, String, usize, Vec<usize>, usize), String>::new();
    for source in sources {
        let range = match &source.kind {
            GeneratedSourceKind::Index(entry) => entry.page_range_bookmark.as_deref(),
            GeneratedSourceKind::Authority(entry) => entry.page_range_bookmark.as_deref(),
            _ => None,
        };
        let cross_reference = matches!(&source.kind, GeneratedSourceKind::Index(entry) if entry.cross_reference.is_some());
        if cross_reference {
            continue;
        }
        if let Some(name) = range {
            let matching = original_ranges.iter().filter(|range| matches!(range.kind(), crate::StoryRangeKind::Bookmark { name: actual, .. } if actual == name)).collect::<Vec<_>>();
            let [range] = matching.as_slice() else {
                return Err(Error::Other(format!(
                    "generated range target {name} is missing or ambiguous"
                )));
            };
            source.target = name.into();
            let endpoint = &range.range().end;
            let key = (
                endpoint.location.story().kind(),
                endpoint.location.story().part_name().into(),
                endpoint.location.story().owner_index(),
                endpoint.location.index_path().to_vec(),
                endpoint.run_index,
            );
            if let Some(target) = targets.get(&key) {
                source.last_target = Some(target.clone());
                continue;
            }
            let existing =
                generated_existing_target(&original_ranges, &endpoint.location, endpoint.run_index);
            let target = if let Some(target) = existing {
                target
            } else {
                let target = generated_target_name(&mut names)?;
                let position = crate::StoryRunPosition {
                    location: generated_source_location(document, &endpoint.location)?,
                    run_index: endpoint.run_index,
                };
                document.add_story_bookmark(
                    &target,
                    crate::StoryRunRange {
                        start: position.clone(),
                        end: position,
                    },
                )?;
                report.bookmark_count += 1;
                target
            };
            targets.insert(key, target.clone());
            source.last_target = Some(target);
        } else {
            let key = (
                source.location.story().kind(),
                source.location.story().part_name().into(),
                source.location.story().owner_index(),
                source.location.index_path().to_vec(),
                source.run,
            );
            if let Some(target) = targets.get(&key) {
                source.target = target.clone();
                continue;
            }
            let target = generated_existing_target(&original_ranges, &source.location, source.run);
            let target = if let Some(target) = target {
                target
            } else {
                let target = generated_target_name(&mut names)?;
                let position = crate::StoryRunPosition {
                    location: generated_source_location(document, &source.location)?,
                    run_index: source.run,
                };
                document.add_story_bookmark(
                    &target,
                    crate::StoryRunRange {
                        start: position.clone(),
                        end: position,
                    },
                )?;
                report.bookmark_count += 1;
                target
            };
            targets.insert(key, target.clone());
            source.target = target;
        }
    }
    Ok(())
}

fn generated_existing_target(
    ranges: &[crate::StoryRangeRef],
    location: &crate::ContentLocation,
    run: usize,
) -> Option<String> {
    ranges.iter().find_map(|range| {
        let crate::StoryRangeKind::Bookmark { name, .. } = range.kind() else {
            return None;
        };
        let same_position = |position: &crate::StoryRunPosition| {
            position.location.story().kind() == location.story().kind()
                && position.location.story().part_name() == location.story().part_name()
                && position.location.story().owner_index() == location.story().owner_index()
                && position.location.index_path() == location.index_path()
                && position.run_index == run
        };
        (same_position(&range.range().start) && same_position(&range.range().end))
            .then_some(name.clone())
    })
}

fn generated_target_name(names: &mut HashSet<String>) -> Result<String> {
    for index in 0u32.. {
        let name = format!("GeneratedTable_{index}");
        if names.insert(name.clone()) {
            return Ok(name);
        }
    }
    Err(Error::Other(
        "generated table bookmark names exhausted".into(),
    ))
}

fn generated_ensure_styles(document: &mut Document, sources: &[GeneratedTableSource]) {
    let mut names = BTreeSet::new();
    for source in sources {
        match &source.kind {
            GeneratedSourceKind::Index(entry) => {
                for level in 1..=entry.levels.len() {
                    names.insert(format!("Index{level}"));
                }
                names.insert("IndexHeading".into());
            }
            GeneratedSourceKind::Authority(_) => {
                names.insert("TableofAuthorities".into());
                names.insert("TOAHeading".into());
            }
            GeneratedSourceKind::Caption { .. } => {
                names.insert("TableofFigures".into());
            }
        }
    }
    for name in names {
        if document.styles.get_by_id(&name).is_some() {
            continue;
        }
        let (mut style, _, _) = style::StyleBuilder::paragraph(&name, &name).build();
        if let Some(level) = name
            .strip_prefix("Index")
            .and_then(|level| level.parse::<i32>().ok())
        {
            style.ppr = Some(CT_PPr {
                ind_left: Some(rdocx_oxml::units::Twips(level.saturating_mul(220))),
                ind_hanging: Some(rdocx_oxml::units::Twips(220)),
                ..Default::default()
            });
        } else if name == "TableofAuthorities" {
            style.ppr = Some(CT_PPr {
                ind_left: Some(rdocx_oxml::units::Twips(220)),
                ind_hanging: Some(rdocx_oxml::units::Twips(220)),
                space_after: Some(rdocx_oxml::units::Twips(0)),
                ..Default::default()
            });
        }
        document.styles.styles.push(style);
    }
}

fn generated_publish_caches(
    document: &mut Document,
    stories: &[GeneratedStorySource],
    owners: &[GeneratedTableOwner],
    sources: &[GeneratedTableSource],
    snapshot: Option<&rdocx_layout::WordLayoutResult>,
    report: &mut GeneratedTablesReport,
) -> Result<()> {
    let mut edits = BTreeMap::<String, Vec<FieldSourceEdit>>::new();
    for owner in owners {
        let story = stories
            .get(owner.story_index)
            .ok_or_else(|| Error::Other("generated story ownership changed".into()))?;
        let span = story
            .spans
            .get(owner.span_index)
            .ok_or_else(|| Error::Other("generated field ownership changed".into()))?;
        let generated = generated_render_table(document, owner, sources, snapshot, report)?;
        let replacement = if span.begin_paragraph == span.end_paragraph {
            let mut replacement = Vec::new();
            append_toc_wrapper_closures(&mut replacement, &span.separator_wrapper_names);
            replacement.extend_from_slice(format!("</{}>", span.start_paragraph_name).as_bytes());
            replacement.extend_from_slice(&generated);
            replacement.extend_from_slice(
                &story.xml[span.end_paragraph_start..span.end_paragraph_content_start],
            );
            for &(start, end) in &span.end_wrapper_prefixes {
                replacement.extend_from_slice(&story.xml[start..end]);
            }
            replacement
        } else {
            dynamic_toc_replacement(&story.xml, span, &generated, &[])?
        };
        let offset = |value: usize| {
            value
                .checked_sub(story.wrapper_len)
                .and_then(|value| value.checked_add(story.range.start))
                .ok_or_else(|| {
                    Error::Other(
                        "generated field source offset is outside its physical owner".into(),
                    )
                })
        };
        edits
            .entry(story.story.part_name().into())
            .or_default()
            .push(FieldSourceEdit {
                start: offset(span.result_start)?,
                end: offset(span.result_end)?,
                replacement,
            });
    }
    for (part, mut edits) in edits {
        edits.sort_by_key(|edit| edit.start);
        if edits.windows(2).any(|pair| pair[0].end > pair[1].start) {
            return Err(Error::Other("generated field cache edits overlap".into()));
        }
        let mut xml = document
            .package
            .get_part(&part)
            .ok_or_else(|| Error::Other("generated story part disappeared".into()))?
            .to_vec();
        for edit in edits.into_iter().rev() {
            xml.splice(edit.start..edit.end, edit.replacement);
        }
        validate_strict_xml_1_0(&xml)
            .map_err(|error| Error::Other(format!("invalid generated story XML: {error:?}")))?;
        document.package.set_part(&part, xml);
    }
    Ok(())
}

fn generated_table_sources(
    document: &Document,
    stories: &[GeneratedStorySource],
) -> Result<Vec<GeneratedTableSource>> {
    let paragraphs = document.generated_table_paragraph_positions()?;
    let snapshot = rdocx_layout::engine::evaluate_sequence_fields(&document.build_layout_input()?)
        .map_err(|error| {
            Error::Other(format!(
                "generated sequence source evaluation failed: {error}"
            ))
        })?;
    let mut result = Vec::new();
    let mut sequence_values = BTreeMap::new();
    for story in stories {
        let mut typed = Vec::new();
        collect_body_paragraphs(&story.body, &mut typed);
        let locations = paragraphs
            .iter()
            .filter(|(location, range)| {
                location.story().part_name() == story.story.part_name()
                    && story.range.start <= range.start
                    && range.end <= story.range.end
                    && (location.story().kind() != StoryKind::TextBox
                        || story.story.kind() == StoryKind::TextBox)
            })
            .collect::<Vec<_>>();
        if locations.len() != typed.len() {
            return Err(Error::Other(format!(
                "generated physical paragraph inventory disagrees with accepted ownership: {:?} locations {} paragraphs {}",
                story.story.kind(),
                locations.len(),
                typed.len()
            )));
        }
        for (paragraph_index, (paragraph, (location, _))) in
            typed.iter().zip(&locations).enumerate()
        {
            let mut caption_runs = Vec::new();
            let mut caption_label = None;
            let mut caption_tail = Vec::new();
            let mut after_sequence = false;
            let mut caption_has_source = false;
            let mut field_index = 0u32;
            for (accepted_index, accepted) in accepted_toc_runs(paragraph).into_iter().enumerate() {
                let position = TocOwnedPosition {
                    paragraph: paragraph_index,
                    run: TocRunPosition {
                        run_boundary: accepted.run_boundary,
                        raw_order: accepted.raw_order,
                        nested_order: accepted.nested_order,
                    },
                };
                let owned = toc_source_position_is_owned(&story.spans, position);
                let mut display_run = accepted.run.clone();
                display_run.content.clear();
                let mut tail_run = display_run.clone();
                for content in &accepted.run.content {
                    let RunContent::Field(field) = content else {
                        if !owned {
                            display_run.content.push(content.clone());
                            if after_sequence {
                                tail_run.content.push(content.clone());
                            }
                        }
                        continue;
                    };
                    let instruction = field.effective_instruction();
                    let physical_field_index = field_index;
                    field_index = field_index
                        .checked_add(
                            1 + u32::try_from(field.all_nested_fields_in_source_order().len())
                                .map_err(|_| {
                                    Error::Other(
                                        "generated source field inventory is too large".into(),
                                    )
                                })?,
                        )
                        .ok_or_else(|| {
                            Error::Other("generated source field inventory overflow".into())
                        })?;
                    if owned {
                        continue;
                    }
                    let kind = match instruction.name.as_str() {
                        "XE" => Some(GeneratedSourceKind::Index(generated_parse_index_marker(
                            &instruction,
                        )?)),
                        "TA" => Some(GeneratedSourceKind::Authority(
                            generated_parse_authority_marker(&instruction)?,
                        )),
                        "SEQ" => {
                            let label = text_argument(&instruction, 0).ok_or_else(|| {
                                Error::Other("caption sequence has no identifier".into())
                            })?;
                            let value = generated_sequence_value(
                                &snapshot,
                                document,
                                &story.story,
                                paragraph_index,
                                physical_field_index,
                                field,
                            )?;
                            if let Ok(number) = value.parse::<i64>() {
                                sequence_values.insert(label.to_ascii_lowercase(), number);
                            }
                            caption_label = Some(label.into());
                            after_sequence = true;
                            display_run
                                .content
                                .push(RunContent::Text(CT_Text::new(&value)));
                            None
                        }
                        _ => {
                            if !matches!(instruction.name.as_str(), "INDEX" | "TOA" | "TOC") {
                                display_run.content.push(content.clone());
                                if after_sequence {
                                    tail_run.content.push(content.clone());
                                }
                            }
                            None
                        }
                    };
                    if let Some(kind) = kind {
                        result.push(GeneratedTableSource {
                            location: (*location).clone(),
                            run: accepted_index,
                            kind,
                            properties: accepted.run.properties.clone(),
                            target: String::new(),
                            last_target: None,
                            sequence_values: sequence_values.clone(),
                            end_sequence_values: sequence_values.clone(),
                        });
                    }
                }
                if !owned {
                    caption_has_source = true;
                }
                if !tail_run.content.is_empty() {
                    caption_tail.push(tail_run);
                }
                if !display_run.content.is_empty() {
                    caption_runs.push(display_run);
                }
            }
            if caption_has_source
                && (caption_label.is_some()
                    || paragraph
                        .properties
                        .as_ref()
                        .and_then(|properties| properties.style_id.as_ref())
                        .is_some())
            {
                if let Some(first) = caption_tail.first_mut() {
                    for content in &mut first.content {
                        if let RunContent::Text(text) = content {
                            text.text = text
                                .text
                                .trim_start_matches([' ', ':', '.', '\t'])
                                .to_owned();
                            break;
                        }
                    }
                }
                result.push(GeneratedTableSource {
                    location: (*location).clone(),
                    run: 0,
                    kind: GeneratedSourceKind::Caption {
                        label: caption_label,
                        runs: caption_runs,
                        tail: caption_tail,
                        style: paragraph
                            .properties
                            .as_ref()
                            .and_then(|properties| properties.style_id.clone()),
                    },
                    properties: None,
                    target: String::new(),
                    last_target: None,
                    sequence_values: sequence_values.clone(),
                    end_sequence_values: sequence_values.clone(),
                });
            }
        }
    }
    let ranges = document.story_ranges()?;
    for source in &mut result {
        let name = match &source.kind {
            GeneratedSourceKind::Index(entry) => entry.page_range_bookmark.as_deref(),
            GeneratedSourceKind::Authority(entry) => entry.page_range_bookmark.as_deref(),
            _ => None,
        };
        let Some(name) = name else {
            continue;
        };
        let matching = ranges.iter().filter(|range| matches!(range.kind(), crate::StoryRangeKind::Bookmark { name: actual, .. } if actual == name)).collect::<Vec<_>>();
        let [range] = matching.as_slice() else {
            return Err(Error::Other(format!(
                "generated range target {name} is missing or ambiguous"
            )));
        };
        source.sequence_values =
            generated_sequence_context_at(&snapshot, document, &range.range().start)?;
        source.end_sequence_values =
            generated_sequence_context_at(&snapshot, document, &range.range().end)?;
    }
    Ok(result)
}

fn generated_sequence_context_at(
    snapshot: &rdocx_layout::WordSequenceSnapshot,
    document: &Document,
    position: &crate::StoryRunPosition,
) -> Result<BTreeMap<String, i64>> {
    let mut values = BTreeMap::new();
    if position.location.story().kind() != StoryKind::Body {
        return Ok(values);
    }
    let path = generated_main_source_path(document, &position.location)?;
    for event in snapshot.main_events() {
        let Some(source) = snapshot.source_node(event.source.node) else {
            return Err(Error::Other(
                "sequence event lost its physical source".into(),
            ));
        };
        if (source.children.as_slice(), event.accepted_run) < (path.as_slice(), position.run_index)
        {
            values.insert(event.identifier.to_ascii_lowercase(), event.value);
        }
    }
    Ok(values)
}

fn generated_main_source_path(
    document: &Document,
    location: &crate::ContentLocation,
) -> Result<Vec<usize>> {
    if location.story().kind() == StoryKind::Body && location.index_path().len() == 1 {
        Ok(document
            .story_items(location.story())?
            .into_iter()
            .find(|item| item.location() == location)
            .ok_or_else(|| {
                Error::Other("generated source paragraph is absent from its checked owner".into())
            })?
            .direct_body_index()?
            .map(|index| vec![index])
            .unwrap_or_default())
    } else {
        Ok(location.index_path().to_vec())
    }
}

fn generated_sequence_value(
    snapshot: &rdocx_layout::WordSequenceSnapshot,
    document: &Document,
    story: &crate::StoryId,
    paragraph_index: usize,
    field_index: u32,
    field: &Field,
) -> Result<String> {
    let input = document.build_layout_input()?;
    let mut nodes = Vec::new();
    for index in 1u32.. {
        let Some(node) = rdocx_layout::SourceNodeId::new(index) else {
            break;
        };
        let Some(path) = snapshot.source_node(node) else {
            break;
        };
        let same_story = match &path.story {
            rdocx_layout::WordStory::Document => {
                story.kind() == StoryKind::Body && story.part_name() == document.doc_part_name
            }
            rdocx_layout::WordStory::Header { .. } => {
                story.kind() == StoryKind::Header
                    && input.story_part_names.get(&path.story).map(String::as_str)
                        == Some(story.part_name())
            }
            rdocx_layout::WordStory::Footer { .. } => {
                story.kind() == StoryKind::Footer
                    && input.story_part_names.get(&path.story).map(String::as_str)
                        == Some(story.part_name())
            }
            rdocx_layout::WordStory::Footnote { id } => {
                document.footnote_story(*id)?.is_some_and(|owner| {
                    owner.kind() == story.kind()
                        && owner.part_name() == story.part_name()
                        && owner.owner_index() == story.owner_index()
                })
            }
            rdocx_layout::WordStory::Endnote { id } => {
                document.endnote_story(*id)?.is_some_and(|owner| {
                    owner.kind() == story.kind()
                        && owner.part_name() == story.part_name()
                        && owner.owner_index() == story.owner_index()
                })
            }
            rdocx_layout::WordStory::TextBox { part_name, .. } => {
                story.kind() == StoryKind::TextBox
                    && part_name == story.part_name()
                    && snapshot.text_box_owner_index(&path.story) == Some(story.owner_index())
            }
        };
        if same_story {
            nodes.push(node);
        }
    }
    let node = *nodes
        .get(paragraph_index)
        .ok_or_else(|| Error::Other("caption source has no qualified physical paragraph".into()))?;
    let source = oxml_layout::FieldSource {
        node,
        index: field_index,
    };
    let instruction = field.effective_instruction_text();
    if !snapshot
        .source_field_context(source)
        .is_some_and(|(actual, _, cached)| !cached && actual.trim() == instruction.trim())
    {
        return Err(Error::Other(
            "caption source disagrees with its physical field identity".into(),
        ));
    }
    snapshot
        .field_value(source)
        .ok_or_else(|| Error::Other("caption source has no physical sequence decision".into()))?
        .map(str::to_owned)
        .map_err(|diagnostic| Error::Other(diagnostic.into()))
}

fn generated_parse_index_marker(instruction: &FieldInstruction) -> Result<IndexEntry> {
    generated_marker_switches(instruction, &["f", "r", "t", "b", "i"])?;
    if instruction.arguments.len() != 1 {
        return Err(Error::Other("XE requires one hierarchy operand".into()));
    }
    let levels = text_argument(instruction, 0)
        .ok_or_else(|| Error::Other("XE hierarchy must be text".into()))?
        .split(':')
        .map(str::to_owned)
        .collect::<Vec<_>>();
    if levels.iter().any(String::is_empty) {
        return Err(Error::Other(
            "XE contains an empty hierarchy component".into(),
        ));
    }
    let entry = IndexEntry {
        levels,
        identifier: switch_text(instruction, "f").map(str::to_owned),
        page_range_bookmark: switch_text(instruction, "r").map(str::to_owned),
        cross_reference: switch_text(instruction, "t").map(str::to_owned),
        bold_page_numbers: has_switch(instruction, "b"),
        italic_page_numbers: has_switch(instruction, "i"),
    };
    if entry.page_range_bookmark.is_some() && entry.cross_reference.is_some() {
        return Err(Error::Other(
            "XE combines a range and cross-reference".into(),
        ));
    }
    Ok(entry)
}

fn generated_parse_authority_marker(instruction: &FieldInstruction) -> Result<AuthorityEntry> {
    generated_marker_switches(instruction, &["l", "s", "c", "r", "b", "i"])?;
    if !instruction.arguments.is_empty() {
        return Err(Error::Other("TA does not take positional operands".into()));
    }
    let long_citation = switch_text(instruction, "l")
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Other("TA requires a long citation".into()))?
        .into();
    let short_citation = switch_text(instruction, "s")
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Other("TA requires a short grouping citation".into()))?
        .into();
    let category = switch_text(instruction, "c")
        .and_then(|value| value.parse::<u8>().ok())
        .ok_or_else(|| Error::Other("TA requires a numbered category".into()))?;
    generated_category(category)?;
    Ok(AuthorityEntry {
        long_citation,
        short_citation,
        category,
        page_range_bookmark: switch_text(instruction, "r").map(str::to_owned),
        bold_page_numbers: has_switch(instruction, "b"),
        italic_page_numbers: has_switch(instruction, "i"),
    })
}

fn generated_marker_switches(instruction: &FieldInstruction, allowed: &[&str]) -> Result<()> {
    let mut seen = HashSet::new();
    for switch in &instruction.switches {
        if !allowed.contains(&switch.name.as_str()) || !seen.insert(&switch.name) {
            return Err(Error::Other(format!(
                "{} source contains unsupported or duplicate switch \\{}",
                instruction.name, switch.name
            )));
        }
        let flag = matches!(switch.name.as_str(), "b" | "i");
        if flag != switch.argument.is_none()
            || switch
                .argument
                .as_ref()
                .is_some_and(|argument| !matches!(argument, FieldArgument::Text(_)))
        {
            return Err(Error::Other(format!(
                "{} source contains malformed switch \\{}",
                instruction.name, switch.name
            )));
        }
    }
    Ok(())
}

fn generated_collation_supported(
    document: &Document,
    definition: &GeneratedTableDefinition,
    sources: &[GeneratedTableSource],
) -> bool {
    sources
        .iter()
        .filter(|source| generated_source_matches(document, definition, source))
        .all(|source| match &source.kind {
            GeneratedSourceKind::Index(entry) => entry.levels.iter().all(|level| level.is_ascii()),
            GeneratedSourceKind::Authority(entry) => entry.long_citation.is_ascii(),
            GeneratedSourceKind::Caption { .. } => true,
        })
}

fn generated_source_matches(
    document: &Document,
    definition: &GeneratedTableDefinition,
    source: &GeneratedTableSource,
) -> bool {
    match (definition, &source.kind) {
        (GeneratedTableDefinition::Index { options, .. }, GeneratedSourceKind::Index(entry)) => {
            entry.identifier == options.identifier
        }
        (GeneratedTableDefinition::Authorities(options), GeneratedSourceKind::Authority(entry)) => {
            Some(entry.category) == options.category
        }
        (
            GeneratedTableDefinition::Figures {
                options, styles, ..
            },
            GeneratedSourceKind::Caption { label, style, .. },
        ) => {
            if styles.is_empty() {
                label.as_deref() == Some(options.label.as_str())
            } else {
                style.as_deref().is_some_and(|id| {
                    styles.iter().any(|(name, _)| {
                        name.eq_ignore_ascii_case(id)
                            || document
                                .styles
                                .get_by_id(id)
                                .and_then(|style| style.name.as_deref())
                                .is_some_and(|actual| actual.eq_ignore_ascii_case(name))
                    })
                })
            }
        }
        _ => false,
    }
}

fn generated_render_table(
    document: &Document,
    owner: &GeneratedTableOwner,
    sources: &[GeneratedTableSource],
    snapshot: Option<&rdocx_layout::WordLayoutResult>,
    report: &mut GeneratedTablesReport,
) -> Result<Vec<u8>> {
    let selected = sources
        .iter()
        .filter(|source| generated_source_matches(document, &owner.definition, source))
        .collect::<Vec<_>>();
    match &owner.definition {
        GeneratedTableDefinition::Index {
            options,
            sequence,
            cross_reference_separator,
        } => generated_render_index(
            owner,
            &selected,
            options,
            sequence.as_deref(),
            cross_reference_separator,
            snapshot,
            report,
        ),
        GeneratedTableDefinition::Authorities(options) => {
            generated_render_authorities(owner, &selected, options, snapshot, report)
        }
        GeneratedTableDefinition::Figures {
            options,
            sequence,
            page_prefix_separator,
            ..
        } => {
            let mut output = Vec::new();
            for source in &selected {
                let GeneratedSourceKind::Caption { runs, tail, .. } = &source.kind else {
                    continue;
                };
                let mut paragraph = generated_entry_paragraph(owner, "TableofFigures", 1);
                paragraph.runs = if options.include_label_and_number {
                    runs.clone()
                } else {
                    tail.clone()
                };
                let label_end = paragraph.runs.len();
                generated_append_separator(&mut paragraph, &options.entry_page_separator);
                paragraph.runs.push(generated_page_run(
                    source,
                    None,
                    snapshot,
                    sequence.as_deref(),
                    page_prefix_separator,
                    false,
                    false,
                )?);
                if options.hyperlink {
                    let end = label_end;
                    paragraph.hyperlinks.push(rdocx_oxml::text::HyperlinkSpan {
                        rel_id: None,
                        anchor: Some(source.target.clone()),
                        tooltip: None,
                        doc_location: None,
                        run_start: 0,
                        run_end: end,
                        extra_attributes: Vec::new(),
                        extra_xml: Vec::new(),
                        preserved_raw_before: None,
                    });
                }
                generated_write_paragraph(&paragraph, &mut output)?;
            }
            if selected.is_empty() {
                let mut paragraph = generated_entry_paragraph(owner, "TableofFigures", 1);
                paragraph.add_run("No table of figures entries found.");
                generated_write_paragraph(&paragraph, &mut output)?;
            }
            if snapshot.is_some() {
                report.figure_entries += selected.len();
            }
            Ok(output)
        }
    }
}

fn generated_entry_paragraph(owner: &GeneratedTableOwner, style: &str, _level: usize) -> CT_P {
    let mut paragraph = CT_P::new();
    let mut properties = owner.properties.clone();
    properties.style_id = Some(style.into());
    properties.num_id = None;
    properties.num_ilvl = None;
    properties.outline_lvl = None;
    if properties.tabs.is_none() {
        properties.tabs = Some(rdocx_oxml::borders::CT_Tabs {
            tabs: vec![rdocx_oxml::borders::CT_TabStop {
                val: ST_TabJc::Right,
                pos: rdocx_oxml::units::Twips(9350),
                leader: Some(rdocx_oxml::shared::ST_TabLeader::Dot),
                source_occurrence: None,
            }],
        });
    }
    paragraph.properties = Some(properties);
    paragraph
}

fn generated_append_separator(paragraph: &mut CT_P, separator: &str) {
    let mut run = CT_R::new("");
    let segments = separator.split('\t').collect::<Vec<_>>();
    for (index, segment) in segments.iter().enumerate() {
        if index != 0 {
            run.content.push(RunContent::Tab);
        }
        if !segment.is_empty() {
            run.content.push(RunContent::Text(CT_Text::new(segment)));
        }
    }
    paragraph.runs.push(run);
}

fn generated_write_paragraph(paragraph: &CT_P, output: &mut Vec<u8>) -> Result<()> {
    let mut writer = quick_xml::Writer::new(Vec::new());
    paragraph.to_xml(&mut writer)?;
    let xml = writer.into_inner();
    output.extend_from_slice(&xml_fragment_with_namespaces(
        &xml,
        &BTreeMap::from([("w".into(), W_NS.into())]),
        "generated entry paragraph",
    )?);
    Ok(())
}

fn generated_page_run(
    source: &GeneratedTableSource,
    last: Option<&str>,
    snapshot: Option<&rdocx_layout::WordLayoutResult>,
    sequence: Option<&str>,
    prefix_separator: &str,
    bold: bool,
    italic: bool,
) -> Result<CT_R> {
    let target = last.unwrap_or(&source.target);
    let mut properties = source.properties.clone().unwrap_or_default();
    properties.vanish = Some(false);
    if bold {
        properties.bold = Some(true);
        properties.bold_cs = Some(true);
    }
    if italic {
        properties.italic = Some(true);
        properties.italic_cs = Some(true);
    }
    let mut run = CT_R::new("");
    run.properties = Some(properties.clone());
    let value = if let Some(snapshot) = snapshot {
        let page = snapshot.bookmark_page_section(target).ok_or_else(|| {
            Error::Other(format!(
                "generated page target {target} is missing, ambiguous or unplaced"
            ))
        })?;
        let prefix = sequence
            .map(|identifier| {
                (if last.is_some() {
                    &source.end_sequence_values
                } else {
                    &source.sequence_values
                })
                .get(&identifier.to_ascii_lowercase())
                .copied()
                .ok_or_else(|| {
                    Error::Other(format!(
                        "generated sequence prefix {identifier} has no source decision"
                    ))
                })
            })
            .transpose()?;
        prefix.map_or_else(
            || page.displayed_page.to_string(),
            |prefix| format!("{prefix}{prefix_separator}{}", page.displayed_page),
        )
    } else {
        "99".into()
    };
    let mut cached = CT_R::new(&value);
    cached.properties = Some(properties);
    let field = Field::from_instruction(
        FieldInstruction::new(
            "PAGEREF",
            vec![FieldArgument::Text(target.into())],
            Vec::new(),
        )?,
        rdocx_oxml::text::FieldForm::Simple,
        vec![cached],
    )?;
    run.content = vec![RunContent::Field(field)];
    Ok(run)
}

fn generated_collation(left: &[String], right: &[String]) -> std::cmp::Ordering {
    for (left, right) in left.iter().zip(right) {
        let folded = left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase());
        if folded != std::cmp::Ordering::Equal {
            return folded;
        }
        // The pinned en-US captures place the distinct lowercase entry before its uppercase counterpart.
        let case = right.cmp(left);
        if case != std::cmp::Ordering::Equal {
            return case;
        }
    }
    left.len().cmp(&right.len())
}

fn generated_append_pages(
    paragraph: &mut CT_P,
    occurrences: &[&GeneratedTableSource],
    page_separator: &str,
    range_separator: &str,
    sequence: Option<&str>,
    passim: bool,
    snapshot: Option<&rdocx_layout::WordLayoutResult>,
) -> Result<()> {
    let mut seen = BTreeSet::new();
    let mut selected = Vec::new();
    for source in occurrences {
        if let Some(snapshot) = snapshot {
            let first = snapshot
                .bookmark_page_section(&source.target)
                .ok_or_else(|| {
                    Error::Other(format!("generated target {} is unplaced", source.target))
                })?;
            let last = source
                .last_target
                .as_deref()
                .map(|target| {
                    snapshot.bookmark_page_section(target).ok_or_else(|| {
                        Error::Other(format!("generated range endpoint {target} is unplaced"))
                    })
                })
                .transpose()?;
            if last.is_some_and(|last| last.physical_page < first.physical_page) {
                return Err(Error::Other(
                    "generated range endpoint precedes its start".into(),
                ));
            }
            let key = (first.physical_page, last.map(|last| last.physical_page));
            if !seen.insert(key) {
                continue;
            }
        }
        selected.push(*source);
    }
    if passim
        && snapshot.is_some()
        && selected
            .iter()
            .filter(|source| source.last_target.is_none())
            .count()
            >= 5
    {
        paragraph.add_run("passim");
        return Ok(());
    }
    for (index, source) in selected.iter().enumerate() {
        if index != 0 {
            generated_append_separator(paragraph, page_separator);
        }
        let (bold, italic) = match &source.kind {
            GeneratedSourceKind::Index(entry) => {
                (entry.bold_page_numbers, entry.italic_page_numbers)
            }
            GeneratedSourceKind::Authority(entry) => {
                (entry.bold_page_numbers, entry.italic_page_numbers)
            }
            _ => (false, false),
        };
        paragraph.runs.push(generated_page_run(
            source, None, snapshot, sequence, "-", bold, italic,
        )?);
        if let Some(last) = &source.last_target {
            generated_append_separator(paragraph, range_separator);
            paragraph.runs.push(generated_page_run(
                source,
                Some(last),
                snapshot,
                sequence,
                "-",
                bold,
                italic,
            )?);
        }
    }
    Ok(())
}

fn generated_render_index(
    owner: &GeneratedTableOwner,
    selected: &[&GeneratedTableSource],
    options: &IndexOptions,
    sequence: Option<&str>,
    cross_reference_separator: &str,
    snapshot: Option<&rdocx_layout::WordLayoutResult>,
    report: &mut GeneratedTablesReport,
) -> Result<Vec<u8>> {
    let mut groups = BTreeMap::<Vec<String>, Vec<&GeneratedTableSource>>::new();
    for source in selected {
        let GeneratedSourceKind::Index(entry) = &source.kind else {
            continue;
        };
        for length in 1..=entry.levels.len() {
            groups.entry(entry.levels[..length].to_vec()).or_default();
        }
        groups.entry(entry.levels.clone()).or_default().push(source);
    }
    let mut groups = groups.into_iter().collect::<Vec<_>>();
    groups.sort_by(|left, right| generated_collation(&left.0, &right.0));
    let mut output = Vec::new();
    let mut previous_heading = None;
    let mut run_in = None::<CT_P>;
    let mut root = None::<String>;
    let mut previous_depth = 0usize;
    for (levels, occurrences) in &groups {
        let heading = levels
            .first()
            .and_then(|level| level.chars().next())
            .map(|letter| letter.to_ascii_uppercase());
        if options.heading_separator.is_some() && previous_heading != heading {
            if let Some(paragraph) = run_in.take() {
                generated_write_paragraph(&paragraph, &mut output)?;
            }
            let mut paragraph = generated_entry_paragraph(owner, "IndexHeading", 1);
            paragraph.add_run(&heading.map(|letter| letter.to_string()).unwrap_or_default());
            generated_write_paragraph(&paragraph, &mut output)?;
            previous_heading = heading;
        }
        let same_root = root.as_deref() == levels.first().map(String::as_str);
        let mut paragraph = if options.run_in && same_root {
            let mut paragraph = run_in
                .take()
                .ok_or_else(|| Error::Other("run-in index lost its parent".into()))?;
            generated_append_separator(
                &mut paragraph,
                if previous_depth == 1 && levels.len() == 2 {
                    ": "
                } else {
                    "; "
                },
            );
            paragraph
        } else {
            if let Some(paragraph) = run_in.take() {
                generated_write_paragraph(&paragraph, &mut output)?;
            }
            generated_entry_paragraph(
                owner,
                &format!("Index{}", if options.run_in { 1 } else { levels.len() }),
                if options.run_in { 1 } else { levels.len() },
            )
        };
        let mut label = CT_R::new(levels.last().map_or("", String::as_str));
        label.properties = occurrences
            .first()
            .and_then(|source| source.properties.clone());
        if let Some(properties) = &mut label.properties {
            properties.vanish = Some(false);
        }
        let label_start = paragraph.runs.len();
        if options.hyperlink {
            label
                .properties
                .get_or_insert_with(CT_RPr::default)
                .style_id = Some("Hyperlink".into());
        }
        paragraph.runs.push(label);
        if options.hyperlink
            && let Some(source) = occurrences.iter().find(|source| !source.target.is_empty())
        {
            paragraph.hyperlinks.push(rdocx_oxml::text::HyperlinkSpan {
                rel_id: None,
                anchor: Some(source.target.clone()),
                tooltip: None,
                doc_location: None,
                run_start: label_start,
                run_end: label_start + 1,
                extra_attributes: Vec::new(),
                extra_xml: Vec::new(),
                preserved_raw_before: None,
            });
        }
        let cross_references = occurrences
            .iter()
            .filter_map(|source| match &source.kind {
                GeneratedSourceKind::Index(entry) => entry.cross_reference.as_deref(),
                _ => None,
            })
            .collect::<BTreeSet<_>>();
        if !cross_references.is_empty() {
            generated_append_separator(&mut paragraph, cross_reference_separator);
            paragraph.add_run(
                &cross_references
                    .into_iter()
                    .collect::<Vec<_>>()
                    .join(&options.page_separator),
            );
        }
        let pages = occurrences.iter().copied().filter(|source| matches!(&source.kind, GeneratedSourceKind::Index(entry) if entry.cross_reference.is_none())).collect::<Vec<_>>();
        if !pages.is_empty() {
            generated_append_separator(&mut paragraph, &options.entry_page_separator);
            generated_append_pages(
                &mut paragraph,
                &pages,
                &options.page_separator,
                &options.range_separator,
                sequence,
                false,
                snapshot,
            )?;
        }
        root = levels.first().cloned();
        previous_depth = levels.len();
        if options.run_in {
            run_in = Some(paragraph);
        } else {
            generated_write_paragraph(&paragraph, &mut output)?;
        }
    }
    if let Some(paragraph) = run_in {
        generated_write_paragraph(&paragraph, &mut output)?;
    }
    if snapshot.is_some() {
        report.index_entries += groups
            .iter()
            .filter(|(_, occurrences)| !occurrences.is_empty())
            .count();
    }
    Ok(output)
}

fn generated_render_authorities(
    owner: &GeneratedTableOwner,
    selected: &[&GeneratedTableSource],
    options: &TableOfAuthoritiesOptions,
    snapshot: Option<&rdocx_layout::WordLayoutResult>,
    report: &mut GeneratedTablesReport,
) -> Result<Vec<u8>> {
    let mut groups = BTreeMap::<String, Vec<&GeneratedTableSource>>::new();
    for source in selected {
        if let GeneratedSourceKind::Authority(entry) = &source.kind {
            groups
                .entry(entry.short_citation.clone())
                .or_default()
                .push(source);
        }
    }
    let mut groups = groups.into_values().collect::<Vec<_>>();
    let long = |group: &[&GeneratedTableSource]| {
        group
            .first()
            .and_then(|source| match &source.kind {
                GeneratedSourceKind::Authority(entry) => Some(entry.long_citation.clone()),
                _ => None,
            })
            .unwrap_or_default()
    };
    groups.sort_by(|left, right| generated_collation(&[long(left)], &[long(right)]));
    let mut output = Vec::new();
    if options.include_category_headings {
        let mut paragraph = generated_entry_paragraph(owner, "TOAHeading", 1);
        paragraph.add_run(match options.category {
            Some(1) => "Cases",
            Some(2) => "Statutes",
            Some(3) => "Other Authorities",
            Some(4) => "Rules",
            Some(5) => "Treatises",
            Some(6) => "Regulations",
            Some(7) => "Constitutional Provisions",
            Some(8) => "8",
            Some(9) => "9",
            Some(10) => "10",
            Some(11) => "11",
            Some(12) => "12",
            Some(13) => "13",
            Some(14) => "14",
            Some(15) => "15",
            Some(16) => "16",
            _ => {
                return Err(Error::Other(
                    "authority heading has no numbered category".into(),
                ));
            }
        });
        generated_write_paragraph(&paragraph, &mut output)?;
    }
    for occurrences in &groups {
        let mut paragraph = generated_entry_paragraph(owner, "TableofAuthorities", 1);
        let mut label = CT_R::new(&long(occurrences));
        label.properties = occurrences
            .first()
            .and_then(|source| source.properties.clone());
        if let Some(properties) = &mut label.properties {
            properties.vanish = Some(false);
        }
        paragraph.runs.push(label);
        generated_append_separator(&mut paragraph, &options.entry_page_separator);
        generated_append_pages(
            &mut paragraph,
            occurrences,
            &options.page_separator,
            &options.range_separator,
            None,
            options.use_passim,
            snapshot,
        )?;
        generated_write_paragraph(&paragraph, &mut output)?;
    }
    if snapshot.is_some() {
        report.authority_entries += groups.len();
    }
    Ok(output)
}

fn generated_simple_span(
    xml: &[u8],
    closed: &DynamicXmlElement,
    elements: &[DynamicXmlElement],
    closing_start: usize,
    end: usize,
    policy: DynamicOwnerPolicy,
) -> Result<Option<DynamicTocSpan>> {
    if !closed.is_word
        || closed.local_name != b"fldSimple"
        || !closed
            .paragraph
            .is_some_and(|paragraph| accepted_simple_field_parent(elements, paragraph))
    {
        return Ok(None);
    }
    let Some(paragraph_index) = closed.paragraph else {
        return Ok(None);
    };
    let raw = xml_fragment_with_namespaces(
        &xml[closed.start..end],
        &closed.inherited_namespaces,
        "simple generated field",
    )?;
    let mut fragment = format!("<w:p xmlns:w=\"{W_NS}\">").into_bytes();
    fragment.extend_from_slice(&raw);
    fragment.extend_from_slice(b"</w:p>");
    let paragraph = CT_P::from_xml_fragment(&fragment)?;
    let field = accepted_toc_runs(&paragraph)
        .into_iter()
        .flat_map(|run| &run.run.content)
        .find_map(|content| match content {
            RunContent::Field(field) => Some(field.clone()),
            _ => None,
        });
    let field = if field.is_none() && policy == DynamicOwnerPolicy::Bibliography {
        // Producer attributes can keep a simple field outside the typed projection.
        // Its qualified instruction still owns source references and global selectors.
        let mut reader = NsReader::from_reader(raw.as_slice());
        let mut buffer = Vec::new();
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid bibliography simple owner: {error}")))?;
        let start = match event {
            Event::Start(start) | Event::Empty(start) => start,
            _ => return Ok(None),
        };
        let instruction = resolved_element_attribute(
            &start,
            reader.resolver(),
            b"instr",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value);
        instruction
            .map(|instruction| {
                let mut field = Field::new(&instruction, "");
                for (name, locked) in [(b"fldLock".as_slice(), true), (b"dirty".as_slice(), false)]
                {
                    if let Some((_, value)) = resolved_element_attribute(
                        &start,
                        reader.resolver(),
                        name,
                        AttributeNamespace::Word,
                    )? {
                        let value = match value.as_str() {
                            "1" | "true" | "on" => true,
                            "0" | "false" | "off" => false,
                            _ => {
                                return Err(Error::Other(
                                    "ambiguous bibliography simple control".into(),
                                ));
                            }
                        };
                        if locked {
                            field.set_locked(Some(value));
                        } else {
                            field.dirty = Some(value);
                        }
                    }
                }
                Ok(field)
            })
            .transpose()?
    } else {
        field
    };
    let Some(field) = field.filter(|field| generated_table_opcode(&field.instruction.name, policy))
    else {
        return Ok(None);
    };
    let parent = elements
        .iter()
        .rev()
        .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph_index))
        .ok_or_else(|| Error::Other("simple generated field has no typed paragraph".into()))?;
    let position = closed.run_position.ok_or_else(|| {
        Error::Other("simple generated field has no accepted run boundary".into())
    })?;
    Ok(Some(DynamicTocSpan {
        instruction: field.effective_instruction_text(),
        field_start: closed.start,
        field_end: end,
        begin_paragraph: paragraph_index,
        end_paragraph: paragraph_index,
        begin_run_start: closed.start,
        instruction_paragraph_start: parent.start,
        result_start: closed.start_tag_end,
        result_end: closing_start,
        result_start_position: position,
        result_end_position: position,
        end_run_end: end,
        start_paragraph_name: parent.qualified_name.clone(),
        start_paragraph_namespaces: parent.inherited_namespaces.clone(),
        separator_wrapper_names: Vec::new(),
        instruction_runs: Vec::new(),
        end_paragraph_start: parent.start,
        end_paragraph_content_start: parent.start_tag_end,
        end_wrapper_prefixes: Vec::new(),
        simple_field: Some(field),
    }))
}

fn generated_simple_cache_xml(raw: &[u8], bindings: &BTreeMap<String, String>) -> Result<Vec<u8>> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start = 0usize;
    let mut copied = 0usize;
    let mut result = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid simple generated cache: {error}")))?;
        let after = reader.buffer_position() as usize;
        let complete = match event {
            Event::Start(_) => {
                if depth == 0 {
                    start = before;
                }
                depth += 1;
                false
            }
            Event::Empty(_) if depth == 0 => {
                start = before;
                true
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Other("simple generated cache has unmatched end".into())
                })?;
                depth == 0
            }
            Event::Eof => {
                if depth != 0 {
                    return Err(Error::Other("simple generated cache is unbalanced".into()));
                }
                result.extend_from_slice(&raw[copied..]);
                return Ok(result);
            }
            _ => false,
        };
        if complete {
            result.extend_from_slice(&raw[copied..start]);
            result.extend_from_slice(&xml_fragment_with_namespaces(
                &raw[start..after],
                bindings,
                "simple generated cache subtree",
            )?);
            copied = after;
        }
        buffer.clear();
    }
}

#[derive(Clone, Copy)]
enum SimpleGeneratedOwnerContext<'a> {
    Tables(&'a [GeneratedTableSource]),
    Bibliography(&'a crate::bibliography::BibliographyUpdateState),
}

fn generated_normalize_simple_tables(
    document: &mut Document,
    context: SimpleGeneratedOwnerContext<'_>,
) -> Result<()> {
    let policy = match context {
        SimpleGeneratedOwnerContext::Tables(_) => DynamicOwnerPolicy::GeneratedTables,
        SimpleGeneratedOwnerContext::Bibliography(_) => DynamicOwnerPolicy::Bibliography,
    };
    let stories = generated_story_inventory(document, policy)?;
    let mut edits = BTreeMap::<String, Vec<FieldSourceEdit>>::new();
    for story in &stories {
        for span in &story.spans {
            let Some(field) = &span.simple_field else {
                continue;
            };
            if field.locked() == Some(true) {
                continue;
            }
            match context {
                SimpleGeneratedOwnerContext::Tables(sources) => {
                    let Ok(Some(definition)) =
                        generated_table_definition(&field.effective_instruction())
                    else {
                        continue;
                    };
                    if !generated_collation_supported(document, &definition, sources) {
                        continue;
                    }
                }
                SimpleGeneratedOwnerContext::Bibliography(state) => {
                    let instruction = field.effective_instruction();
                    if instruction.name != "BIBLIOGRAPHY"
                        || state
                            .bibliography_blocks(
                                &instruction,
                                bibliography_story_text_width(document, story, span),
                            )?
                            .is_none()
                    {
                        continue;
                    }
                }
            }
            let raw = xml_fragment_with_namespaces(
                &story.xml[span.field_start..span.field_end],
                &span.start_paragraph_namespaces,
                "simple generated owner attributes",
            )?;
            let mut reader = NsReader::from_reader(raw.as_slice());
            let mut buffer = Vec::new();
            let (_, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| {
                    Error::Other(format!("invalid simple generated field: {error}"))
                })?;
            let start = match event {
                Event::Start(start) | Event::Empty(start) => start,
                _ => continue,
            };
            let mut cache_bindings = span.start_paragraph_namespaces.clone();
            let mut convertible = true;
            for attribute in start.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::Other(format!("invalid simple generated attribute: {error}"))
                })?;
                if attribute.key.as_ref() == b"xmlns"
                    || attribute.key.as_ref().starts_with(b"xmlns:")
                {
                    let prefix = attribute
                        .key
                        .as_ref()
                        .strip_prefix(b"xmlns:")
                        .map(|prefix| String::from_utf8_lossy(prefix).into_owned())
                        .unwrap_or_default();
                    let namespace = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
                        .map_err(|error| {
                            Error::Other(format!("invalid simple generated namespace: {error}"))
                        })?;
                    cache_bindings.insert(prefix, namespace.into_owned());
                    continue;
                }
                let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                if !namespace_is_word(&namespace)
                    || !matches!(local.as_ref(), b"instr" | b"dirty" | b"fldLock")
                {
                    convertible = false;
                }
            }
            if !convertible {
                continue;
            }
            // Retain the exact producer instruction and result XML during form expansion.
            let mut begin = format!("<w:r xmlns:w=\"{W_NS}\"><w:fldChar w:fldCharType=\"begin\"");
            if let Some(locked) = field.locked() {
                begin.push_str(if locked {
                    " w:fldLock=\"1\""
                } else {
                    " w:fldLock=\"0\""
                });
            }
            if let Some(dirty) = field.dirty {
                begin.push_str(if dirty {
                    " w:dirty=\"1\""
                } else {
                    " w:dirty=\"0\""
                });
            }
            begin.push_str("/></w:r>");
            let mut replacement = begin.into_bytes();
            replacement.extend_from_slice(format!("<w:r xmlns:w=\"{W_NS}\"><w:instrText xml:space=\"preserve\">{}</w:instrText></w:r><w:r xmlns:w=\"{W_NS}\"><w:fldChar w:fldCharType=\"separate\"/></w:r>", xml_escape_text(&span.instruction)).as_bytes());
            replacement.extend_from_slice(&generated_simple_cache_xml(
                &story.xml[span.result_start..span.result_end],
                &cache_bindings,
            )?);
            replacement.extend_from_slice(
                format!("<w:r xmlns:w=\"{W_NS}\"><w:fldChar w:fldCharType=\"end\"/></w:r>")
                    .as_bytes(),
            );
            let offset = |value: usize| {
                value
                    .checked_sub(story.wrapper_len)
                    .and_then(|value| value.checked_add(story.range.start))
                    .ok_or_else(|| {
                        Error::Other("simple generated owner is outside its physical story".into())
                    })
            };
            edits
                .entry(story.story.part_name().into())
                .or_default()
                .push(FieldSourceEdit {
                    start: offset(span.field_start)?,
                    end: offset(span.field_end)?,
                    replacement,
                });
        }
    }
    for (part, mut edits) in edits {
        edits.sort_by_key(|edit| edit.start);
        if edits.windows(2).any(|pair| pair[0].end > pair[1].start) {
            return Err(Error::Other("simple generated owners overlap".into()));
        }
        let mut xml = document
            .package
            .get_part(&part)
            .ok_or_else(|| Error::Other("simple generated part disappeared".into()))?
            .to_vec();
        for edit in edits.into_iter().rev() {
            xml.splice(edit.start..edit.end, edit.replacement);
        }
        validate_story_document_declarations_and_doctype(&xml)?;
        document.package.set_part(&part, xml);
    }
    let reopened = document.clone_for_staging().reopen_prepared_staged()?;
    *document = reopened;
    Ok(())
}

pub(crate) fn bibliography_reference_instructions(
    document: &Document,
) -> Result<Vec<FieldInstruction>> {
    let mut instructions = Vec::new();
    for story in generated_story_inventory(document, DynamicOwnerPolicy::Bibliography)? {
        for span in &story.spans {
            let field = parse_dynamic_toc_field(&story.xml, span)?;
            let instruction = field.effective_instruction();
            if instruction.name == "CITATION" {
                instructions.push(instruction);
            }
        }
    }
    Ok(instructions)
}

struct BibliographyInstructionSource {
    text: String,
    units: Vec<(std::ops::Range<usize>, std::ops::Range<usize>)>,
    insertion: Option<usize>,
    attribute: bool,
}

fn bibliography_push_instruction_text(
    source: &mut BibliographyInstructionSource,
    raw: &[u8],
    start: usize,
) -> Result<()> {
    let text = std::str::from_utf8(raw)
        .map_err(|_| Error::Other("bibliography instruction is not UTF-8".into()))?;
    let mut characters = text.char_indices().peekable();
    while let Some((offset, mut character)) = characters.next() {
        let mut end = offset + character.len_utf8();
        if character == '\r' {
            if characters
                .peek()
                .is_some_and(|(_, character)| *character == '\n')
            {
                end = characters.next().expect("peeked character exists").0 + 1;
            }
            character = '\n';
        }
        if source.attribute && matches!(character, '\n' | '\t') {
            character = ' ';
        }
        let begin = source.text.len();
        source.text.push(character);
        source
            .units
            .push((begin..source.text.len(), start + offset..start + end));
    }
    source.insertion = Some(start + raw.len());
    Ok(())
}

fn bibliography_push_instruction_reference(
    source: &mut BibliographyInstructionSource,
    raw: &[u8],
    start: usize,
) -> Result<()> {
    let raw_text = std::str::from_utf8(raw)
        .map_err(|_| Error::Other("bibliography reference is not UTF-8".into()))?;
    let value = quick_xml::escape::unescape(raw_text).map_err(|error| {
        Error::Other(format!("invalid bibliography character reference: {error}"))
    })?;
    if value.chars().count() != 1 {
        return Err(Error::Other(
            "ambiguous bibliography character reference".into(),
        ));
    }
    let begin = source.text.len();
    source.text.push_str(&value);
    source
        .units
        .push((begin..source.text.len(), start..start + raw.len()));
    source.insertion = Some(start + raw.len());
    Ok(())
}

fn bibliography_instruction_source(
    xml: &[u8],
    span: &DynamicTocSpan,
) -> Result<BibliographyInstructionSource> {
    let mut source = BibliographyInstructionSource {
        text: String::new(),
        units: Vec::new(),
        insertion: None,
        attribute: span.simple_field.is_some(),
    };
    let fragments = if span.simple_field.is_some() {
        vec![(
            span.field_start,
            span.field_end,
            &span.start_paragraph_namespaces,
        )]
    } else {
        span.instruction_runs
            .iter()
            .map(|run| (run.start, run.end, &run.inherited_namespaces))
            .collect()
    };
    for (begin, end, namespaces) in fragments {
        let mut wrapper = String::from("<instructionSource");
        for (prefix, namespace) in namespaces {
            if prefix == "xml" {
                continue;
            }
            if prefix.is_empty() {
                wrapper.push_str(" xmlns=\"");
            } else {
                wrapper.push_str(&format!(" xmlns:{prefix}=\""));
            }
            wrapper.push_str(&xml_escape_attribute(namespace));
            wrapper.push('"');
        }
        wrapper.push('>');
        let wrapper_len = wrapper.len();
        let mut wrapped = wrapper.into_bytes();
        wrapped.extend_from_slice(&xml[begin..end]);
        wrapped.extend_from_slice(b"</instructionSource>");
        let mut reader = NsReader::from_reader(wrapped.as_slice());
        let mut buffer = Vec::new();
        let mut elements = Vec::<(bool, Vec<u8>)>::new();
        let mut active = None;
        loop {
            let before = reader.buffer_position() as usize;
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::Other(format!("bibliography instruction XML: {error}")))?;
            let word = namespace_is_word(&namespace);
            let event = event.into_owned();
            let after = reader.buffer_position() as usize;
            let physical = |offset: usize| -> Result<usize> {
                offset
                    .checked_sub(wrapper_len)
                    .and_then(|offset| begin.checked_add(offset))
                    .filter(|offset| *offset <= end)
                    .ok_or_else(|| {
                        Error::Other("bibliography instruction range escapes its source".into())
                    })
            };
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if active.is_some() {
                        return Err(Error::Other("mixed bibliography instruction XML".into()));
                    }
                    let local = element.local_name().as_ref().to_vec();
                    let empty = wrapped.get(after.saturating_sub(2)..after) == Some(b"/>");
                    if source.attribute && word && local == b"fldSimple" {
                        for attribute in element.attributes() {
                            let attribute =
                                attribute.map_err(|error| Error::Other(error.to_string()))?;
                            let (namespace, local) =
                                reader.resolver().resolve_attribute(attribute.key);
                            if namespace_is_word(&namespace) && local.as_ref() == b"instr" {
                                let (start, finish) = attribute_value_span(
                                    &wrapped[before..after],
                                    attribute.key.as_ref(),
                                )
                                .ok_or_else(|| {
                                    Error::Other(
                                        "bibliography instruction attribute has no source span"
                                            .into(),
                                    )
                                })?;
                                let start = before + start;
                                let finish = before + finish;
                                let mut text_reader =
                                    quick_xml::Reader::from_reader(&wrapped[start..finish]);
                                let mut text_buffer = Vec::new();
                                loop {
                                    let text_before = text_reader.buffer_position() as usize;
                                    let event = text_reader
                                        .read_event_into(&mut text_buffer)
                                        .map_err(|error| Error::Other(error.to_string()))?;
                                    let text_after = text_reader.buffer_position() as usize;
                                    match event {
                                        Event::Text(_) => bibliography_push_instruction_text(
                                            &mut source,
                                            &wrapped[start + text_before..start + text_after],
                                            physical(start + text_before)?,
                                        )?,
                                        Event::GeneralRef(_) => {
                                            bibliography_push_instruction_reference(
                                                &mut source,
                                                &wrapped[start + text_before..start + text_after],
                                                physical(start + text_before)?,
                                            )?
                                        }
                                        Event::Eof => break,
                                        _ => {
                                            return Err(Error::Other(
                                                "ambiguous bibliography instruction attribute"
                                                    .into(),
                                            ));
                                        }
                                    }
                                    text_buffer.clear();
                                }
                                source.insertion = Some(physical(finish)?);
                                let expected = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        element.decoder(),
                                    )
                                    .map_err(|error| Error::Other(error.to_string()))?;
                                if source.text != expected {
                                    return Err(Error::Other("bibliography instruction lexical map disagrees with XML decoding".into()));
                                }
                            }
                        }
                    } else if !source.attribute
                        && word
                        && local == b"instrText"
                        && elements
                            .last()
                            .is_some_and(|(word, local)| *word && local == b"r")
                        && !empty
                    {
                        active = Some(elements.len() + 1);
                        source.insertion = Some(physical(after)?);
                    }
                    if !empty {
                        elements.push((word, local));
                    }
                }
                Event::End(_) => {
                    if active == Some(elements.len()) {
                        active = None;
                        source.insertion = Some(physical(before)?);
                    }
                    elements.pop();
                }
                Event::Text(_) if active.is_some() => bibliography_push_instruction_text(
                    &mut source,
                    &wrapped[before..after],
                    physical(before)?,
                )?,
                Event::GeneralRef(_) if active.is_some() => {
                    bibliography_push_instruction_reference(
                        &mut source,
                        &wrapped[before..after],
                        physical(before)?,
                    )?
                }
                Event::CData(value) if active.is_some() => {
                    bibliography_push_instruction_text(
                        &mut source,
                        value.as_ref(),
                        physical(before + 9)?,
                    )?;
                    source.insertion = Some(physical(after)?);
                }
                Event::DocType(_) => {
                    return Err(Error::Other("bibliography instruction DOCTYPE".into()));
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
    }
    if source.insertion.is_none() || source.text.trim() != span.instruction.trim() {
        return Err(Error::Other(
            "bibliography instruction source and owned field disagree".into(),
        ));
    }
    Ok(source)
}

pub(crate) fn patch_bibliography_instruction_options(
    document: &mut Document,
    switches: &[rdocx_oxml::text::FieldSwitch],
) -> Result<()> {
    let stories = generated_story_inventory(document, DynamicOwnerPolicy::Bibliography)?;
    let mut edits = BTreeMap::<String, Vec<FieldSourceEdit>>::new();
    for story in stories {
        for span in &story.spans {
            if Field::new(&span.instruction, "").instruction.name != "BIBLIOGRAPHY" {
                continue;
            }
            let source = bibliography_instruction_source(&story.xml, span)?;
            let replacements =
                rdocx_oxml::text::bibliography_option_switch_edits(&source.text, switches)?;
            let offset = |offset: usize| -> Result<usize> {
                offset
                    .checked_sub(story.wrapper_len)
                    .and_then(|offset| story.range.start.checked_add(offset))
                    .filter(|offset| *offset <= story.range.end)
                    .ok_or_else(|| {
                        Error::Other(
                            "bibliography instruction edit escapes its physical owner".into(),
                        )
                    })
            };
            let part_edits = edits.entry(story.story.part_name().to_owned()).or_default();
            for (range, replacement) in replacements {
                if range.is_empty() {
                    let insertion = offset(source.insertion.ok_or_else(|| {
                        Error::Other("bibliography code insertion boundary disappeared".into())
                    })?)?;
                    let replacement = if source.attribute {
                        xml_escape_attribute(&replacement)
                            .replace('\r', "&#13;")
                            .replace('\n', "&#10;")
                            .replace('\t', "&#9;")
                    } else {
                        xml_escape_text(&replacement).replace('\r', "&#13;")
                    };
                    part_edits.push(FieldSourceEdit {
                        start: insertion,
                        end: insertion,
                        replacement: replacement.into_bytes(),
                    });
                } else {
                    for (logical, physical) in &source.units {
                        if logical.end <= range.start || range.end <= logical.start {
                            continue;
                        }
                        if logical.start < range.start || range.end < logical.end {
                            return Err(Error::Other(
                                "bibliography token splits an XML character".into(),
                            ));
                        }
                        part_edits.push(FieldSourceEdit {
                            start: offset(physical.start)?,
                            end: offset(physical.end)?,
                            replacement: Vec::new(),
                        });
                    }
                }
            }
        }
    }
    for (part, mut part_edits) in edits {
        part_edits.sort_by_key(|edit| (edit.start, edit.end));
        if part_edits
            .windows(2)
            .any(|pair| pair[0].end > pair[1].start)
        {
            return Err(Error::Other(
                "bibliography instruction source edits overlap".into(),
            ));
        }
        let mut xml = document
            .package
            .get_part(&part)
            .ok_or_else(|| Error::Other("bibliography physical source disappeared".into()))?
            .to_vec();
        for edit in part_edits.into_iter().rev() {
            xml.splice(edit.start..edit.end, edit.replacement);
        }
        validate_strict_xml_1_0(&xml).map_err(|error| {
            Error::Other(format!("invalid bibliography instruction edit: {error:?}"))
        })?;
        document.package.set_part(&part, xml);
    }
    Ok(())
}

// Source instructions alone establish numbering. Generated caches are excluded by the shared inventory.
pub(crate) fn bibliography_citation_encounter_tags(document: &Document) -> Result<Vec<String>> {
    use rdocx_oxml::text::FieldArgument;
    let mut tags = Vec::new();
    for story in generated_story_inventory(document, DynamicOwnerPolicy::Bibliography)? {
        for span in &story.spans {
            let field = parse_dynamic_toc_field(&story.xml, span)?;
            let instruction = field.effective_instruction();
            if instruction.name != "CITATION" {
                continue;
            }
            let Some(FieldArgument::Text(tag)) = instruction.arguments.first() else {
                return Err(Error::Other(
                    "citation source identity is missing or nested".into(),
                ));
            };
            if tag.trim().is_empty() {
                return Err(Error::Other("empty citation source identity".into()));
            }
            tags.push(tag.clone());
            for switch in instruction
                .switches
                .iter()
                .take(crate::bibliography::NUMERIC_CITATION_SWITCH_LIMIT)
            {
                if switch.name != "m" {
                    continue;
                }
                let Some(FieldArgument::Text(tag)) = &switch.argument else {
                    return Err(Error::Other(
                        "citation grouped identity is missing or nested".into(),
                    ));
                };
                if tag.trim().is_empty() {
                    return Err(Error::Other(
                        "empty grouped citation source identity".into(),
                    ));
                }
                tags.push(tag.clone());
            }
        }
    }
    Ok(tags)
}

fn preserve_bibliography_div_group(
    story: &GeneratedStorySource,
    span: &DynamicTocSpan,
    blocks: &mut [rdocx_oxml::document::BodyContent],
) -> Result<()> {
    use rdocx_oxml::document::BodyContent;
    if !blocks
        .iter()
        .any(|block| matches!(block, BodyContent::Table(_)))
    {
        return Ok(());
    }
    let mut reader = NsReader::from_reader(story.xml.as_slice());
    let mut buffer = Vec::new();
    let mut parents = Vec::<Vec<u8>>::new();
    let mut table_depth = 0;
    let mut group = None;
    let mut row_particle = None;
    loop {
        let start_offset = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid bibliography producer identity: {error}"))
        })?;
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let empty = story.xml[reader.buffer_position() as usize - 2] == b'/';
                let (namespace, local) = reader.resolver().resolve_element(start.name());
                let word = matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == W_NS.as_bytes());
                let local = if word {
                    local.as_ref().to_vec()
                } else {
                    Vec::new()
                };
                if local == b"divId"
                    && start_offset >= span.result_start
                    && start_offset < span.result_end
                {
                    let row = parents.last().is_some_and(|parent| parent == b"trPr");
                    let trailer =
                        table_depth == 0 && parents.last().is_some_and(|parent| parent == b"pPr");
                    if !row && !trailer {
                        return Err(Error::Other(
                            "ambiguous bibliography divId association".into(),
                        ));
                    }
                    let mut identity = None;
                    for attribute in start.attributes() {
                        let attribute =
                            attribute.map_err(|error| Error::Other(error.to_string()))?;
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        if local.as_ref() == b"val"
                            && matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == W_NS.as_bytes())
                        {
                            if identity.is_some() {
                                return Err(Error::Other(
                                    "duplicate bibliography divId value".into(),
                                ));
                            }
                            identity = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        start.decoder(),
                                    )
                                    .map_err(|error| Error::Other(error.to_string()))?
                                    .parse::<u32>()
                                    .map_err(|_| {
                                        Error::Other("invalid bibliography divId value".into())
                                    })?,
                            );
                        } else if attribute.key.as_ref() != b"xmlns"
                            && !attribute.key.as_ref().starts_with(b"xmlns:")
                        {
                            return Err(Error::Other("unmodeled bibliography divId attributes have no proven group association".into()));
                        }
                    }
                    let identity = identity
                        .ok_or_else(|| Error::Other("missing bibliography divId value".into()))?;
                    if group.is_some_and(|value| value != identity) {
                        return Err(Error::Other(
                            "conflicting bibliography divId group identities".into(),
                        ));
                    }
                    group = Some(identity);
                    if !empty {
                        let content_start = reader.buffer_position() as usize;
                        let end = reader
                            .read_to_end_into(start.name(), &mut Vec::new())
                            .map_err(|error| Error::Other(error.to_string()))?;
                        if !story.xml[content_start..end.end as usize]
                            .iter()
                            .all(u8::is_ascii_whitespace)
                        {
                            return Err(Error::Other(
                                "unmodeled bibliography divId content has no proven group association".into(),
                            ));
                        }
                    }
                    if row {
                        let raw = &story.xml[start_offset..reader.buffer_position() as usize];
                        let scope =
                            crate::document::story_namespace_scope_at(&story.xml, start_offset)?;
                        let closed =
                            crate::document::close_content_fragment_namespaces(raw, &scope)?;
                        if row_particle.as_ref().is_some_and(|value| value != &closed) {
                            return Err(Error::Other("different bibliography divId row particles have no proven group association".into()));
                        }
                        row_particle = Some(closed);
                    }
                    buffer.clear();
                    continue;
                }
                if !empty {
                    if local == b"tbl" {
                        table_depth += 1;
                    }
                    parents.push(local);
                }
            }
            Event::End(_) => {
                if parents.pop().as_deref() == Some(b"tbl") {
                    table_depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if let Some(identity) = group {
        let particle = row_particle
            .ok_or_else(|| Error::Other("bibliography divId lacks a proven row group".into()))?;
        for block in blocks {
            match block {
                BodyContent::Table(table) => {
                    for row in &mut table.rows {
                        row.properties
                            .get_or_insert_with(Default::default)
                            .extra_xml
                            .insert(0, (0, particle.clone()));
                    }
                }
                BodyContent::Paragraph(paragraph)
                    if paragraph
                        .properties
                        .as_ref()
                        .and_then(|properties| properties.rpr.as_ref())
                        .is_some_and(|properties| {
                            properties.font_east_asia.as_deref() == Some("Times New Roman")
                        }) =>
                {
                    paragraph.properties.as_mut().unwrap().div_id = Some(identity);
                }
                _ => {}
            }
        }
    }
    Ok(())
}

fn bibliography_story_text_width(
    document: &Document,
    story: &GeneratedStorySource,
    span: &DynamicTocSpan,
) -> Option<i32> {
    if story.story.kind() != crate::StoryKind::Body {
        return None;
    }
    // Physical owner ranges exclude the final section particle. Keep that existing
    // document section as the fallback behind any intervening paragraph section.
    let mut body = story.body.clone();
    body.sect_pr = document.document.body.sect_pr.clone();
    Some(toc_section_text_width(&body, span.begin_paragraph))
}

pub(crate) fn update_bibliography_caches(
    document: &mut Document,
    state: &crate::bibliography::BibliographyUpdateState,
) -> Result<crate::BibliographyUpdateReport> {
    generated_normalize_simple_tables(document, SimpleGeneratedOwnerContext::Bibliography(state))?;
    let mut report = crate::BibliographyUpdateReport {
        updated_citations: 0,
        rebuilt_bibliographies: 0,
        diagnostics: Vec::new(),
    };
    let mut edits = BTreeMap::<String, Vec<FieldSourceEdit>>::new();
    for story in generated_story_inventory(document, DynamicOwnerPolicy::Bibliography)? {
        let paragraph_properties =
            crate::bibliography::bibliography_paragraph_property_ranges(&story.xml)?;
        for span in &story.spans {
            let field = parse_dynamic_toc_field(&story.xml, span)?;
            let instruction = field.effective_instruction();
            if !matches!(instruction.name.as_str(), "CITATION" | "BIBLIOGRAPHY") {
                continue;
            }
            if field.locked() == Some(true) {
                report.diagnostics.push(format!(
                    "locked {} retains its complete cache",
                    instruction.name
                ));
                continue;
            }
            if instruction.name == "BIBLIOGRAPHY" {
                let Some(mut blocks) = state.bibliography_blocks(
                    &instruction,
                    bibliography_story_text_width(document, &story, span),
                )?
                else {
                    report
                        .diagnostics
                        .push("noncatalogue bibliography retains its complete cache".into());
                    continue;
                };
                preserve_bibliography_div_group(&story, span, &mut blocks)?;
                if span.simple_field.is_some() {
                    report.diagnostics.push("simple bibliography with unmodeled producer attributes retains its complete owner".into());
                    continue;
                }
                if !span.separator_wrapper_names.is_empty() || !span.end_wrapper_prefixes.is_empty()
                {
                    return Err(Error::Other(
                        "catalogued bibliography owner expansion is still being implemented".into(),
                    ));
                }
                let Some(rdocx_oxml::document::BodyContent::Paragraph(first)) = blocks.first()
                else {
                    return Err(Error::Other(
                        "bibliography formatter returned no entry boundary".into(),
                    ));
                };
                let mut writer = quick_xml::Writer::new(Vec::new());
                first
                    .properties
                    .as_ref()
                    .ok_or_else(|| Error::Other("bibliography first properties missing".into()))?
                    .to_xml(&mut writer)?;
                let properties = xml_fragment_with_namespaces(
                    &writer.into_inner(),
                    &BTreeMap::from([("w".into(), W_NS.into())]),
                    "native bibliography paragraph properties",
                )?;
                let paragraph_start = span.instruction_paragraph_start;
                let mut reader =
                    quick_xml::Reader::from_reader(&story.xml[paragraph_start..span.result_start]);
                let mut buffer = Vec::new();
                let Event::Start(start) = reader.read_event_into(&mut buffer).map_err(|error| {
                    Error::Other(format!("invalid bibliography first paragraph: {error}"))
                })?
                else {
                    return Err(Error::Other(
                        "bibliography first paragraph boundary changed".into(),
                    ));
                };
                if start.name().as_ref() != span.start_paragraph_name.as_bytes() {
                    return Err(Error::Other(
                        "bibliography first paragraph ownership changed".into(),
                    ));
                }
                let property_offset = paragraph_start + reader.buffer_position() as usize;
                let existing_properties = paragraph_properties.get(&paragraph_start).cloned();
                let relocated_properties = if let Some(range) = &existing_properties {
                    if story.xml[range.clone()] != properties {
                        if span.begin_paragraph != span.end_paragraph {
                            None
                        } else {
                            Some(crate::bibliography::bibliography_end_paragraph_properties(
                                &story.xml[range.clone()],
                                &span.start_paragraph_namespaces,
                            )?)
                        }
                    } else {
                        None
                    }
                } else {
                    None
                };
                let offset = |offset: usize| {
                    offset
                        .checked_sub(story.wrapper_len)
                        .and_then(|offset| offset.checked_add(story.range.start))
                        .filter(|offset| *offset <= story.range.end)
                        .ok_or_else(|| {
                            Error::Other("bibliography cache escaped its physical owner".into())
                        })
                };
                if let Some(range) = &existing_properties {
                    let replacement = if story.xml[range.clone()] == properties {
                        properties.clone()
                    } else {
                        crate::bibliography::bibliography_first_paragraph_properties(
                            &story.xml[range.clone()],
                            &properties,
                            &span.start_paragraph_namespaces,
                        )?
                    };
                    if story.xml[range.clone()] != replacement {
                        edits
                            .entry(story.story.part_name().to_owned())
                            .or_default()
                            .push(FieldSourceEdit {
                                start: offset(range.start)?,
                                end: offset(range.end)?,
                                replacement,
                            });
                    }
                }
                if existing_properties.is_none() {
                    edits
                        .entry(story.story.part_name().to_owned())
                        .or_default()
                        .push(FieldSourceEdit {
                            start: offset(property_offset)?,
                            end: offset(property_offset)?,
                            replacement: properties,
                        });
                }
                let mut writer = quick_xml::Writer::new(Vec::new());
                for run in &first.runs {
                    run.to_xml(&mut writer)?;
                }
                let first_runs = writer.into_inner();
                let mut replacement = if first_runs.is_empty() {
                    Vec::new()
                } else {
                    xml_fragment_with_namespaces(
                        &first_runs,
                        &BTreeMap::from([("w".into(), W_NS.into())]),
                        "native bibliography first entry",
                    )?
                };
                replacement
                    .extend_from_slice(format!("</{}>", span.start_paragraph_name).as_bytes());
                for block in blocks.iter().skip(1) {
                    let mut writer = quick_xml::Writer::new(Vec::new());
                    match block {
                        rdocx_oxml::document::BodyContent::Paragraph(paragraph) => {
                            paragraph.to_xml(&mut writer)?;
                        }
                        rdocx_oxml::document::BodyContent::Table(table) => {
                            table.to_xml(&mut writer)?;
                        }
                        _ => {
                            return Err(Error::Other(
                                "bibliography formatter returned an unsupported owned block".into(),
                            ));
                        }
                    }
                    replacement.extend_from_slice(&xml_fragment_with_namespaces(
                        &writer.into_inner(),
                        &BTreeMap::from([("w".into(), W_NS.into())]),
                        "native bibliography interior entry",
                    )?);
                }
                if span.begin_paragraph == span.end_paragraph {
                    // The new end boundary has no producer paragraph identity to duplicate.
                    // Original end controls and their outside suffix remain byte-for-byte in place.
                    let mut end_start = format!("<{}", span.start_paragraph_name);
                    for (prefix, namespace) in &span.start_paragraph_namespaces {
                        if prefix == "xml" {
                            continue;
                        }
                        let name = if prefix.is_empty() {
                            "xmlns".to_owned()
                        } else {
                            format!("xmlns:{prefix}")
                        };
                        end_start
                            .push_str(&format!(" {name}=\"{}\"", xml_escape_attribute(namespace)));
                    }
                    end_start.push('>');
                    replacement.extend_from_slice(end_start.as_bytes());
                    if let Some(properties) = &relocated_properties {
                        replacement.extend_from_slice(properties);
                    }
                } else {
                    replacement.extend_from_slice(
                        &story.xml[span.end_paragraph_start..span.end_paragraph_content_start],
                    );
                }
                edits
                    .entry(story.story.part_name().to_owned())
                    .or_default()
                    .push(FieldSourceEdit {
                        start: offset(span.result_start)?,
                        end: offset(span.result_end)?,
                        replacement,
                    });
                report.rebuilt_bibliographies += 1;
                continue;
            }
            let Some(mut runs) = state.citation_runs(&instruction)? else {
                report
                    .diagnostics
                    .push("noncatalogue citation retains its complete cache".into());
                continue;
            };
            if span.begin_paragraph != span.end_paragraph {
                report.diagnostics.push(
                    "citation with producer block topology retains its complete cache".into(),
                );
                continue;
            }
            if instruction.switches.iter().any(|switch| switch.name == "*"
                && matches!(&switch.argument, Some(rdocx_oxml::text::FieldArgument::Text(value)) if value.eq_ignore_ascii_case("MERGEFORMAT"))) {
                if !span.separator_wrapper_names.is_empty() || !span.end_wrapper_prefixes.is_empty() {
                    return Err(Error::Other("ambiguous citation cache-format wrapper".into()));
                }
                let scope = crate::document::story_namespace_scope_at(&story.xml, span.result_start)?;
                let cache = crate::document::close_content_fragment_namespaces(
                    &story.xml[span.result_start..span.result_end], &scope)?;
                let mut paragraph = format!("<w:p xmlns:w=\"{W_NS}\">").into_bytes();
                paragraph.extend_from_slice(&cache);
                paragraph.extend_from_slice(b"</w:p>");
                let previous = CT_P::from_xml_fragment(&paragraph)?;
                let properties = previous.runs.first().and_then(|run| run.properties.as_ref());
                if previous.runs.is_empty() || !previous.extra_xml.is_empty() || !previous.hyperlinks.is_empty() || !previous.comment_ranges.is_empty() || !previous.bookmark_markers.is_empty() || !previous.content_controls.is_empty() || !previous.revisions.is_empty() || !previous.equations.is_empty() || !previous.rubies.is_empty() || previous.runs.iter().any(|run|
                    !run.extra_xml.is_empty() || run.properties.as_ref() != properties
                    || run.content.iter().any(|content| !matches!(content, RunContent::Text(_))))
                    || runs.windows(2).any(|pair| pair[0].properties != pair[1].properties)
                {
                    return Err(Error::Other("ambiguous citation cache-format association".into()));
                }
                let mut replacement = rdocx_oxml::text::CT_R::new(
                    &runs.iter().map(|run| run.text()).collect::<String>());
                // merge_from cascades modeled properties. Keep the original raw particles on
                // the source clone while supplying generated defaults for missing values.
                let mut merged = properties.cloned().unwrap_or_default();
                let mut effective = runs.first().and_then(|run| run.properties.clone()).unwrap_or_default();
                effective.merge_from(&merged);
                merged.merge_from(&effective);
                replacement.properties = Some(merged);
                runs = vec![replacement];
            }
            let mut writer = quick_xml::Writer::new(Vec::new());
            for run in &runs {
                run.to_xml(&mut writer)?;
            }
            let replacement = xml_fragment_with_namespaces(
                &writer.into_inner(),
                &BTreeMap::from([("w".into(), W_NS.into())]),
                "native citation cache",
            )?;
            let offset = |offset: usize| {
                offset
                    .checked_sub(story.wrapper_len)
                    .and_then(|offset| offset.checked_add(story.range.start))
                    .filter(|offset| *offset <= story.range.end)
                    .ok_or_else(|| Error::Other("citation cache escaped its physical owner".into()))
            };
            edits
                .entry(story.story.part_name().to_owned())
                .or_default()
                .push(FieldSourceEdit {
                    start: offset(span.result_start)?,
                    end: offset(span.result_end)?,
                    replacement,
                });
            report.updated_citations += 1;
        }
    }
    for (part, mut edits) in edits {
        edits.sort_by_key(|edit| edit.start);
        if edits.windows(2).any(|pair| pair[0].end > pair[1].start) {
            return Err(Error::Other("citation cache edits overlap".into()));
        }
        let mut xml = document
            .package
            .get_part(&part)
            .ok_or_else(|| Error::Other("citation story disappeared".into()))?
            .to_vec();
        for edit in edits.into_iter().rev() {
            xml.splice(edit.start..edit.end, edit.replacement);
        }
        validate_strict_xml_1_0(&xml)
            .map_err(|error| Error::Other(format!("invalid citation cache XML: {error:?}")))?;
        document.package.set_part(&part, xml);
    }
    Ok(report)
}
