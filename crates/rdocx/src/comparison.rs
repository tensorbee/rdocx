//! Deterministic native document comparison and tracked-revision generation.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer};
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Document};
use rdocx_oxml::namespace::W_NS;
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_TblPr, CT_Tc, CT_TrPr, CellContent};
use rdocx_oxml::text::{CT_P, CT_R, CT_Text, RunContent};

use crate::revision::validate_revision_timestamp;
use crate::{Document, Error, Result};

use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;

#[cfg(test)]
thread_local! {
    static FAIL_AFTER_COMPARISON_STAGING: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

type ControlPropertySignature<'a> = Option<(
    Option<&'a str>,
    Option<&'a str>,
    Option<i32>,
    Option<rdocx_oxml::content_control::SdtType>,
    Option<&'a rdocx_oxml::content_control::CT_DataBinding>,
)>;

/// A comparison difference that cannot be represented as a content revision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComparisonDiagnostic {
    pub location: String,
    pub message: String,
}

/// A Word story category that can be excluded from document comparison.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ComparisonStoryKind {
    Main,
    Header,
    Footer,
    Comment,
    TextBox,
    Footnote,
    Endnote,
}

/// The text boundary used when generating content revisions.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ComparisonGranularity {
    /// Preserve the legacy whole-run comparison behavior.
    #[default]
    Run,
    /// Compare maximal word, whitespace, and punctuation or symbol units.
    Word,
    /// Compare individual Unicode scalar values.
    Character,
}

/// Policy controls for native document comparison.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ComparisonOptions {
    pub granularity: ComparisonGranularity,
    pub ignore_formatting: bool,
    pub ignore_whitespace: bool,
    pub ignore_fields: bool,
    pub ignore_comments: bool,
    pub ignored_stories: Vec<ComparisonStoryKind>,
}

struct Metadata<'a> {
    author: &'a str,
    timestamp: &'a str,
    options: &'a ComparisonOptions,
    ids: IdAllocator,
}

struct IdAllocator {
    used: HashSet<i32>,
    next: i32,
}

impl ComparisonStoryKind {
    fn label(self) -> &'static str {
        match self {
            Self::Header => "header",
            Self::Footer => "footer",
            Self::Comment => "comments",
            Self::Footnote => "footnotes",
            Self::Endnote => "endnotes",
            Self::Main => "body",
            Self::TextBox => "text-box",
        }
    }

    fn root_local(self) -> &'static str {
        match self {
            Self::Header => "hdr",
            Self::Footer => "ftr",
            Self::Comment => "comments",
            Self::Footnote => "footnotes",
            Self::Endnote => "endnotes",
            Self::Main | Self::TextBox => "document",
        }
    }

    fn owner_local(self) -> Option<&'static str> {
        match self {
            Self::Comment => Some("comment"),
            Self::Footnote => Some("footnote"),
            Self::Endnote => Some("endnote"),
            Self::Main | Self::Header | Self::Footer | Self::TextBox => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StoryPart {
    kind: ComparisonStoryKind,
    part_name: String,
}

#[derive(Default)]
struct TextBoxMarkers {
    main: String,
    related: HashMap<String, String>,
}

impl IdAllocator {
    fn new(used: HashSet<i32>) -> Self {
        Self { used, next: 0 }
    }

    fn allocate(&mut self) -> Result<i32> {
        while self.used.contains(&self.next) {
            self.next = self
                .next
                .checked_add(1)
                .ok_or_else(|| Error::Other("comparison revision ids are exhausted".to_owned()))?;
        }
        let id = self.next;
        self.used.insert(id);
        self.next = self.next.saturating_add(1);
        Ok(id)
    }

    fn revision(
        &mut self,
        kind: &str,
        author: &str,
        timestamp: &str,
        inner: &str,
    ) -> Result<String> {
        let id = self.allocate()?;
        Ok(Self::revision_with_id(kind, author, timestamp, inner, id))
    }

    fn revision_with_id(kind: &str, author: &str, timestamp: &str, inner: &str, id: i32) -> String {
        let author = quick_xml::escape::escape(author);
        let timestamp = quick_xml::escape::escape(timestamp);
        format!(
            r#"<w:{kind} w:id="{id}" w:author="{author}" w:date="{timestamp}">{inner}</w:{kind}>"#
        )
    }

    fn marker(&mut self, kind: &str, author: &str, timestamp: &str) -> Result<String> {
        let id = self.allocate()?;
        Ok(Self::marker_with_id(kind, author, timestamp, id))
    }

    fn marker_with_id(kind: &str, author: &str, timestamp: &str, id: i32) -> String {
        let author = quick_xml::escape::escape(author);
        let timestamp = quick_xml::escape::escape(timestamp);
        format!(r#"<w:{kind} w:id="{id}" w:author="{author}" w:date="{timestamp}"/>"#)
    }
}

struct MovePairs {
    original: Vec<Option<i32>>,
    edited: Vec<Option<i32>>,
}

fn pair_moves(
    aligned: &[(Option<usize>, Option<usize>)],
    original: &[String],
    edited: &[String],
    ids: &mut IdAllocator,
) -> Result<MovePairs> {
    let mut pairs = MovePairs {
        original: vec![None; original.len()],
        edited: vec![None; edited.len()],
    };
    let unmatched_edited = aligned
        .iter()
        .filter_map(|(left, right)| left.is_none().then_some(*right).flatten())
        .collect::<Vec<_>>();
    for left in aligned
        .iter()
        .filter_map(|(left, right)| right.is_none().then_some(*left).flatten())
    {
        let Some(right) = unmatched_edited
            .iter()
            .copied()
            .find(|right| pairs.edited[*right].is_none() && original[left] == edited[*right])
        else {
            continue;
        };
        let id = ids.allocate()?;
        pairs.original[left] = Some(id);
        pairs.edited[right] = Some(id);
    }
    Ok(pairs)
}

impl Document {
    /// Compare every supported Word story with `edited` and record tracked changes.
    pub fn compare(
        &mut self,
        edited: &Document,
        author: &str,
        timestamp: &str,
    ) -> Result<Vec<ComparisonDiagnostic>> {
        self.compare_with_options(edited, author, timestamp, &ComparisonOptions::default())
    }

    /// Compare supported Word stories using an explicit granularity and ignore policy.
    pub fn compare_with_options(
        &mut self,
        edited: &Document,
        author: &str,
        timestamp: &str,
        options: &ComparisonOptions,
    ) -> Result<Vec<ComparisonDiagnostic>> {
        validate_revision_timestamp(timestamp)?;
        validate_comparison_options(options)?;
        let mut original = self.clone_for_staging();
        original.flush_to_package()?;
        let mut edited = edited.clone_for_staging();
        edited.flush_to_package()?;
        let original_stories = story_parts_with_options(&original, options)?;
        let edited_stories = story_parts_with_options(&edited, options)?;
        if original_stories != edited_stories {
            return Err(Error::Other(
                "document comparison requires identical related-story shells".to_owned(),
            ));
        }
        if contains_modeled_revisions(&original, &original_stories, options)?
            || contains_modeled_revisions(&edited, &edited_stories, options)?
        {
            return Err(Error::Other(
                "document comparison requires inputs without existing modeled revisions".to_owned(),
            ));
        }

        reject_cross_story_moves(&original, &edited, &original_stories, options)?;

        let original_xml = original.document.to_xml()?;
        let edited_xml = edited.document.to_xml()?;
        let text_box_markers =
            comparison_text_box_markers(&original, &edited, &original_stories, options)?;
        let mut used_ids = if story_ignored(options, ComparisonStoryKind::Main) {
            HashSet::new()
        } else {
            word_ids_with_options(&original_xml, options)?
        };
        if !story_ignored(options, ComparisonStoryKind::Main) {
            used_ids.extend(word_ids_with_options(&edited_xml, options)?);
        }
        for story in &original_stories {
            used_ids.extend(word_ids_with_options(
                story_xml(&original, story)?,
                options,
            )?);
            used_ids.extend(word_ids_with_options(story_xml(&edited, story)?, options)?);
        }
        let mut metadata = Metadata {
            author,
            timestamp,
            options,
            ids: IdAllocator::new(used_ids),
        };
        let mut diagnostics = Vec::new();
        let tracked_body = if story_ignored(options, ComparisonStoryKind::Main) {
            extract_body_inner(&original_xml)?.to_owned()
        } else if story_ignored(options, ComparisonStoryKind::TextBox) {
            compare_story_inner(
                extract_body_inner(&original_xml)?,
                extract_body_inner(&edited_xml)?,
                "body",
                "w",
                &mut metadata,
                &mut diagnostics,
            )?
        } else {
            compare_body(
                &original.document,
                &edited.document,
                "body",
                None,
                &mut metadata,
                &mut diagnostics,
            )?
        };
        let tracked_xml = replace_body_inner(&original_xml, &tracked_body)?;
        let tracked = CT_Document::from_xml(tracked_xml.as_bytes())?;
        tracked.to_xml()?;

        let mut candidate = original.clone_for_staging();
        candidate.document = tracked;
        candidate
            .package
            .set_part(&candidate.doc_part_name, tracked_xml.into_bytes());
        for story in &original_stories {
            let tracked_story = compare_story_part(
                story_xml(&original, story)?,
                story_xml(&edited, story)?,
                story,
                &mut metadata,
                &mut diagnostics,
            )?;
            candidate.package.set_part(&story.part_name, tracked_story);
        }
        candidate = reopen_staged(candidate)?;
        #[cfg(test)]
        FAIL_AFTER_COMPARISON_STAGING.with(|fail| {
            if fail.replace(false) {
                return Err(Error::Other(
                    "injected staged comparison postcondition failure".to_owned(),
                ));
            }
            Ok(())
        })?;
        let mut accepted = candidate.clone_for_staging();
        accepted.accept_all()?;
        let accepted_body =
            normalized_package(&accepted, &original_stories, options, &text_box_markers)?;
        let mut edited_package =
            normalized_package(&edited, &edited_stories, options, &text_box_markers)?;
        if story_ignored(options, ComparisonStoryKind::Main) {
            edited_package.0 =
                normalized_package(&original, &original_stories, options, &text_box_markers)?.0;
        }
        if accepted_body != edited_package {
            return Err(Error::Other(format!(
                "comparison acceptance does not reproduce the edited stories: {accepted_body:?} != {edited_package:?}"
            )));
        }
        let mut rejected = candidate.clone_for_staging();
        rejected.reject_all()?;
        let rejected_package =
            normalized_package(&rejected, &original_stories, options, &text_box_markers)?;
        let original_package =
            normalized_package(&original, &original_stories, options, &text_box_markers)?;
        if rejected_package != original_package {
            return Err(Error::Other(format!(
                "comparison rejection does not reproduce the original stories: {rejected_package:?} != {original_package:?}"
            )));
        }

        self.commit_staged_mutation(candidate);
        Ok(diagnostics)
    }
}

fn story_ignored(options: &ComparisonOptions, kind: ComparisonStoryKind) -> bool {
    options.ignored_stories.contains(&kind)
        || (options.ignore_comments && kind == ComparisonStoryKind::Comment)
}

fn validate_comparison_options(options: &ComparisonOptions) -> Result<()> {
    let mut unique = HashSet::new();
    if options
        .ignored_stories
        .iter()
        .any(|kind| !unique.insert(*kind))
    {
        return Err(Error::Other(
            "comparison options contain a duplicate ignored story".to_owned(),
        ));
    }
    Ok(())
}

fn story_parts(document: &Document) -> Result<Vec<StoryPart>> {
    story_parts_with_options(document, &ComparisonOptions::default())
}

fn story_parts_with_options(
    document: &Document,
    options: &ComparisonOptions,
) -> Result<Vec<StoryPart>> {
    let relationships = document.package.get_part_rels(&document.doc_part_name);
    let mut stories = Vec::new();
    let mut seen = HashSet::new();
    let sections = document
        .document
        .body
        .content
        .iter()
        .filter_map(|content| match content {
            BodyContent::Paragraph(paragraph) => paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.sect_pr.as_ref()),
            BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => None,
        })
        .chain(document.document.body.sect_pr.iter());
    for section in sections {
        for (kind, references) in [
            (ComparisonStoryKind::Header, section.header_refs.as_slice()),
            (ComparisonStoryKind::Footer, section.footer_refs.as_slice()),
        ] {
            if story_ignored(options, kind) {
                continue;
            }
            for reference in references {
                let relationships = relationships.ok_or_else(|| {
                    Error::Other(format!(
                        "{} reference {} has no relationship set",
                        kind.label(),
                        reference.rel_id
                    ))
                })?;
                let matching = relationships
                    .items
                    .iter()
                    .filter(|relationship| relationship.id == reference.rel_id)
                    .collect::<Vec<_>>();
                if matching.len() != 1 {
                    return Err(Error::Other(format!(
                        "{} reference {} has {} relationships",
                        kind.label(),
                        reference.rel_id,
                        matching.len()
                    )));
                }
                let relationship = matching[0];
                let expected_type = if kind == ComparisonStoryKind::Header {
                    rel_types::HEADER
                } else {
                    rel_types::FOOTER
                };
                if relationship.rel_type != expected_type
                    || relationship.target_mode.as_deref() == Some("External")
                {
                    return Err(Error::Other(format!(
                        "{} reference {} has an invalid relationship target",
                        kind.label(),
                        reference.rel_id
                    )));
                }
                let part_name =
                    OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
                if document.package.get_part(&part_name).is_none() {
                    return Err(Error::Other(format!(
                        "{} relationship {} targets missing part {part_name}",
                        kind.label(),
                        reference.rel_id
                    )));
                }
                if seen.insert((kind, part_name.clone())) {
                    stories.push(StoryPart { kind, part_name });
                }
            }
        }
    }

    if let Some(relationships) = relationships {
        for (kind, relationship_type) in [
            (ComparisonStoryKind::Comment, rel_types::COMMENTS),
            (ComparisonStoryKind::Footnote, rel_types::FOOTNOTES),
            (ComparisonStoryKind::Endnote, rel_types::ENDNOTES),
        ] {
            if story_ignored(options, kind) {
                continue;
            }
            for relationship in &relationships.items {
                if relationship.rel_type != relationship_type {
                    continue;
                }
                if relationship.target_mode.as_deref() == Some("External") {
                    return Err(Error::Other(format!(
                        "{} story has an external relationship target",
                        kind.label()
                    )));
                }
                let part_name =
                    OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
                if document.package.get_part(&part_name).is_none() {
                    return Err(Error::Other(format!(
                        "{} relationship {} targets missing part {part_name}",
                        kind.label(),
                        relationship.id
                    )));
                }
                if !seen.insert((kind, part_name.clone())) {
                    return Err(Error::Other(format!(
                        "{} story part {part_name} is referenced more than once",
                        kind.label()
                    )));
                }
                stories.push(StoryPart { kind, part_name });
            }
        }
    }
    Ok(stories)
}

pub(crate) fn related_story_part_names(document: &Document) -> Result<Vec<String>> {
    story_parts(document).map(|stories| stories.into_iter().map(|story| story.part_name).collect())
}

fn story_xml<'a>(document: &'a Document, story: &StoryPart) -> Result<&'a [u8]> {
    document.package.get_part(&story.part_name).ok_or_else(|| {
        Error::Other(format!(
            "missing {} story {}",
            story.kind.label(),
            story.part_name
        ))
    })
}

fn contains_modeled_revisions(
    document: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
) -> Result<bool> {
    if !story_ignored(options, ComparisonStoryKind::Main)
        && modeled_revision_count_with_options(
            document
                .package
                .get_part(&document.doc_part_name)
                .ok_or_else(|| {
                    Error::Other(format!("missing main story {}", document.doc_part_name))
                })?,
            options,
        )? > 0
    {
        return Ok(true);
    }
    stories
        .iter()
        .map(|story| modeled_revision_count_with_options(story_xml(document, story)?, options))
        .try_fold(false, |found, count| count.map(|count| found || count > 0))
}

fn reject_cross_story_moves(
    original: &Document,
    edited: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
) -> Result<()> {
    let mut original_by_story = Vec::new();
    let mut edited_by_story = Vec::new();
    if !story_ignored(options, ComparisonStoryKind::Main) {
        original_by_story.push((
            "document".to_owned(),
            normalized_body_with_options(&original.document, options),
        ));
        edited_by_story.push((
            "document".to_owned(),
            normalized_body_with_options(&edited.document, options),
        ));
    }
    for story in stories {
        let identity = format!("{}:{}", story.kind.label(), story.part_name);
        let original_xml = story_xml(original, story)?;
        let edited_xml = story_xml(edited, story)?;
        let marker = if story_ignored(options, ComparisonStoryKind::TextBox) {
            let original_source = std::str::from_utf8(original_xml).map_err(utf8_error)?;
            let edited_source = std::str::from_utf8(edited_xml).map_err(utf8_error)?;
            let prefix = root_prefix(original_source, story.kind.root_local())?;
            Some(text_box_marker_local(
                original_source,
                edited_source,
                &prefix,
            )?)
        } else {
            None
        };
        original_by_story.push((
            identity.clone(),
            normalized_story_part(original_xml, story.kind, options, marker.as_deref())?,
        ));
        edited_by_story.push((
            identity,
            normalized_story_part(edited_xml, story.kind, options, marker.as_deref())?,
        ));
    }

    let mut removed = HashMap::<String, Vec<String>>::new();
    let mut inserted = HashMap::<String, Vec<String>>::new();
    for ((identity, before), (_, after)) in original_by_story.iter().zip(&edited_by_story) {
        let before = signature_counts(before);
        let after = signature_counts(after);
        for (signature, count) in &before {
            let delta = count.saturating_sub(*after.get(signature).unwrap_or(&0));
            removed
                .entry(signature.clone())
                .or_default()
                .extend(std::iter::repeat_n(identity.clone(), delta));
        }
        for (signature, count) in &after {
            let delta = count.saturating_sub(*before.get(signature).unwrap_or(&0));
            inserted
                .entry(signature.clone())
                .or_default()
                .extend(std::iter::repeat_n(identity.clone(), delta));
        }
    }
    for (signature, sources) in removed {
        let Some(destinations) = inserted.get(&signature) else {
            continue;
        };
        if sources
            .iter()
            .any(|source| destinations.iter().any(|destination| source != destination))
        {
            return Err(Error::Other(
                "comparison cannot represent a move between Word stories".to_owned(),
            ));
        }
    }
    Ok(())
}

fn signature_counts(signatures: &[String]) -> HashMap<String, usize> {
    let mut counts = HashMap::new();
    for signature in signatures {
        *counts.entry(signature.clone()).or_default() += 1;
    }
    counts
}

fn compare_story_part(
    original: &[u8],
    edited: &[u8],
    story: &StoryPart,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<Vec<u8>> {
    let original = std::str::from_utf8(original).map_err(utf8_error)?;
    let edited = std::str::from_utf8(edited).map_err(utf8_error)?;
    let original_root = root_inner_range(original, story.kind.root_local())?;
    let edited_root = root_inner_range(edited, story.kind.root_local())?;
    let word_prefix = root_prefix(original, story.kind.root_local())?;
    let tracked = if let Some(owner_local) = story.kind.owner_local() {
        compare_owned_story(
            original,
            edited,
            original_root,
            edited_root,
            owner_local,
            story,
            &word_prefix,
            metadata,
            diagnostics,
        )?
    } else {
        let tracked_inner = compare_story_inner(
            &original[original_root.clone()],
            &edited[edited_root],
            &format!("{}:{}", story.kind.label(), story.part_name),
            &word_prefix,
            metadata,
            diagnostics,
        )?;
        let mut tracked = original.to_owned();
        tracked.replace_range(original_root, &tracked_inner);
        tracked
    };
    crate::revision::modeled_revision_count(tracked.as_bytes())?;
    Ok(tracked.into_bytes())
}

#[allow(clippy::too_many_arguments)]
fn compare_owned_story(
    original: &str,
    edited: &str,
    original_root: Range<usize>,
    edited_root: Range<usize>,
    owner_local: &str,
    story: &StoryPart,
    word_prefix: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_spans = direct_word_element_spans_prefix_aware(original, owner_local)?;
    let edited_spans = direct_word_element_spans_prefix_aware(edited, owner_local)?;
    if original_spans.len() != edited_spans.len() {
        return Err(Error::Other(format!(
            "{} story shell count changed in {}",
            story.kind.label(),
            story.part_name
        )));
    }
    let original_skeleton = story_skeleton(original, &original_spans);
    let edited_skeleton = story_skeleton(edited, &edited_spans);
    if original_skeleton != edited_skeleton {
        return Err(Error::Other(format!(
            "{} story root shell changed in {}",
            story.kind.label(),
            story.part_name
        )));
    }
    let mut replacements = Vec::with_capacity(original_spans.len());
    for (index, (left, right)) in original_spans.iter().zip(&edited_spans).enumerate() {
        let left_xml = &original[left.clone()];
        let right_xml = &edited[right.clone()];
        if owner_start_signature(left_xml)? != owner_start_signature(right_xml)? {
            return Err(Error::Other(format!(
                "{} owner shell changed at {}[{index}]",
                story.kind.label(),
                story.part_name
            )));
        }
        if matches!(
            story.kind,
            ComparisonStoryKind::Footnote | ComparisonStoryKind::Endnote
        ) && !normal_note_owner(left_xml)?
        {
            if left_xml != right_xml {
                return Err(Error::Other(format!(
                    "{} separator shell changed at {}[{index}]",
                    story.kind.label(),
                    story.part_name
                )));
            }
            replacements.push(left_xml.to_owned());
            continue;
        }
        let left_inner = element_inner_range_any_prefix(left_xml, owner_local)?;
        let right_inner = element_inner_range_any_prefix(right_xml, owner_local)?;
        let tracked_inner = compare_story_inner(
            &left_xml[left_inner.clone()],
            &right_xml[right_inner],
            &format!(
                "{}:{}/{}[{index}]",
                story.kind.label(),
                story.part_name,
                owner_local
            ),
            word_prefix,
            metadata,
            diagnostics,
        )?;
        let mut owner = left_xml.to_owned();
        owner.replace_range(left_inner, &tracked_inner);
        replacements.push(owner);
    }
    let mut tracked = original.to_owned();
    for (span, replacement) in original_spans.into_iter().zip(replacements).rev() {
        tracked.replace_range(span, &replacement);
    }
    if root_inner_range(&tracked, story.kind.root_local())?.start != original_root.start
        || edited_root.start == usize::MAX
    {
        return Err(Error::Other(
            "comparison story root moved unexpectedly".to_owned(),
        ));
    }
    Ok(tracked)
}

fn compare_story_inner(
    original: &str,
    edited: &str,
    location: &str,
    word_prefix: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    compare_story_inner_impl(
        original,
        edited,
        location,
        word_prefix,
        metadata,
        diagnostics,
        true,
    )
}

#[allow(clippy::too_many_arguments)]
fn compare_story_inner_impl(
    original: &str,
    edited: &str,
    location: &str,
    word_prefix: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
    scan_text_boxes: bool,
) -> Result<String> {
    let original_text_boxes = if scan_text_boxes {
        text_box_spans(original, word_prefix)?
    } else {
        Vec::new()
    };
    let edited_text_boxes = if scan_text_boxes {
        text_box_spans(edited, word_prefix)?
    } else {
        Vec::new()
    };
    if !original_text_boxes.is_empty() || !edited_text_boxes.is_empty() {
        if text_box_host_skeletons(original, &original_text_boxes)?
            != text_box_host_skeletons(edited, &edited_text_boxes)?
        {
            return Err(Error::Other(format!(
                "comparison cannot change nested text-box host shells at {location}"
            )));
        }
        if story_ignored(metadata.options, ComparisonStoryKind::TextBox) {
            let marker = text_box_marker_local(original, edited, word_prefix)?;
            let (masked_original, preserved) =
                mask_text_box_subtrees(original, word_prefix, &marker)?;
            let (masked_edited, _) = mask_text_box_subtrees(edited, word_prefix, &marker)?;
            let mut tracked = compare_story_inner_impl(
                &masked_original,
                &masked_edited,
                location,
                word_prefix,
                metadata,
                diagnostics,
                false,
            )?;
            restore_text_box_subtrees(&mut tracked, &preserved, word_prefix, &marker, location)?;
            return Ok(tracked);
        }
        if original_text_boxes.len() != edited_text_boxes.len() {
            return Err(Error::Other(format!(
                "comparison cannot change nested text-box hosts at {location}"
            )));
        }
        let text_box_coordinates = text_box_coordinates(original, &original_text_boxes)?;
        let mut tracked_boxes = Vec::with_capacity(original_text_boxes.len());
        for ((run_ordinal, box_ordinal), (left, right)) in text_box_coordinates
            .into_iter()
            .zip(original_text_boxes.iter().zip(&edited_text_boxes))
        {
            let left_box = &original[left.clone()];
            let right_box = &edited[right.clone()];
            let left_inner = element_inner_range_any_prefix(left_box, "txbxContent")?;
            let right_inner = element_inner_range_any_prefix(right_box, "txbxContent")?;
            let tracked_inner = compare_story_inner(
                &left_box[left_inner.clone()],
                &right_box[right_inner],
                &format!("{location}/text-box[{run_ordinal}:{box_ordinal}]"),
                word_prefix,
                metadata,
                diagnostics,
            )?;
            let mut tracked_box = left_box.to_owned();
            tracked_box.replace_range(left_inner, &tracked_inner);
            tracked_boxes.push(tracked_box);
        }
        let tracked_hosts = replace_text_boxes_in_hosts(
            original,
            &original_text_boxes,
            &tracked_boxes,
            word_prefix,
        )?;
        let marker = text_box_marker_local(original, edited, word_prefix)?;
        let (masked_original, _) = mask_text_box_subtrees(original, word_prefix, &marker)?;
        let (masked_edited, _) = mask_text_box_subtrees(edited, word_prefix, &marker)?;
        let mut tracked = compare_story_inner_impl(
            &masked_original,
            &masked_edited,
            location,
            word_prefix,
            metadata,
            diagnostics,
            false,
        )?;
        restore_text_box_subtrees(&mut tracked, &tracked_hosts, word_prefix, &marker, location)?;
        return Ok(tracked);
    }
    let original_spans = story_content_spans(original, word_prefix)?;
    let edited_spans = story_content_spans(edited, word_prefix)?;
    let original_document = story_document(original)?;
    let edited_document = story_document(edited)?;
    if original_spans.len() != original_document.body.content.len()
        || edited_spans.len() != edited_document.body.content.len()
    {
        return Err(Error::Other(format!(
            "comparison could not correlate related-story owners at {location}"
        )));
    }
    compare_body(
        &original_document,
        &edited_document,
        location,
        Some((original, &original_spans)),
        metadata,
        diagnostics,
    )
    .map_err(|error| Error::Other(format!("comparison failed in {location}: {error}")))
}

fn story_content_spans(xml: &str, word_prefix: &str) -> Result<Vec<Range<usize>>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open_elements = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_section = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes()
        );
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let is_section = is_section && element.local_name().as_ref() == b"sectPr";
                open_elements.push((depth == 1 && !is_section).then_some(before));
                depth += 1;
            }
            Event::Empty(element)
                if depth == 1 && !(is_section && element.local_name().as_ref() == b"sectPr") =>
            {
                spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
            }
            Event::Empty(_) => {}
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open_elements.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn text_box_coordinates(xml: &str, boxes: &[Range<usize>]) -> Result<Vec<(usize, usize)>> {
    let mut coordinates = Vec::with_capacity(boxes.len());
    let mut previous_run = None;
    let mut run_ordinal = 0usize;
    let mut box_ordinal = 0usize;
    for text_box in boxes {
        let run = containing_run(xml, text_box.start)?;
        if previous_run.as_ref() != Some(&run) {
            run_ordinal += usize::from(previous_run.is_some());
            box_ordinal = 0;
            previous_run = Some(run);
        }
        coordinates.push((run_ordinal, box_ordinal));
        box_ordinal += 1;
    }
    Ok(coordinates)
}

fn text_box_marker_local(original: &str, edited: &str, word_prefix: &str) -> Result<String> {
    for ordinal in 0usize.. {
        let local = format!("textBoxHost{ordinal}");
        if !contains_private_text_box_marker(original, word_prefix, &local)?
            && !contains_private_text_box_marker(edited, word_prefix, &local)?
        {
            return Ok(local);
        }
    }
    unreachable!()
}

fn contains_private_text_box_marker(xml: &str, word_prefix: &str, local: &str) -> Result<bool> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let private = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == b"urn:rdocx:comparison:private"
        );
        match event {
            Event::Start(element) | Event::Empty(element)
                if private && element.local_name().as_ref() == local.as_bytes() =>
            {
                return Ok(true);
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn mask_text_box_subtrees(
    xml: &str,
    word_prefix: &str,
    marker_local: &str,
) -> Result<(String, Vec<String>)> {
    let boxes = text_box_spans(xml, word_prefix)?;
    let spans = text_box_host_spans(xml, &boxes)?;
    let preserved = spans
        .iter()
        .map(|span| xml[span.clone()].to_owned())
        .collect::<Vec<_>>();
    let mut masked = xml.to_owned();
    for (index, span) in spans.into_iter().enumerate().rev() {
        masked.replace_range(
            span,
            &format!(
                r#"<rdocxcmp:{marker_local} xmlns:rdocxcmp="urn:rdocx:comparison:private" rdocxcmp:index="{index}"/>"#
            ),
        );
    }
    Ok((masked, preserved))
}

fn restore_text_box_subtrees(
    xml: &mut String,
    preserved: &[String],
    word_prefix: &str,
    marker_local: &str,
    location: &str,
) -> Result<()> {
    let spans = private_text_box_host_spans(xml, word_prefix, marker_local)?;
    if spans.len() != preserved.len() {
        return Err(Error::Other(format!(
            "comparison lost an ignored text-box subtree at {location}"
        )));
    }
    for (index, span) in spans.into_iter().enumerate().rev() {
        xml.replace_range(span, &preserved[index]);
    }
    Ok(())
}

fn text_box_host_spans(xml: &str, boxes: &[Range<usize>]) -> Result<Vec<Range<usize>>> {
    let mut hosts = Vec::new();
    for text_box in boxes {
        let run = containing_run(xml, text_box.start)?;
        let relative = text_box.start - run.0;
        let children = direct_element_spans(&xml[run.0..run.1])?;
        let child = children
            .into_iter()
            .find(|child| child.start <= relative && relative < child.end)
            .ok_or_else(|| Error::Other("text-box subtree has no run-child host".to_owned()))?;
        let host = run.0 + child.start..run.0 + child.end;
        if hosts.last() != Some(&host) {
            hosts.push(host);
        }
    }
    Ok(hosts)
}

fn text_box_host_skeletons(xml: &str, boxes: &[Range<usize>]) -> Result<Vec<String>> {
    text_box_host_spans(xml, boxes)?
        .into_iter()
        .map(|host| {
            let host = &xml[host];
            let boxes = text_box_spans(host, "w")?;
            Ok(story_skeleton(host, &boxes))
        })
        .collect()
}

fn replace_text_boxes_in_hosts(
    xml: &str,
    boxes: &[Range<usize>],
    replacements: &[String],
    word_prefix: &str,
) -> Result<Vec<String>> {
    let mut replacement_index = 0usize;
    let mut hosts = Vec::new();
    for span in text_box_host_spans(xml, boxes)? {
        let mut host = xml[span].to_owned();
        let host_boxes = text_box_spans(&host, word_prefix)?;
        let count = host_boxes.len();
        for (box_span, replacement) in host_boxes
            .into_iter()
            .zip(&replacements[replacement_index..replacement_index + count])
            .rev()
        {
            host.replace_range(box_span, replacement);
        }
        replacement_index += count;
        hosts.push(host);
    }
    if replacement_index != replacements.len() {
        return Err(Error::Other(
            "comparison could not correlate nested text-box owners".to_owned(),
        ));
    }
    Ok(hosts)
}

fn direct_element_spans(xml: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(_) => {
                open.push((depth == 1).then_some(before));
                depth += 1;
            }
            Event::Empty(_) if depth == 1 => spans.push(before..after),
            Event::Empty(_) => {}
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open.pop() {
                    spans.push(start..after);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn private_text_box_host_spans(
    xml: &str,
    word_prefix: &str,
    marker_local: &str,
) -> Result<Vec<Range<usize>>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut stack = Vec::<Option<usize>>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_marker = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == b"urn:rdocx:comparison:private"
        );
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => stack.push(
                (is_marker && element.local_name().as_ref() == marker_local.as_bytes())
                    .then_some(before),
            ),
            Event::Empty(element)
                if is_marker && element.local_name().as_ref() == marker_local.as_bytes() =>
            {
                spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
            }
            Event::Empty(_) => {}
            Event::End(_) => {
                if let Some(Some(start)) = stack.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn without_text_box_subtrees(xml: &[u8]) -> Result<Vec<u8>> {
    let source = std::str::from_utf8(xml).map_err(utf8_error)?;
    let boxes = text_box_spans(source, "w")?;
    if boxes.is_empty() {
        return Ok(xml.to_vec());
    }
    let mut output = source.to_owned();
    for span in boxes.into_iter().rev() {
        output.replace_range(span, "");
    }
    Ok(output.into_bytes())
}

fn modeled_revision_count_with_options(xml: &[u8], options: &ComparisonOptions) -> Result<usize> {
    if story_ignored(options, ComparisonStoryKind::TextBox) {
        crate::revision::modeled_revision_count(&without_text_box_subtrees(xml)?)
    } else {
        crate::revision::modeled_revision_count(xml)
    }
}

fn word_ids_with_options(xml: &[u8], options: &ComparisonOptions) -> Result<HashSet<i32>> {
    if story_ignored(options, ComparisonStoryKind::TextBox) {
        word_ids(&without_text_box_subtrees(xml)?)
    } else {
        word_ids(xml)
    }
}

fn text_box_spans(xml: &str, word_prefix: &str) -> Result<Vec<Range<usize>>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut stack = Vec::<Option<usize>>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_text_box =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                stack.push(
                    (is_text_box && element.local_name().as_ref() == b"txbxContent")
                        .then_some(before),
                );
            }
            Event::Empty(element)
                if is_text_box && element.local_name().as_ref() == b"txbxContent" =>
            {
                spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
            }
            Event::End(_) => {
                if let Some(Some(start)) = stack.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn root_prefix(xml: &str, local: &str) -> Result<String> {
    let suffix = format!(":{local}");
    let suffix_at = xml
        .find(&suffix)
        .ok_or_else(|| Error::Other(format!("comparison {local} root has no prefix")))?;
    let name_start = xml[..suffix_at]
        .rfind('<')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("comparison {local} root has no start")))?;
    Ok(xml[name_start..suffix_at].to_owned())
}

fn element_inner_range_any_prefix(xml: &str, local: &str) -> Result<Range<usize>> {
    let open_end = xml
        .find('>')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("{local} source has no start tag")))?;
    let name_end = xml[..open_end]
        .find(local)
        .map(|at| at + local.len())
        .ok_or_else(|| Error::Other(format!("{local} source has no owner name")))?;
    let name_start = xml[..name_end]
        .rfind('<')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("{local} source has no owner start")))?;
    let close = format!("</{}>", &xml[name_start..name_end]);
    let close_start = xml
        .rfind(&close)
        .ok_or_else(|| Error::Other(format!("{local} source has no end tag")))?;
    Ok(open_end..close_start)
}

fn story_document(inner: &str) -> Result<CT_Document> {
    let xml = format!(
        r#"<rdocxcmp:document xmlns:rdocxcmp="{W_NS}" xmlns:w="{W_NS}"><rdocxcmp:body>{inner}</rdocxcmp:body></rdocxcmp:document>"#
    );
    CT_Document::from_xml(xml.as_bytes()).map_err(Into::into)
}

fn root_inner_range(xml: &str, expected_local: &str) -> Result<Range<usize>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if element.local_name().as_ref() != expected_local.as_bytes() || !is_word {
                        return Err(Error::Other(format!(
                            "comparison expected a Word {expected_local} root"
                        )));
                    }
                    start = Some(after);
                }
                depth += 1;
            }
            Event::Empty(element) if depth == 0 => {
                if element.local_name().as_ref() == expected_local.as_bytes() && is_word {
                    return Ok(after..after);
                }
                return Err(Error::Other(format!(
                    "comparison expected a Word {expected_local} root"
                )));
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok(start.unwrap_or(before)..before);
                }
            }
            Event::Eof => {
                return Err(Error::Other(format!(
                    "comparison XML has no closed {expected_local} root"
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn direct_word_element_spans_prefix_aware(xml: &str, local: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let target =
                    depth == 1 && element.local_name().as_ref() == local.as_bytes() && is_word;
                open.push(target.then_some(before));
                depth += 1;
            }
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == local.as_bytes() && is_word {
                    spans.push(before..after);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open.pop() {
                    spans.push(start..after);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn story_skeleton(xml: &str, spans: &[Range<usize>]) -> String {
    let mut skeleton = xml.to_owned();
    for span in spans.iter().rev() {
        skeleton.replace_range(span.clone(), "<owner/>");
    }
    skeleton
}

fn owner_start_signature(xml: &str) -> Result<String> {
    let end = xml
        .find('>')
        .ok_or_else(|| Error::Other("comparison owner has no start tag".to_owned()))?;
    Ok(xml[..=end].to_owned())
}

fn normal_note_owner(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (_, event) = reader
        .read_resolved_event_into(&mut buffer)
        .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
    let (Event::Start(element) | Event::Empty(element)) = event else {
        return Err(Error::Other(
            "comparison note has no owner element".to_owned(),
        ));
    };
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == b"type"
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
        {
            return Ok(attribute.value.as_ref().is_empty());
        }
    }
    Ok(true)
}

pub(crate) fn reopen_staged(candidate: Document) -> Result<Document> {
    let mut bytes = std::io::Cursor::new(Vec::new());
    candidate.package.write_to(&mut bytes)?;
    Document::from_bytes(bytes.get_ref())
}

type NormalizedPackage = (Vec<String>, Vec<(ComparisonStoryKind, String, Vec<String>)>);

fn comparison_text_box_markers(
    original: &Document,
    edited: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
) -> Result<TextBoxMarkers> {
    if !story_ignored(options, ComparisonStoryKind::TextBox) {
        return Ok(TextBoxMarkers::default());
    }
    let original_main = original
        .package
        .get_part(&original.doc_part_name)
        .ok_or_else(|| Error::Other(format!("missing main story {}", original.doc_part_name)))?;
    let edited_main = edited
        .package
        .get_part(&edited.doc_part_name)
        .ok_or_else(|| Error::Other(format!("missing main story {}", edited.doc_part_name)))?;
    let original_main = std::str::from_utf8(original_main).map_err(utf8_error)?;
    let edited_main = std::str::from_utf8(edited_main).map_err(utf8_error)?;
    let mut markers = TextBoxMarkers {
        main: text_box_marker_local(original_main, edited_main, "w")?,
        related: HashMap::new(),
    };
    for story in stories {
        let original_xml = std::str::from_utf8(story_xml(original, story)?).map_err(utf8_error)?;
        let edited_xml = std::str::from_utf8(story_xml(edited, story)?).map_err(utf8_error)?;
        let prefix = root_prefix(original_xml, story.kind.root_local())?;
        markers.related.insert(
            story.part_name.clone(),
            text_box_marker_local(original_xml, edited_xml, &prefix)?,
        );
    }
    Ok(markers)
}

fn normalized_package(
    document: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
    text_box_markers: &TextBoxMarkers,
) -> Result<NormalizedPackage> {
    let mut related = Vec::with_capacity(stories.len());
    for story in stories {
        related.push((
            story.kind,
            story.part_name.clone(),
            normalized_story_part(
                story_xml(document, story)?,
                story.kind,
                options,
                text_box_markers
                    .related
                    .get(&story.part_name)
                    .map(String::as_str),
            )?,
        ));
    }
    let main = if story_ignored(options, ComparisonStoryKind::TextBox) {
        let source = document
            .package
            .get_part(&document.doc_part_name)
            .ok_or_else(|| {
                Error::Other(format!("missing main story {}", document.doc_part_name))
            })?;
        let source = std::str::from_utf8(source).map_err(utf8_error)?;
        let (masked, _) = mask_text_box_subtrees(source, "w", &text_box_markers.main)?;
        normalized_body_with_options(&CT_Document::from_xml(masked.as_bytes())?, options)
    } else {
        normalized_body_with_options(&document.document, options)
    };
    Ok((main, related))
}

fn normalized_story_part(
    xml: &[u8],
    kind: ComparisonStoryKind,
    options: &ComparisonOptions,
    text_box_marker: Option<&str>,
) -> Result<Vec<String>> {
    let masked;
    let xml = if story_ignored(options, ComparisonStoryKind::TextBox) {
        let source = std::str::from_utf8(xml).map_err(utf8_error)?;
        let prefix = root_prefix(source, kind.root_local())?;
        let marker = text_box_marker
            .ok_or_else(|| Error::Other("missing shared comparison text-box marker".to_owned()))?;
        masked = mask_text_box_subtrees(source, &prefix, marker)?.0;
        &masked
    } else {
        std::str::from_utf8(xml).map_err(utf8_error)?
    };
    let root = root_inner_range(xml, kind.root_local())?;
    if let Some(owner_local) = kind.owner_local() {
        let mut normalized = Vec::new();
        for owner in direct_word_element_spans_prefix_aware(xml, owner_local)? {
            let owner_xml = &xml[owner];
            if matches!(
                kind,
                ComparisonStoryKind::Footnote | ComparisonStoryKind::Endnote
            ) && !normal_note_owner(owner_xml)?
            {
                continue;
            }
            let inner = element_inner_range_any_prefix(owner_xml, owner_local)?;
            let document = story_document(&owner_xml[inner])?;
            normalized.extend(normalized_body_with_options(&document, options));
        }
        Ok(normalized)
    } else {
        Ok(normalized_body_with_options(
            &story_document(&xml[root])?,
            options,
        ))
    }
}

fn compare_body(
    original: &CT_Document,
    edited: &CT_Document,
    location: &str,
    original_source: Option<(&str, &[Range<usize>])>,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_signatures = original
        .body
        .content
        .iter()
        .map(|content| body_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .body
        .content
        .iter()
        .map(|content| body_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    let aligned = expand_body_alignment(
        align(&original_signatures, &edited_signatures),
        &original.body.content,
        &edited.body.content,
    );
    let moves = pair_moves(
        &aligned,
        &original_signatures,
        &edited_signatures,
        &mut metadata.ids,
    )?;
    let mut output: Vec<(bool, String)> = Vec::new();
    for (position, (original_index, edited_index)) in aligned.iter().copied().enumerate() {
        let next_is_paragraph = aligned.get(position + 1).is_some_and(|(left, right)| {
            right
                .and_then(|index| edited.body.content.get(index))
                .or_else(|| left.and_then(|index| original.body.content.get(index)))
                .is_some_and(|content| matches!(content, BodyContent::Paragraph(_)))
        });
        match (original_index, edited_index) {
            (Some(left), Some(right)) => output.push((
                matches!(edited.body.content[right], BodyContent::Paragraph(_)),
                compare_body_content(
                    &original.body.content[left],
                    &edited.body.content[right],
                    &body_location(location, &edited.body.content[right], right),
                    metadata,
                    diagnostics,
                )?,
            )),
            (Some(left), None) => {
                let content = &original.body.content[left];
                if let Some(id) = moves.original[left] {
                    if matches!(content, BodyContent::Paragraph(_)) && !next_is_paragraph {
                        mark_previous_paragraph_with_id(&mut output, "moveFrom", id, metadata)?;
                        output.push((
                            true,
                            moved_paragraph_content(content, "moveFrom", id, metadata)?,
                        ));
                    } else {
                        output.push((
                            matches!(content, BodyContent::Paragraph(_)),
                            moved_body_content(content, "moveFrom", id, metadata)?,
                        ));
                    }
                } else if matches!(content, BodyContent::Paragraph(_)) && !next_is_paragraph {
                    mark_previous_paragraph(&mut output, "del", metadata)?;
                    output.push((true, deleted_paragraph_content(content, metadata)?));
                } else {
                    output.push((
                        matches!(content, BodyContent::Paragraph(_)),
                        deleted_body_content(content, metadata)?,
                    ));
                }
            }
            (None, Some(right)) => {
                let content = &edited.body.content[right];
                if let Some(id) = moves.edited[right] {
                    if matches!(content, BodyContent::Paragraph(_)) && !next_is_paragraph {
                        mark_previous_paragraph_with_id(&mut output, "moveTo", id, metadata)?;
                        output.push((
                            true,
                            moved_paragraph_content(content, "moveTo", id, metadata)?,
                        ));
                    } else {
                        output.push((
                            matches!(content, BodyContent::Paragraph(_)),
                            moved_body_content(content, "moveTo", id, metadata)?,
                        ));
                    }
                } else if matches!(content, BodyContent::Paragraph(_)) && !next_is_paragraph {
                    mark_previous_paragraph(&mut output, "ins", metadata)?;
                    output.push((true, inserted_paragraph_content(content, metadata)?));
                } else {
                    output.push((
                        matches!(content, BodyContent::Paragraph(_)),
                        inserted_body_content(content, metadata)?,
                    ));
                }
            }
            (None, None) => unreachable!(),
        }
    }
    let mut output = if let Some((source, spans)) = original_source {
        interleave_story_source(source, spans, &aligned, output)?
    } else {
        output.into_iter().map(|(_, xml)| xml).collect::<String>()
    };
    output.push_str(&section_properties_xml(
        original.body.sect_pr.as_ref(),
        edited.body.sect_pr.as_ref(),
        "section",
        metadata,
        diagnostics,
    )?);
    Ok(output)
}

fn interleave_story_source(
    source: &str,
    spans: &[Range<usize>],
    aligned: &[(Option<usize>, Option<usize>)],
    output: Vec<(bool, String)>,
) -> Result<String> {
    if spans.len() != aligned.iter().filter(|(left, _)| left.is_some()).count()
        || aligned.len() != output.len()
    {
        return Err(Error::Other(
            "comparison source spans do not match related-story owners".to_owned(),
        ));
    }
    let mut tracked = String::new();
    if let Some(first) = spans.first() {
        tracked.push_str(&source[..first.start]);
    } else {
        tracked.push_str(source);
    }
    let mut next_original = 0usize;
    for ((left, _), (_, xml)) in aligned.iter().zip(output) {
        if let Some(index) = left {
            if *index != next_original {
                return Err(Error::Other(
                    "comparison related-story owner order is not monotonic".to_owned(),
                ));
            }
            if *index > 0 {
                tracked.push_str(&source[spans[index - 1].end..spans[*index].start]);
            }
            next_original += 1;
        }
        tracked.push_str(&xml);
    }
    if let Some(last) = spans.last() {
        tracked.push_str(&source[last.end..]);
    }
    Ok(tracked)
}

fn moved_body_content(
    content: &BodyContent,
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => moved_paragraph(paragraph, kind, id, metadata),
        BodyContent::Table(table) => moved_table(table, kind, id, metadata),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Err(Error::Other(
            "comparison cannot move a content-control or opaque body node".to_owned(),
        )),
    }
}

fn moved_paragraph_content(
    content: &BodyContent,
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<String> {
    let BodyContent::Paragraph(paragraph) = content else {
        unreachable!("caller checked paragraph content")
    };
    let mut output = String::from("<w:p>");
    if let Some(properties) = &paragraph.properties {
        output.push_str(&property_xml(properties)?);
    }
    for run in &paragraph.runs {
        output.push_str(&IdAllocator::revision_with_id(
            kind,
            metadata.author,
            metadata.timestamp,
            &run_xml(run)?,
            id,
        ));
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn moved_paragraph(
    paragraph: &CT_P,
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<String> {
    let marker = IdAllocator::marker_with_id(kind, metadata.author, metadata.timestamp, id);
    let mut properties = paragraph.properties.clone().unwrap_or_default();
    properties.rpr = Some(properties.rpr.take().unwrap_or_default());
    let mut properties = property_xml(&properties)?;
    let run_properties = direct_word_element_spans(&properties, "rPr")?;
    if let Some(span) = run_properties.first() {
        let updated = append_word_child(&properties[span.clone()], "rPr", &marker)?;
        properties.replace_range(span.clone(), &updated);
    } else {
        properties = append_word_child(&properties, "pPr", &format!("<w:rPr>{marker}</w:rPr>"))?;
    }
    let mut output = format!("<w:p>{properties}");
    for run in &paragraph.runs {
        output.push_str(&IdAllocator::revision_with_id(
            kind,
            metadata.author,
            metadata.timestamp,
            &run_xml(run)?,
            id,
        ));
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn moved_table(table: &CT_Tbl, kind: &str, id: i32, metadata: &Metadata<'_>) -> Result<String> {
    let source = table_xml(table)?;
    let replacements = table
        .rows
        .iter()
        .map(|row| moved_row(row, kind, id, metadata))
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(&source, "tr", &replacements)
}

fn moved_row(row: &CT_Row, kind: &str, id: i32, metadata: &Metadata<'_>) -> Result<String> {
    let marker = IdAllocator::marker_with_id(kind, metadata.author, metadata.timestamp, id);
    let mut xml = row_xml(row)?;
    let properties = direct_word_element_spans(&xml, "trPr")?;
    if let Some(span) = properties.first() {
        let updated = append_word_child(&xml[span.clone()], "trPr", &marker)?;
        xml.replace_range(span.clone(), &updated);
        Ok(xml)
    } else {
        let open = xml
            .find('>')
            .ok_or_else(|| Error::Other("row XML has no start".to_owned()))?
            + 1;
        Ok(format!(
            "{}<w:trPr>{marker}</w:trPr>{}",
            &xml[..open],
            &xml[open..]
        ))
    }
}

fn mark_previous_paragraph(
    output: &mut [(bool, String)],
    kind: &str,
    metadata: &mut Metadata<'_>,
) -> Result<()> {
    let Some((true, paragraph)) = output.last_mut() else {
        return Err(Error::Other(
            "comparison needs an adjacent paragraph for a final paragraph change".to_owned(),
        ));
    };
    let marker = metadata
        .ids
        .marker(kind, metadata.author, metadata.timestamp)?;
    let paragraph_properties = direct_word_element_spans(paragraph, "pPr")?;
    if let Some(properties_span) = paragraph_properties.first() {
        let properties = &paragraph[properties_span.clone()];
        let run_properties = direct_word_element_spans(properties, "rPr")?;
        let updated = if let Some(run_span) = run_properties.first() {
            let run_properties = &properties[run_span.clone()];
            let updated_run = append_word_child(run_properties, "rPr", &marker)?;
            let mut updated = properties.to_owned();
            updated.replace_range(run_span.clone(), &updated_run);
            updated
        } else {
            append_word_child(properties, "pPr", &format!("<w:rPr>{marker}</w:rPr>"))?
        };
        paragraph.replace_range(properties_span.clone(), &updated);
    } else {
        let open = paragraph
            .find('>')
            .ok_or_else(|| Error::Other("paragraph XML has no start".to_owned()))?
            + 1;
        *paragraph = format!(
            "{}<w:pPr><w:rPr>{marker}</w:rPr></w:pPr>{}",
            &paragraph[..open],
            &paragraph[open..]
        );
    }
    Ok(())
}

fn mark_previous_paragraph_with_id(
    output: &mut [(bool, String)],
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<()> {
    let Some((true, paragraph)) = output.last_mut() else {
        return Err(Error::Other(
            "comparison needs an adjacent paragraph for a final paragraph move".to_owned(),
        ));
    };
    let marker = IdAllocator::marker_with_id(kind, metadata.author, metadata.timestamp, id);
    let paragraph_properties = direct_word_element_spans(paragraph, "pPr")?;
    if let Some(properties_span) = paragraph_properties.first() {
        let properties = &paragraph[properties_span.clone()];
        let run_properties = direct_word_element_spans(properties, "rPr")?;
        let updated = if let Some(run_span) = run_properties.first() {
            let updated_run = append_word_child(&properties[run_span.clone()], "rPr", &marker)?;
            let mut updated = properties.to_owned();
            updated.replace_range(run_span.clone(), &updated_run);
            updated
        } else {
            append_word_child(properties, "pPr", &format!("<w:rPr>{marker}</w:rPr>"))?
        };
        paragraph.replace_range(properties_span.clone(), &updated);
    } else {
        let open = paragraph
            .find('>')
            .ok_or_else(|| Error::Other("paragraph XML has no start".to_owned()))?
            + 1;
        *paragraph = format!(
            "{}<w:pPr><w:rPr>{marker}</w:rPr></w:pPr>{}",
            &paragraph[..open],
            &paragraph[open..]
        );
    }
    Ok(())
}

fn deleted_paragraph_content(content: &BodyContent, metadata: &mut Metadata<'_>) -> Result<String> {
    let BodyContent::Paragraph(paragraph) = content else {
        unreachable!("caller checked paragraph content")
    };
    let mut output = String::from("<w:p>");
    if let Some(properties) = &paragraph.properties {
        output.push_str(&property_xml(properties)?);
    }
    for run in &paragraph.runs {
        let run = deleted_run_xml(run)?;
        output.push_str(&metadata.ids.revision(
            "del",
            metadata.author,
            metadata.timestamp,
            &run,
        )?);
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn inserted_paragraph_content(
    content: &BodyContent,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let BodyContent::Paragraph(paragraph) = content else {
        unreachable!("caller checked paragraph content")
    };
    let mut output = String::from("<w:p>");
    if let Some(properties) = &paragraph.properties {
        output.push_str(&property_xml(properties)?);
    }
    for run in &paragraph.runs {
        let run = run_xml(run)?;
        output.push_str(&metadata.ids.revision(
            "ins",
            metadata.author,
            metadata.timestamp,
            &run,
        )?);
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn compare_body_content(
    original: &BodyContent,
    edited: &BodyContent,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    match (original, edited) {
        (BodyContent::Paragraph(left), BodyContent::Paragraph(right)) => {
            compare_paragraph(left, right, location, metadata, diagnostics)
        }
        (BodyContent::Table(left), BodyContent::Table(right)) => {
            compare_table(left, right, location, metadata, diagnostics)
        }
        (BodyContent::ContentControl(left), BodyContent::ContentControl(right)) => {
            compare_control(left, right, location, metadata, diagnostics)
        }
        _ if body_signature(original) == body_signature(edited) => body_content_xml(original),
        _ => Err(Error::Other(format!(
            "comparison cannot replace unlike body structures at {location}"
        ))),
    }
}

fn deleted_body_content(content: &BodyContent, metadata: &mut Metadata<'_>) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => deleted_paragraph(paragraph, metadata),
        BodyContent::Table(table) => marked_table(table, "del", metadata),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Err(Error::Other(
            "comparison cannot delete an unmatched content-control or opaque body node".to_owned(),
        )),
    }
}

fn inserted_body_content(content: &BodyContent, metadata: &mut Metadata<'_>) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => inserted_paragraph(paragraph, metadata),
        BodyContent::Table(table) => marked_table(table, "ins", metadata),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Err(Error::Other(
            "comparison cannot insert an unmatched content-control or opaque body node".to_owned(),
        )),
    }
}

fn compare_paragraph(
    original: &CT_P,
    edited: &CT_P,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if uses_attributed_run_path(metadata.options) {
        return compare_granular_paragraph(original, edited, location, metadata, diagnostics);
    }
    if !original.hyperlinks.is_empty()
        || !edited.hyperlinks.is_empty()
        || !original.comment_ranges.is_empty()
        || !edited.comment_ranges.is_empty()
        || !original.bookmark_markers.is_empty()
        || !edited.bookmark_markers.is_empty()
        || !original.content_controls.is_empty()
        || !edited.content_controls.is_empty()
        || !original.extra_xml.is_empty()
        || !edited.extra_xml.is_empty()
    {
        return compare_complex_paragraph(original, edited, location, metadata, diagnostics);
    }

    let mut output = String::from("<w:p>");
    output.push_str(&paragraph_properties_xml(
        original,
        edited,
        location,
        metadata,
        diagnostics,
    )?);
    let original_signatures = original
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let aligned = align(&original_signatures, &edited_signatures);
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original.runs,
            &edited.runs,
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    for (left, right) in aligned {
        match (left, right) {
            (Some(i), Some(j)) => {
                if original_signatures[i] == edited_signatures[j] {
                    output.push_str(&compared_run_xml(
                        &original.runs[i],
                        &edited.runs[j],
                        &format!("{location}/run[{j}]"),
                        metadata,
                        diagnostics,
                    )?);
                } else {
                    let deleted = deleted_run_xml(&original.runs[i])?;
                    output.push_str(&metadata.ids.revision(
                        "del",
                        metadata.author,
                        metadata.timestamp,
                        &deleted,
                    )?);
                    let inserted = policy_inserted_run_xml(
                        &edited.runs[j],
                        Some(&original.runs[i]),
                        metadata.options,
                    )?;
                    output.push_str(&metadata.ids.revision(
                        "ins",
                        metadata.author,
                        metadata.timestamp,
                        &inserted,
                    )?);
                }
            }
            (Some(i), None) => {
                if run_is_ignored(&original.runs[i], metadata.options) {
                    output.push_str(&paragraph_owned_run_xml(&original.runs[i])?);
                    continue;
                }
                let run = deleted_run_xml(&original.runs[i])?;
                output.push_str(&metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &run,
                )?);
            }
            (None, Some(j)) => {
                if run_is_ignored(&edited.runs[j], metadata.options) {
                    continue;
                }
                let run = paragraph_owned_run_xml(&edited.runs[j])?;
                output.push_str(&metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &run,
                )?);
            }
            (None, None) => unreachable!(),
        }
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn uses_attributed_run_path(options: &ComparisonOptions) -> bool {
    options.granularity != ComparisonGranularity::Run
        || options.ignore_formatting
        || options.ignore_whitespace
        || options.ignore_fields
        || options.ignore_comments
        || story_ignored(options, ComparisonStoryKind::TextBox)
}

#[derive(Clone)]
struct AttributedRunUnit {
    run: CT_R,
    ignored: bool,
    owner: usize,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum GranularAction {
    Preserve,
    Equal,
    Replace,
    Delete,
    Insert,
    Drop,
}

fn compare_granular_paragraph(
    original: &CT_P,
    edited: &CT_P,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if hyperlink_shells(original, metadata.options) != hyperlink_shells(edited, metadata.options)
        || (!metadata.options.ignore_comments && original.comment_ranges != edited.comment_ranges)
        || original.bookmark_markers != edited.bookmark_markers
        || paragraph_control_boundaries(original) != paragraph_control_boundaries(edited)
    {
        return Err(Error::Other(format!(
            "comparison cannot revise paragraph boundary structures at {location}"
        )));
    }

    let original_run_signatures = original.runs.iter().map(run_signature).collect::<Vec<_>>();
    let edited_run_signatures = edited.runs.iter().map(run_signature).collect::<Vec<_>>();
    if original_run_signatures == edited_run_signatures
        && original.content_controls == edited.content_controls
    {
        let properties =
            paragraph_properties_xml(original, edited, location, metadata, diagnostics)?;
        let replacements = original
            .runs
            .iter()
            .zip(&edited.runs)
            .enumerate()
            .map(|(index, (left, right))| {
                compared_run_xml(
                    left,
                    right,
                    &format!("{location}/run[{index}]"),
                    metadata,
                    diagnostics,
                )
            })
            .collect::<Result<Vec<_>>>()?;
        return replace_paragraph_properties_and_runs(original, &properties, &replacements);
    }

    let original_units = attributed_run_units(&original.runs, metadata.options);
    let edited_units = attributed_run_units(&edited.runs, metadata.options);
    let original_signatures = original_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();
    let edited_signatures = edited_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();

    let properties = paragraph_properties_xml(original, edited, location, metadata, diagnostics)?;
    let aligned = align(&original_signatures, &edited_signatures);
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &edited_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let grouped = coalesced_granular_alignment(
        &aligned,
        &original_units,
        &edited_units,
        &original_signatures,
        &edited_signatures,
    );
    let mut grouped_alignment = Vec::with_capacity(grouped.len());
    let mut replacements = Vec::with_capacity(grouped.len());
    for (action, members) in grouped {
        let first = aligned[members.start];
        grouped_alignment.push(first);
        let left_indices = members
            .clone()
            .filter_map(|index| aligned[index].0)
            .collect::<Vec<_>>();
        let right_indices = members
            .filter_map(|index| aligned[index].1)
            .collect::<Vec<_>>();
        let left = merge_unit_runs(&original_units, &left_indices);
        let right = merge_unit_runs(&edited_units, &right_indices);
        replacements.push(match action {
            GranularAction::Preserve => paragraph_owned_run_xml(left.as_ref().unwrap())?,
            GranularAction::Equal => compared_run_xml(
                left.as_ref().unwrap(),
                right.as_ref().unwrap(),
                &format!("{location}/run-unit[{}]", right_indices[0]),
                metadata,
                diagnostics,
            )?,
            GranularAction::Replace => {
                let left = left.as_ref().unwrap();
                let deleted = metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted_run_xml(left)?,
                )?;
                let inserted = metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &policy_inserted_run_xml(
                        right.as_ref().unwrap(),
                        Some(left),
                        metadata.options,
                    )?,
                )?;
                format!("{deleted}{inserted}")
            }
            GranularAction::Delete => metadata.ids.revision(
                "del",
                metadata.author,
                metadata.timestamp,
                &deleted_run_xml(left.as_ref().unwrap())?,
            )?,
            GranularAction::Insert => metadata.ids.revision(
                "ins",
                metadata.author,
                metadata.timestamp,
                &paragraph_owned_run_xml(right.as_ref().unwrap())?,
            )?,
            GranularAction::Drop => String::new(),
        });
    }
    interleave_granular_paragraph(
        original,
        edited,
        &original_units,
        &grouped_alignment,
        &replacements,
        &properties,
        location,
        metadata,
        diagnostics,
    )
}

fn hyperlink_shells(paragraph: &CT_P, options: &ComparisonOptions) -> Vec<String> {
    paragraph
        .hyperlinks
        .iter()
        .map(|link| {
            format!(
                "{:?}:{:?}:{:?}:{:?}:{:?}:{:?}:{:?}:{}:{}",
                link.rel_id,
                link.anchor,
                link.tooltip,
                link.doc_location,
                link.extra_attributes,
                hyperlink_raw_boundaries(paragraph, link, options),
                link.preserved_raw_before,
                policy_run_boundary(paragraph, link.run_start, options),
                policy_run_boundary(paragraph, link.run_end, options)
            )
        })
        .collect()
}

fn hyperlink_raw_boundaries(
    paragraph: &CT_P,
    link: &rdocx_oxml::text::HyperlinkSpan,
    options: &ComparisonOptions,
) -> Vec<(usize, usize, Vec<u8>)> {
    let logical_start = policy_run_boundary(paragraph, link.run_start, options);
    link.extra_xml
        .iter()
        .map(|(boundary, revisions_before, raw)| {
            let absolute = link.run_start.saturating_add(*boundary);
            (
                policy_run_boundary(paragraph, absolute, options).saturating_sub(logical_start),
                *revisions_before,
                raw.clone(),
            )
        })
        .collect()
}

fn policy_run_boundary(
    paragraph: &CT_P,
    run_boundary: usize,
    options: &ComparisonOptions,
) -> usize {
    attributed_run_units(
        &paragraph.runs[..run_boundary.min(paragraph.runs.len())],
        options,
    )
    .iter()
    .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
    .count()
}

fn unit_is_ignorable(unit: &AttributedRunUnit) -> bool {
    unit.ignored && unit.run.extra_xml.is_empty() && unit.run.alt_drawings.is_empty()
}

fn unit_is_empty(unit: &AttributedRunUnit) -> bool {
    unit.run.content.is_empty()
        && unit.run.extra_xml.is_empty()
        && unit.run.alt_drawings.is_empty()
        && unit.run.properties.is_none()
}

fn granular_action(
    pair: (Option<usize>, Option<usize>),
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    original_signatures: &[String],
    edited_signatures: &[String],
) -> GranularAction {
    match pair {
        (Some(i), Some(j))
            if original_units[i].ignored
                && edited_units[j].ignored
                && original_signatures[i] == edited_signatures[j] =>
        {
            GranularAction::Preserve
        }
        (Some(i), Some(j)) if original_signatures[i] == edited_signatures[j] => {
            GranularAction::Equal
        }
        (Some(_), Some(_)) => GranularAction::Replace,
        (Some(i), None) if unit_is_ignorable(&original_units[i]) => GranularAction::Preserve,
        (None, Some(j)) if unit_is_ignorable(&edited_units[j]) => GranularAction::Drop,
        (Some(_), None) => GranularAction::Delete,
        (None, Some(_)) => GranularAction::Insert,
        (None, None) => unreachable!(),
    }
}

fn coalesced_granular_alignment(
    aligned: &[(Option<usize>, Option<usize>)],
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    original_signatures: &[String],
    edited_signatures: &[String],
) -> Vec<(GranularAction, Range<usize>)> {
    let mut groups = Vec::new();
    let mut start = 0usize;
    while start < aligned.len() {
        let action = granular_action(
            aligned[start],
            original_units,
            edited_units,
            original_signatures,
            edited_signatures,
        );
        let owners = (
            aligned[start].0.map(|index| original_units[index].owner),
            aligned[start].1.map(|index| edited_units[index].owner),
        );
        let mut end = start + 1;
        if matches!(
            action,
            GranularAction::Replace | GranularAction::Delete | GranularAction::Insert
        ) {
            while end < aligned.len()
                && granular_action(
                    aligned[end],
                    original_units,
                    edited_units,
                    original_signatures,
                    edited_signatures,
                ) == action
                && (
                    aligned[end].0.map(|index| original_units[index].owner),
                    aligned[end].1.map(|index| edited_units[index].owner),
                ) == owners
            {
                end += 1;
            }
        }
        groups.push((action, start..end));
        start = end;
    }
    groups
}

fn merge_unit_runs(units: &[AttributedRunUnit], indices: &[usize]) -> Option<CT_R> {
    let mut merged = units.get(*indices.first()?)?.run.clone();
    let property_boundary = usize::from(merged.properties.is_some());
    for &index in &indices[1..] {
        let next = &units[index].run;
        let content_offset = merged.content.len();
        merged.content.extend(next.content.iter().cloned());
        merged
            .alt_drawings
            .extend(next.alt_drawings.iter().cloned());
        for (raw, &encoded) in next.extra_xml.iter().zip(&next.extra_xml_positions) {
            let boundary = CT_R::raw_child_position(encoded);
            let mut rebased = encoded;
            CT_R::set_raw_child_position(&mut rebased, boundary + content_offset);
            merged.extra_xml.push(raw.clone());
            merged.extra_xml_positions.push(rebased);
        }
        if property_boundary == 0 && next.properties.is_some() {
            merged.properties = next.properties.clone();
        }
    }
    Some(merged)
}

fn replace_paragraph_properties_and_runs(
    paragraph: &CT_P,
    properties: &str,
    runs: &[String],
) -> Result<String> {
    let mut source = paragraph_xml(paragraph)?;
    let property_spans = direct_word_element_spans(&source, "pPr")?;
    match (property_spans.first(), properties.is_empty()) {
        (Some(span), false) => source.replace_range(span.clone(), properties),
        (Some(span), true) => source.replace_range(span.clone(), ""),
        (None, false) => {
            let open = source
                .find('>')
                .ok_or_else(|| Error::Other("paragraph XML has no start".to_owned()))?
                + 1;
            source.insert_str(open, properties);
        }
        (None, true) => {}
    }
    replace_paragraph_run_elements(&source, runs)
}

#[allow(clippy::too_many_arguments)]
fn interleave_granular_paragraph(
    original: &CT_P,
    edited: &CT_P,
    original_units: &[AttributedRunUnit],
    aligned: &[(Option<usize>, Option<usize>)],
    replacements: &[String],
    properties: &str,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let mut source = paragraph_xml(original)?;
    let property_spans = direct_word_element_spans(&source, "pPr")?;
    match (property_spans.first(), properties.is_empty()) {
        (Some(span), false) => source.replace_range(span.clone(), properties),
        (Some(span), true) => source.replace_range(span.clone(), ""),
        (None, false) => {
            let open = source
                .find('>')
                .ok_or_else(|| Error::Other("paragraph XML has no start".to_owned()))?
                + 1;
            source.insert_str(open, properties);
        }
        (None, true) => {}
    }
    let spans = paragraph_run_spans(&source)?;
    if spans.len() != original.runs.len() || aligned.len() != replacements.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate granular run owners at {location}"
        )));
    }
    let insertion_boundary = spans
        .first()
        .map_or_else(|| paragraph_close_start(&source), |span| Ok(span.start))?;
    let mut output = source[..insertion_boundary].to_owned();
    let mut cursor = insertion_boundary;
    let mut consumed_owner = None;
    for ((left, _), replacement) in aligned.iter().zip(replacements) {
        if let Some(unit) = left.map(|index| &original_units[index])
            && consumed_owner != Some(unit.owner)
        {
            let span = &spans[unit.owner];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
            consumed_owner = Some(unit.owner);
        }
        output.push_str(replacement);
    }
    output.push_str(&source[cursor..]);

    let control_spans = direct_word_element_spans(&output, "sdt")?;
    if control_spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate granular content controls at {location}"
        )));
    }
    let mut control_replacements = Vec::with_capacity(control_spans.len());
    for (index, ((_, _, _, left), (_, _, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        control_replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[control_spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &control_replacements)
}

fn attributed_run_units(runs: &[CT_R], options: &ComparisonOptions) -> Vec<AttributedRunUnit> {
    let mut units = Vec::new();
    for (owner, run) in runs.iter().enumerate() {
        let unit_start = units.len();
        let mut content_units = Vec::<Range<usize>>::with_capacity(run.content.len());
        for content in &run.content {
            let start = units.len() - unit_start;
            let fragments = match content {
                RunContent::Text(text) => granular_text(text, options)
                    .into_iter()
                    .map(RunContent::Text)
                    .collect::<Vec<_>>(),
                RunContent::DeletedText(text) => granular_text(text, options)
                    .into_iter()
                    .map(RunContent::DeletedText)
                    .collect::<Vec<_>>(),
                content => vec![content.clone()],
            };
            for fragment in fragments {
                let ignored = ignored_run_content(&fragment, options);
                units.push(AttributedRunUnit {
                    run: CT_R {
                        properties: run.properties.clone(),
                        content: vec![fragment],
                        extra_xml: Vec::new(),
                        extra_xml_positions: Vec::new(),
                        alt_drawings: Vec::new(),
                    },
                    ignored,
                    owner,
                });
            }
            content_units.push(start..units.len() - unit_start);
        }
        if units.len() == unit_start {
            units.push(AttributedRunUnit {
                run: CT_R {
                    properties: run.properties.clone(),
                    content: Vec::new(),
                    extra_xml: Vec::new(),
                    extra_xml_positions: Vec::new(),
                    alt_drawings: run.alt_drawings.clone(),
                },
                ignored: false,
                owner,
            });
        } else {
            units[unit_start].run.alt_drawings = run.alt_drawings.clone();
        }
        attribute_raw_children(run, &content_units, &mut units[unit_start..]);
        let property_boundary = usize::from(run.properties.is_some());
        let mut inserted = 0usize;
        for (raw, &encoded) in run.extra_xml.iter().zip(&run.extra_xml_positions) {
            if !is_text_box_host_marker(raw) {
                continue;
            }
            let boundary = CT_R::raw_child_position(encoded);
            let offset = if boundary <= property_boundary || content_units.is_empty() {
                0
            } else {
                let content_index = (boundary - property_boundary - 1).min(content_units.len() - 1);
                content_units[content_index].end
            };
            let mut position = encoded;
            CT_R::set_raw_child_position(&mut position, property_boundary);
            units.insert(
                unit_start + offset + inserted,
                AttributedRunUnit {
                    run: CT_R {
                        properties: run.properties.clone(),
                        content: Vec::new(),
                        extra_xml: vec![raw.clone()],
                        extra_xml_positions: vec![position],
                        alt_drawings: Vec::new(),
                    },
                    ignored: false,
                    owner,
                },
            );
            inserted += 1;
        }
    }
    units
}

fn attribute_raw_children(
    run: &CT_R,
    content_units: &[Range<usize>],
    units: &mut [AttributedRunUnit],
) {
    let property_boundary = usize::from(run.properties.is_some());
    for (raw, &encoded) in run.extra_xml.iter().zip(&run.extra_xml_positions) {
        if is_text_box_host_marker(raw) {
            continue;
        }
        let boundary = CT_R::raw_child_position(encoded);
        let (unit_index, position) = if boundary <= property_boundary || content_units.is_empty() {
            (0, boundary.min(property_boundary))
        } else {
            let content_index = (boundary - property_boundary - 1).min(content_units.len() - 1);
            let range = &content_units[content_index];
            (range.end.saturating_sub(1), property_boundary + 1)
        };
        units[unit_index].run.extra_xml.push(raw.clone());
        let mut attributed = encoded;
        CT_R::set_raw_child_position(&mut attributed, position);
        units[unit_index].run.extra_xml_positions.push(attributed);
    }
}

fn is_text_box_host_marker(raw: &[u8]) -> bool {
    raw.windows(b"urn:rdocx:comparison:private".len())
        .any(|window| window == b"urn:rdocx:comparison:private")
        && raw
            .windows(b"textBoxHost".len())
            .any(|window| window == b"textBoxHost")
}

fn granular_text(text: &CT_Text, options: &ComparisonOptions) -> Vec<CT_Text> {
    let fragments = match options.granularity {
        ComparisonGranularity::Run if options.ignore_whitespace => whitespace_fragments(&text.text),
        ComparisonGranularity::Run => vec![text.text.clone()],
        ComparisonGranularity::Character => text.text.chars().map(String::from).collect(),
        ComparisonGranularity::Word => word_fragments(&text.text),
    };
    if fragments.is_empty() {
        return vec![text.clone()];
    }
    fragments
        .into_iter()
        .map(|value| CT_Text {
            text: value,
            preserve_space: text.preserve_space,
        })
        .collect()
}

fn whitespace_fragments(text: &str) -> Vec<String> {
    let mut output = Vec::new();
    let mut current = String::new();
    let mut whitespace = None;
    for character in text.chars() {
        let next = character.is_whitespace();
        if whitespace.is_some_and(|value| value != next) {
            output.push(std::mem::take(&mut current));
        }
        current.push(character);
        whitespace = Some(next);
    }
    if !current.is_empty() {
        output.push(current);
    }
    output
}

fn word_fragments(text: &str) -> Vec<String> {
    fn class(character: char) -> u8 {
        if character.is_alphanumeric() || character == '_' {
            0
        } else if character.is_whitespace() {
            1
        } else {
            2
        }
    }
    let mut output = Vec::new();
    let mut current = String::new();
    let mut current_class = None;
    for character in text.chars() {
        let next_class = class(character);
        if current_class.is_some_and(|value| value != next_class) {
            output.push(std::mem::take(&mut current));
        }
        current.push(character);
        current_class = Some(next_class);
    }
    if !current.is_empty() {
        output.push(current);
    }
    output
}

fn ignored_run_content(content: &RunContent, options: &ComparisonOptions) -> bool {
    (options.ignore_fields && matches!(content, RunContent::Field(_)))
        || (options.ignore_comments && matches!(content, RunContent::CommentReference { .. }))
        || (options.ignore_whitespace
            && matches!(
                content,
                RunContent::Text(text) | RunContent::DeletedText(text)
                    if text.text.chars().all(char::is_whitespace)
            ))
}

fn run_is_ignored(run: &CT_R, options: &ComparisonOptions) -> bool {
    !run.content.is_empty()
        && run
            .content
            .iter()
            .all(|content| ignored_run_content(content, options))
        && run.extra_xml.is_empty()
}

fn attributed_unit_signature(unit: &AttributedRunUnit) -> String {
    if unit.ignored {
        let kind = match unit.run.content.first() {
            Some(RunContent::Field(_)) => "ignored:field".to_owned(),
            Some(RunContent::CommentReference { .. }) => "ignored:comment".to_owned(),
            _ => "ignored:whitespace".to_owned(),
        };
        format!(
            "{kind}:{:?}:{:?}:{:?}",
            unit.run.extra_xml, unit.run.extra_xml_positions, unit.run.alt_drawings
        )
    } else {
        match unit.run.content.first() {
            Some(RunContent::CommentReference { id, .. }) if unit.run.content.len() == 1 => {
                format!(
                    "comment:{id}:{:?}:{:?}:{:?}",
                    unit.run.properties, unit.run.extra_xml, unit.run.alt_drawings
                )
            }
            _ => run_signature(&unit.run),
        }
    }
}

fn compare_complex_paragraph(
    original: &CT_P,
    edited: &CT_P,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if paragraph_signature_with_options(original, metadata.options)
        == paragraph_signature_with_options(edited, metadata.options)
    {
        if !metadata.options.ignore_formatting
            && paragraph_formatting(original) != paragraph_formatting(edited)
        {
            formatting_diagnostic(diagnostics, location.to_owned());
        }
        for (index, (left, right)) in original.runs.iter().zip(&edited.runs).enumerate() {
            if !metadata.options.ignore_formatting && left.properties != right.properties {
                formatting_diagnostic(diagnostics, format!("{location}/run[{index}]"));
            }
        }
        for (index, ((_, _, _, left), (_, _, _, right))) in original
            .content_controls
            .iter()
            .zip(&edited.content_controls)
            .enumerate()
        {
            compare_control(
                left,
                right,
                &format!("{location}/content-control[{index}]"),
                metadata,
                diagnostics,
            )?;
        }
        return paragraph_xml(original);
    }
    if original.hyperlinks != edited.hyperlinks
        || (!metadata.options.ignore_comments && original.comment_ranges != edited.comment_ranges)
        || original.bookmark_markers != edited.bookmark_markers
        || original.extra_xml != edited.extra_xml
        || paragraph_control_boundaries(original) != paragraph_control_boundaries(edited)
        || (!metadata.options.ignore_formatting && original.properties != edited.properties)
    {
        return Err(Error::Other(format!(
            "comparison cannot revise paragraph boundary structures at {location}"
        )));
    }

    let source = paragraph_xml(original)?;
    let original_signatures = original
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let aligned = align(&original_signatures, &edited_signatures);
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original.runs,
            &edited.runs,
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let mut output = String::new();
    let mut cursor = 0usize;
    for (left, right) in aligned {
        let old_run = left
            .map(|index| paragraph_owned_run_xml(&original.runs[index]))
            .transpose()?;
        if let Some(old_run) = &old_run {
            let relative = source[cursor..].find(old_run).ok_or_else(|| {
                Error::Other(format!(
                    "comparison could not locate a serialized run at {location}"
                ))
            })?;
            let start = cursor + relative;
            output.push_str(&source[cursor..start]);
            cursor = start + old_run.len();
        }
        match (left, right) {
            (Some(i), Some(j)) if original_signatures[i] == edited_signatures[j] => {
                output.push_str(&compared_run_xml(
                    &original.runs[i],
                    &edited.runs[j],
                    &format!("{location}/run[{j}]"),
                    metadata,
                    diagnostics,
                )?);
            }
            (Some(i), Some(j)) => {
                let deleted = deleted_run_xml(&original.runs[i])?;
                output.push_str(&metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted,
                )?);
                let inserted = policy_inserted_run_xml(
                    &edited.runs[j],
                    Some(&original.runs[i]),
                    metadata.options,
                )?;
                output.push_str(&metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &inserted,
                )?);
            }
            (Some(i), None) => {
                if run_is_ignored(&original.runs[i], metadata.options) {
                    output.push_str(&paragraph_owned_run_xml(&original.runs[i])?);
                    continue;
                }
                let deleted = deleted_run_xml(&original.runs[i])?;
                output.push_str(&metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted,
                )?);
            }
            (None, Some(j)) => {
                if run_is_ignored(&edited.runs[j], metadata.options) {
                    continue;
                }
                let inserted = paragraph_owned_run_xml(&edited.runs[j])?;
                output.push_str(&metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &inserted,
                )?);
            }
            (None, None) => unreachable!(),
        }
    }
    output.push_str(&source[cursor..]);
    let spans = direct_word_element_spans(&output, "sdt")?;
    if spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate paragraph content controls at {location}"
        )));
    }
    let mut replacements = Vec::with_capacity(spans.len());
    for (index, ((_, _, _, left), (_, _, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &replacements)
}

fn paragraph_control_boundaries(paragraph: &CT_P) -> Vec<(usize, usize, usize)> {
    paragraph
        .content_controls
        .iter()
        .map(|(at, raw_before, markers_before, _)| (*at, *raw_before, *markers_before))
        .collect()
}

fn paragraph_properties_xml(
    original: &CT_P,
    edited: &CT_P,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_section = original
        .properties
        .as_ref()
        .and_then(|properties| properties.sect_pr.as_ref());
    let edited_section = edited
        .properties
        .as_ref()
        .and_then(|properties| properties.sect_pr.as_ref());
    let tracked_section = section_properties_xml(
        original_section,
        edited_section,
        &format!("{location}/section"),
        metadata,
        diagnostics,
    )?;
    if metadata.options.ignore_formatting {
        return original
            .properties
            .as_ref()
            .map(property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }

    let original_modeled = modeled_paragraph_properties(original.properties.as_ref());
    let edited_modeled = modeled_paragraph_properties(edited.properties.as_ref());
    let changed = original_modeled != edited_modeled;
    let mut current = if changed {
        edited.properties.clone().unwrap_or_default()
    } else {
        original.properties.clone().unwrap_or_default()
    };
    current.sect_pr = None;
    current.change = None;
    let (original_numbering_xml, original_revision_xml, original_revision_positions) = original
        .properties
        .as_ref()
        .map(|properties| {
            (
                properties.numbering_revision_xml.clone(),
                properties.revision_xml.clone(),
                properties.revision_xml_positions.clone(),
            )
        })
        .unwrap_or_default();
    if current.numbering_revision_xml != original_numbering_xml
        || current.revision_xml != original_revision_xml
        || current.revision_xml_positions != original_revision_positions
    {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    current.numbering_revision_xml = original_numbering_xml;
    current.revision_xml = original_revision_xml;
    current.revision_xml_positions = original_revision_positions;
    let needs_owner = changed || !tracked_section.is_empty() || original.properties.is_some();
    if !needs_owner {
        return Ok(String::new());
    }
    let mut current_xml = property_xml(&current)?;
    if current_xml.is_empty() {
        current_xml.push_str("<w:pPr></w:pPr>");
    }
    if !tracked_section.is_empty() {
        current_xml = inject_before_close(&current_xml, "</w:pPr>", &tracked_section)?;
    }
    if changed {
        let previous = original_modeled
            .as_ref()
            .map(property_xml)
            .transpose()?
            .filter(|xml| !xml.is_empty())
            .unwrap_or_else(|| "<w:pPr/>".to_owned());
        let change =
            metadata
                .ids
                .revision("pPrChange", metadata.author, metadata.timestamp, &previous)?;
        current_xml = inject_before_close(&current_xml, "</w:pPr>", &change)?;
    }
    Ok(current_xml)
}

fn modeled_paragraph_properties(properties: Option<&CT_PPr>) -> Option<CT_PPr> {
    properties.cloned().map(|mut properties| {
        properties.sect_pr = None;
        properties.numbering_revision = None;
        properties.numbering_revision_xml.clear();
        properties.change = None;
        properties.revision_xml.clear();
        properties.revision_xml_positions.clear();
        properties
    })
}

fn section_properties_xml(
    original: Option<&rdocx_oxml::document::CT_SectPr>,
    edited: Option<&rdocx_oxml::document::CT_SectPr>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let (Some(original), Some(edited)) = (original, edited) else {
        return match (original, edited) {
            (None, None) => Ok(String::new()),
            _ => Err(Error::Other(format!(
                "comparison cannot create or remove a section shell at {location}"
            ))),
        };
    };
    if original.header_refs != edited.header_refs || original.footer_refs != edited.footer_refs {
        return Err(Error::Other(format!(
            "comparison cannot change related-story references at {location}"
        )));
    }
    if metadata.options.ignore_formatting {
        return section_property_xml(original);
    }
    let mut original_modeled = original.clone();
    original_modeled.change = None;
    let mut edited_modeled = edited.clone();
    edited_modeled.change = None;
    if original_modeled == edited_modeled {
        return section_property_xml(original);
    }
    if original.extra_xml != edited.extra_xml {
        formatting_diagnostic(diagnostics, location.to_owned());
        edited_modeled.extra_xml = original.extra_xml.clone();
    }
    let previous = section_property_xml(&original_modeled)?;
    let change = metadata.ids.revision(
        "sectPrChange",
        metadata.author,
        metadata.timestamp,
        &previous,
    )?;
    let current = section_property_xml(&edited_modeled)?;
    inject_before_close(&current, "</w:sectPr>", &change)
}

fn deleted_paragraph(paragraph: &CT_P, metadata: &mut Metadata<'_>) -> Result<String> {
    let mut output = String::from("<w:p>");
    output.push_str(&paragraph_mark_properties(
        paragraph.properties.as_ref(),
        "del",
        metadata,
    )?);
    for run in &paragraph.runs {
        let run = deleted_run_xml(run)?;
        output.push_str(&metadata.ids.revision(
            "del",
            metadata.author,
            metadata.timestamp,
            &run,
        )?);
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn inserted_paragraph(paragraph: &CT_P, metadata: &mut Metadata<'_>) -> Result<String> {
    let mut output = String::from("<w:p>");
    output.push_str(&paragraph_mark_properties(
        paragraph.properties.as_ref(),
        "ins",
        metadata,
    )?);
    for run in &paragraph.runs {
        let run = run_xml(run)?;
        output.push_str(&metadata.ids.revision(
            "ins",
            metadata.author,
            metadata.timestamp,
            &run,
        )?);
    }
    output.push_str("</w:p>");
    Ok(output)
}

fn paragraph_mark_properties(
    properties: Option<&CT_PPr>,
    kind: &str,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let marker = metadata
        .ids
        .marker(kind, metadata.author, metadata.timestamp)?;
    let mut properties = properties.cloned().unwrap_or_default();
    properties.rpr = Some(properties.rpr.take().unwrap_or_default());
    let mut xml = property_xml(&properties)?;
    let run_properties = direct_word_element_spans(&xml, "rPr")?;
    if let Some(run_span) = run_properties.first() {
        let updated = append_word_child(&xml[run_span.clone()], "rPr", &marker)?;
        xml.replace_range(run_span.clone(), &updated);
        Ok(xml)
    } else {
        append_word_child(&xml, "pPr", &format!("<w:rPr>{marker}</w:rPr>"))
    }
}

fn compare_table(
    original: &CT_Tbl,
    edited: &CT_Tbl,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if original.grid != edited.grid {
        return Err(Error::Other(format!(
            "comparison cannot revise a table grid change at {location}"
        )));
    }
    if original.extra_xml != edited.extra_xml
        || table_control_boundaries(original) != table_control_boundaries(edited)
    {
        return Err(Error::Other(format!(
            "comparison cannot revise table boundary structures at {location}"
        )));
    }
    let original_signatures = original
        .rows
        .iter()
        .map(|row| row_signature_with_options(row, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .rows
        .iter()
        .map(|row| row_signature_with_options(row, metadata.options))
        .collect::<Vec<_>>();
    let mut source = table_xml(original)?;
    let properties = table_properties_xml(
        original.properties.as_ref(),
        edited.properties.as_ref(),
        location,
        metadata,
        diagnostics,
    )?;
    let property_spans = direct_word_element_spans(&source, "tblPr")?;
    match (property_spans.first(), properties.is_empty()) {
        (Some(span), false) => source.replace_range(span.clone(), &properties),
        (Some(span), true) => source.replace_range(span.clone(), ""),
        (None, false) => {
            let open = source
                .find('>')
                .ok_or_else(|| Error::Other("table XML has no start".to_owned()))?
                + 1;
            source = format!("{}{}{}", &source[..open], properties, &source[open..]);
        }
        (None, true) => {}
    }
    let row_spans = direct_word_element_spans(&source, "tr")?;
    if row_spans.len() != original.rows.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate direct table rows at {location}"
        )));
    }
    let aligned = align(&original_signatures, &edited_signatures);
    let mut output = String::new();
    let mut cursor = 0usize;
    for (alignment_index, &(left, right)) in aligned.iter().enumerate() {
        if let Some(index) = left {
            let span = &row_spans[index];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
        } else {
            let next_row_start = aligned[alignment_index + 1..]
                .iter()
                .find_map(|(next_left, _)| next_left.map(|index| row_spans[index].start))
                .unwrap_or_else(|| source.rfind("</w:tbl>").unwrap_or(source.len()));
            output.push_str(&source[cursor..next_row_start]);
            cursor = next_row_start;
        }
        match (left, right) {
            (Some(i), Some(j)) => {
                output.push_str(&compare_row(
                    &original.rows[i],
                    &edited.rows[j],
                    &format!("{location}/row[{j}]"),
                    metadata,
                    diagnostics,
                )?);
            }
            (Some(i), None) => output.push_str(&marked_row(&original.rows[i], "del", metadata)?),
            (None, Some(j)) => output.push_str(&marked_row(&edited.rows[j], "ins", metadata)?),
            (None, None) => unreachable!(),
        }
    }
    output.push_str(&source[cursor..]);
    let spans = direct_word_element_spans(&output, "sdt")?;
    if spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate table content controls at {location}"
        )));
    }
    let mut replacements = Vec::with_capacity(spans.len());
    for (index, ((_, _, left), (_, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &replacements)
}

fn table_properties_xml(
    original: Option<&CT_TblPr>,
    edited: Option<&CT_TblPr>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if metadata.options.ignore_formatting {
        return original
            .map(table_property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }
    let original_modeled = modeled_table_properties(original);
    let edited_modeled = modeled_table_properties(edited);
    if original_modeled == edited_modeled {
        return original
            .map(table_property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }
    let mut current = edited.cloned().unwrap_or_default();
    current.change = None;
    let (original_extra_xml, original_revision_xml) = original
        .map(|properties| {
            (
                properties.extra_xml.clone(),
                properties.revision_xml.clone(),
            )
        })
        .unwrap_or_default();
    if current.extra_xml != original_extra_xml || current.revision_xml != original_revision_xml {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    current.extra_xml = original_extra_xml;
    current.revision_xml = original_revision_xml;
    let previous = original_modeled
        .as_ref()
        .map(table_property_xml)
        .transpose()?
        .unwrap_or_else(|| "<w:tblPr/>".to_owned());
    let change = metadata.ids.revision(
        "tblPrChange",
        metadata.author,
        metadata.timestamp,
        &previous,
    )?;
    let current = table_property_xml(&current)?;
    inject_before_close(&current, "</w:tblPr>", &change)
}

fn modeled_table_properties(properties: Option<&CT_TblPr>) -> Option<CT_TblPr> {
    properties.cloned().map(|mut properties| {
        properties.change = None;
        properties.revision_xml.clear();
        properties.extra_xml.clear();
        properties
    })
}

fn table_property_xml(properties: &CT_TblPr) -> Result<String> {
    let mut bytes = Vec::new();
    properties.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn table_control_boundaries(table: &CT_Tbl) -> Vec<(usize, usize)> {
    table
        .content_controls
        .iter()
        .map(|(at, raw_before, _)| (*at, *raw_before))
        .collect()
}

fn marked_table(table: &CT_Tbl, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    let source = table_xml(table)?;
    let replacements = table
        .rows
        .iter()
        .map(|row| marked_row(row, kind, metadata))
        .collect::<Result<Vec<_>>>()?;
    let output = replace_direct_word_elements(&source, "tr", &replacements)?;
    let control_spans = direct_word_element_spans(&output, "sdt")?;
    if control_spans.len() != table.content_controls.len() {
        return Err(Error::Other(
            "comparison could not correlate table content controls while marking rows".to_owned(),
        ));
    }
    let replacements = table
        .content_controls
        .iter()
        .zip(&control_spans)
        .map(|((_, _, control), span)| {
            mark_control_owned_rows(control, kind, metadata, &output[span.clone()])
        })
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(&output, "sdt", &replacements)
}

fn mark_control_owned_rows(
    control: &CT_Sdt,
    kind: &str,
    metadata: &mut Metadata<'_>,
    original_xml: &str,
) -> Result<String> {
    let mut content = String::new();
    for child in &control.content {
        content.push_str(&match child {
            SdtContent::Paragraph(paragraph) => paragraph_xml(paragraph)?,
            SdtContent::Table(table) => marked_table(table, kind, metadata)?,
            SdtContent::Row(row) => marked_row(row, kind, metadata)?,
            SdtContent::Cell(cell) => cell_xml(cell)?,
            SdtContent::Run(run) => paragraph_owned_run_xml(run)?,
            SdtContent::ContentControl(nested) => {
                let nested_xml = control_xml(nested)?;
                mark_control_owned_rows(nested, kind, metadata, &nested_xml)?
            }
            SdtContent::RawXml(raw) => String::from_utf8(raw.clone()).map_err(utf8_error)?,
        });
    }
    replace_element_inner(original_xml, "w:sdtContent", &content)
}

fn marked_row(row: &CT_Row, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    let marker = metadata
        .ids
        .marker(kind, metadata.author, metadata.timestamp)?;
    let mut xml = row_xml(row)?;
    let properties = direct_word_element_spans(&xml, "trPr")?;
    if let Some(properties_span) = properties.first() {
        let updated = append_word_child(&xml[properties_span.clone()], "trPr", &marker)?;
        xml.replace_range(properties_span.clone(), &updated);
        Ok(xml)
    } else {
        let open = xml
            .find('>')
            .ok_or_else(|| Error::Other("row XML has no start".to_owned()))?
            + 1;
        Ok(format!(
            "{}<w:trPr>{marker}</w:trPr>{}",
            &xml[..open],
            &xml[open..]
        ))
    }
}

fn compare_row(
    original: &CT_Row,
    edited: &CT_Row,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if original.cells.len() != edited.cells.len()
        || original.extra_xml != edited.extra_xml
        || row_control_boundaries(original) != row_control_boundaries(edited)
    {
        return Err(Error::Other(format!(
            "comparison cannot revise row boundary structures at {location}"
        )));
    }
    if !metadata.options.ignore_formatting && row_formatting(original) != row_formatting(edited) {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    let mut output = row_xml(original)?;
    let cell_spans = direct_word_element_spans(&output, "tc")?;
    if cell_spans.len() != original.cells.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate row cells at {location}"
        )));
    }
    let mut cell_replacements = Vec::with_capacity(cell_spans.len());
    for (index, (left, right)) in original.cells.iter().zip(&edited.cells).enumerate() {
        cell_replacements.push(compare_cell_from_xml(
            left,
            right,
            &format!("{location}/cell[{index}]"),
            metadata,
            diagnostics,
            &output[cell_spans[index].clone()],
        )?);
    }
    output = replace_direct_word_elements(&output, "tc", &cell_replacements)?;

    let control_spans = direct_word_element_spans(&output, "sdt")?;
    if control_spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate row content controls at {location}"
        )));
    }
    let mut control_replacements = Vec::with_capacity(control_spans.len());
    for (index, ((_, _, left), (_, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        control_replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[control_spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &control_replacements)
}

fn row_control_boundaries(row: &CT_Row) -> Vec<(usize, usize)> {
    row.content_controls
        .iter()
        .map(|(at, raw_before, _)| (*at, *raw_before))
        .collect()
}

fn compare_cell_from_xml(
    original: &CT_Tc,
    edited: &CT_Tc,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
    original_xml: &str,
) -> Result<String> {
    if original.content.len() != edited.content.len() || original.extra_xml != edited.extra_xml {
        return Err(Error::Other(format!(
            "comparison cannot revise cell boundary structures at {location}"
        )));
    }
    if !metadata.options.ignore_formatting && original.properties != edited.properties {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    let children = cell_child_spans(original, original_xml, location)?;
    let mut replacements = Vec::with_capacity(children.len());
    for (index, ((left, right), (_, span))) in original
        .content
        .iter()
        .zip(&edited.content)
        .zip(&children)
        .enumerate()
    {
        let child_location = format!("{location}/content[{index}]");
        let source = &original_xml[span.clone()];
        replacements.push(match (left, right) {
            (CellContent::Paragraph(left), CellContent::Paragraph(right)) => {
                compare_paragraph(left, right, &child_location, metadata, diagnostics)?
            }
            (CellContent::Table(left), CellContent::Table(right)) => {
                compare_table(left, right, &child_location, metadata, diagnostics)?
            }
            (CellContent::ContentControl(left), CellContent::ContentControl(right)) => {
                compare_control_from_xml(
                    left,
                    right,
                    &child_location,
                    metadata,
                    diagnostics,
                    source,
                )?
            }
            _ => {
                return Err(Error::Other(format!(
                    "comparison cannot revise incompatible cell children at {child_location}"
                )));
            }
        });
    }
    replace_ranges(original_xml, &children, &replacements)
}

fn compare_control(
    original: &CT_Sdt,
    edited: &CT_Sdt,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_xml = control_xml(original)?;
    compare_control_from_xml(
        original,
        edited,
        location,
        metadata,
        diagnostics,
        &original_xml,
    )
}

fn compare_control_from_xml(
    original: &CT_Sdt,
    edited: &CT_Sdt,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
    original_xml: &str,
) -> Result<String> {
    if control_property_signature(original) != control_property_signature(edited) {
        return Err(Error::Other(format!(
            "comparison cannot revise content-control properties at {location}"
        )));
    }
    let original_content = modeled_control_content(original);
    let edited_content = modeled_control_content(edited);
    let direct_run_or_raw =
        |content: &&SdtContent| matches!(content, SdtContent::Run(_) | SdtContent::RawXml(_));
    if uses_attributed_run_path(metadata.options)
        && original_content.iter().all(direct_run_or_raw)
        && edited_content.iter().all(direct_run_or_raw)
        && (!original_content.is_empty() || !edited_content.is_empty())
        && inline_control_raw_boundaries(&original_content, metadata.options)
            == inline_control_raw_boundaries(&edited_content, metadata.options)
    {
        let original_runs = original_content
            .iter()
            .filter_map(|content| match content {
                SdtContent::Run(run) => Some(run.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        let edited_runs = edited_content
            .iter()
            .filter_map(|content| match content {
                SdtContent::Run(run) => Some(run.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        let content_spans = direct_word_element_spans(original_xml, "sdtContent")?;
        let content_span = content_spans.first().ok_or_else(|| {
            Error::Other(format!(
                "comparison could not find inline-control content at {location}"
            ))
        })?;
        let content_source = &original_xml[content_span.clone()];
        let inner = element_inner_range_any_prefix(content_source, "sdtContent")?;
        let compared = compare_control_runs_with_options(
            &original_runs,
            &edited_runs,
            Some(&content_source[inner]),
            location,
            metadata,
            diagnostics,
        )?;
        return replace_element_inner(original_xml, "w:sdtContent", &compared);
    }
    let whitespace_slots = control_whitespace_slots(original)?;
    let original_signatures = original_content
        .iter()
        .map(|content| control_content_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited_content
        .iter()
        .map(|content| control_content_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    let aligned = expand_control_alignment(
        align(&original_signatures, &edited_signatures),
        &original_content,
        &edited_content,
    );
    let mut content: Vec<(bool, String)> = Vec::new();
    let mut whitespace_emitted = vec![false; whitespace_slots.len()];
    for (position, (left, right)) in aligned.iter().copied().enumerate() {
        let whitespace_boundary = left.or_else(|| {
            let right = right?;
            let (Some(next_left), None) = aligned.get(position + 1).copied()? else {
                return None;
            };
            (matches!(original_content[next_left], SdtContent::Paragraph(_))
                && matches!(edited_content[right], SdtContent::Table(_)))
            .then_some(next_left)
        });
        if let Some(index) = whitespace_boundary
            && !whitespace_emitted[index]
            && !whitespace_slots[index].is_empty()
        {
            content.push((false, whitespace_slots[index].clone()));
            whitespace_emitted[index] = true;
        }
        let next_is_paragraph = aligned.get(position + 1).is_some_and(|(left, right)| {
            right
                .and_then(|index| edited_content.get(index).copied())
                .or_else(|| left.and_then(|index| original_content.get(index).copied()))
                .is_some_and(|content| matches!(content, SdtContent::Paragraph(_)))
        });
        let child_location = format!("{location}/content[{position}]");
        match (left, right) {
            (Some(i), Some(j)) => content.push((
                matches!(edited_content[j], SdtContent::Paragraph(_)),
                compare_control_content(
                    original_content[i],
                    edited_content[j],
                    &child_location,
                    metadata,
                    diagnostics,
                )?,
            )),
            (Some(i), None) => {
                if matches!(original_content[i], SdtContent::Paragraph(_)) && !next_is_paragraph {
                    mark_previous_paragraph(&mut content, "del", metadata)?;
                    content.push((
                        true,
                        marked_control_content(original_content[i], "del", false, metadata)?,
                    ));
                } else {
                    content.push((
                        matches!(original_content[i], SdtContent::Paragraph(_)),
                        marked_control_content(original_content[i], "del", true, metadata)?,
                    ));
                }
            }
            (None, Some(j)) => {
                if matches!(edited_content[j], SdtContent::Paragraph(_)) && !next_is_paragraph {
                    mark_previous_paragraph(&mut content, "ins", metadata)?;
                    content.push((
                        true,
                        marked_control_content(edited_content[j], "ins", false, metadata)?,
                    ));
                } else {
                    content.push((
                        matches!(edited_content[j], SdtContent::Paragraph(_)),
                        marked_control_content(edited_content[j], "ins", true, metadata)?,
                    ));
                }
            }
            (None, None) => unreachable!(),
        }
    }
    if let Some(trailing) = whitespace_slots.last()
        && !trailing.is_empty()
    {
        content.push((false, trailing.clone()));
    }
    let content = content.into_iter().map(|(_, xml)| xml).collect::<String>();
    replace_element_inner(original_xml, "w:sdtContent", &content)
}

fn inline_control_raw_boundaries(
    content: &[&SdtContent],
    options: &ComparisonOptions,
) -> Vec<(usize, Vec<u8>)> {
    let mut runs = Vec::new();
    let mut raw = Vec::new();
    for child in content {
        match child {
            SdtContent::Run(run) => runs.push((*run).clone()),
            SdtContent::RawXml(bytes) => {
                let boundary = attributed_run_units(&runs, options)
                    .into_iter()
                    .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
                    .count();
                raw.push((boundary, bytes.clone()));
            }
            _ => unreachable!(),
        }
    }
    raw
}

fn compare_control_content(
    original: &SdtContent,
    edited: &SdtContent,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    match (original, edited) {
        (SdtContent::Paragraph(left), SdtContent::Paragraph(right)) => {
            compare_paragraph(left, right, location, metadata, diagnostics)
        }
        (SdtContent::Table(left), SdtContent::Table(right)) => {
            compare_table(left, right, location, metadata, diagnostics)
        }
        (SdtContent::ContentControl(left), SdtContent::ContentControl(right)) => {
            compare_control(left, right, location, metadata, diagnostics)
        }
        (SdtContent::Row(left), SdtContent::Row(right)) => {
            compare_row(left, right, location, metadata, diagnostics)
        }
        (SdtContent::Cell(left), SdtContent::Cell(right)) => compare_cell_from_xml(
            left,
            right,
            location,
            metadata,
            diagnostics,
            &cell_xml(left)?,
        ),
        (SdtContent::Run(left), SdtContent::Run(right)) => {
            if uses_attributed_run_path(metadata.options) {
                compare_control_runs_with_options(
                    std::slice::from_ref(left),
                    std::slice::from_ref(right),
                    None,
                    location,
                    metadata,
                    diagnostics,
                )
            } else if run_signature(left) == run_signature(right) {
                if !metadata.options.ignore_formatting && left.properties != right.properties {
                    formatting_diagnostic(diagnostics, location.to_owned());
                }
                run_xml(left)
            } else {
                let deleted = metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted_run_xml(left)?,
                )?;
                let inserted = metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &run_xml(right)?,
                )?;
                Ok(format!("{deleted}{inserted}"))
            }
        }
        (SdtContent::RawXml(left), SdtContent::RawXml(right)) if left == right => {
            String::from_utf8(left.clone()).map_err(utf8_error)
        }
        _ => Err(Error::Other(format!(
            "comparison cannot revise incompatible content-control children at {location}"
        ))),
    }
}

fn compare_control_runs_with_options(
    original: &[CT_R],
    edited: &[CT_R],
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_units = attributed_run_units(original, metadata.options);
    let edited_units = attributed_run_units(edited, metadata.options);
    let original_signatures = original_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();
    let edited_signatures = edited_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();
    let aligned = align(&original_signatures, &edited_signatures);
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &edited_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let grouped = coalesced_granular_alignment(
        &aligned,
        &original_units,
        &edited_units,
        &original_signatures,
        &edited_signatures,
    );
    let mut grouped_alignment = Vec::with_capacity(grouped.len());
    let mut replacements = Vec::with_capacity(grouped.len());
    for (action, members) in grouped {
        grouped_alignment.push(aligned[members.start]);
        let left_indices = members
            .clone()
            .filter_map(|index| aligned[index].0)
            .collect::<Vec<_>>();
        let right_indices = members
            .filter_map(|index| aligned[index].1)
            .collect::<Vec<_>>();
        let left = merge_unit_runs(&original_units, &left_indices);
        let right = merge_unit_runs(&edited_units, &right_indices);
        replacements.push(match action {
            GranularAction::Preserve => paragraph_owned_run_xml(left.as_ref().unwrap())?,
            GranularAction::Equal => compared_run_xml(
                left.as_ref().unwrap(),
                right.as_ref().unwrap(),
                location,
                metadata,
                diagnostics,
            )?,
            GranularAction::Replace => {
                let left = left.as_ref().unwrap();
                let deleted = metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted_run_xml(left)?,
                )?;
                let inserted = metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &policy_inserted_run_xml(
                        right.as_ref().unwrap(),
                        Some(left),
                        metadata.options,
                    )?,
                )?;
                format!("{deleted}{inserted}")
            }
            GranularAction::Delete => metadata.ids.revision(
                "del",
                metadata.author,
                metadata.timestamp,
                &deleted_run_xml(left.as_ref().unwrap())?,
            )?,
            GranularAction::Insert => metadata.ids.revision(
                "ins",
                metadata.author,
                metadata.timestamp,
                &paragraph_owned_run_xml(right.as_ref().unwrap())?,
            )?,
            GranularAction::Drop => String::new(),
        });
    }
    let Some(source) = original_source else {
        return Ok(replacements.concat());
    };
    let open = "<w:inlineControl>";
    let wrapped = format!("{open}{source}</w:inlineControl>");
    let spans = direct_word_element_spans(&wrapped, "r")?
        .into_iter()
        .map(|span| span.start - open.len()..span.end - open.len())
        .collect::<Vec<_>>();
    if spans.len() != original.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate inline-control run owners at {location}"
        )));
    }
    let insertion_boundary = spans.first().map_or(source.len(), |span| span.start);
    let mut output = source[..insertion_boundary].to_owned();
    let mut cursor = insertion_boundary;
    let mut consumed_owner = None;
    for ((left, _), replacement) in grouped_alignment.iter().zip(&replacements) {
        if let Some(unit) = left.map(|index| &original_units[index])
            && consumed_owner != Some(unit.owner)
        {
            let span = &spans[unit.owner];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
            consumed_owner = Some(unit.owner);
        }
        output.push_str(replacement);
    }
    output.push_str(&source[cursor..]);
    Ok(output)
}

fn marked_control_content(
    content: &SdtContent,
    kind: &str,
    paragraph_marker: bool,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    match content {
        SdtContent::Paragraph(paragraph) if paragraph_marker && kind == "del" => {
            deleted_paragraph(paragraph, metadata)
        }
        SdtContent::Paragraph(paragraph) if paragraph_marker => {
            inserted_paragraph(paragraph, metadata)
        }
        SdtContent::Paragraph(paragraph) if kind == "del" => {
            deleted_paragraph_content(&BodyContent::Paragraph(paragraph.clone()), metadata)
        }
        SdtContent::Paragraph(paragraph) => {
            inserted_paragraph_content(&BodyContent::Paragraph(paragraph.clone()), metadata)
        }
        SdtContent::Table(table) => marked_table(table, kind, metadata),
        SdtContent::Row(row) => marked_row(row, kind, metadata),
        SdtContent::Run(run) if kind == "del" => metadata.ids.revision(
            kind,
            metadata.author,
            metadata.timestamp,
            &deleted_run_xml(run)?,
        ),
        SdtContent::Run(run) => {
            metadata
                .ids
                .revision(kind, metadata.author, metadata.timestamp, &run_xml(run)?)
        }
        SdtContent::RawXml(raw) => String::from_utf8(raw.clone()).map_err(utf8_error),
        SdtContent::Cell(_) | SdtContent::ContentControl(_) => Err(Error::Other(
            "comparison cannot add or remove an unmatched cell or nested content control"
                .to_owned(),
        )),
    }
}

fn compared_run_xml(
    original: &CT_R,
    edited: &CT_R,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if metadata.options.ignore_fields && (run_is_field(original) || run_is_field(edited)) {
        return paragraph_owned_run_xml(original);
    }
    if metadata.options.ignore_formatting {
        return paragraph_owned_run_xml(original);
    }
    if let (Some(original_field), Some(edited_field)) = (run_field(original), run_field(edited)) {
        return compare_field_xml(original_field, edited_field, metadata);
    }
    let original_properties = modeled_run_properties(original);
    let edited_properties = modeled_run_properties(edited);
    if original_properties == edited_properties {
        if original.properties != edited.properties {
            formatting_diagnostic(diagnostics, location.to_owned());
        }
        return paragraph_owned_run_xml(original);
    }
    let previous = run_property_xml(original_properties.as_ref())?;
    let change =
        metadata
            .ids
            .revision("rPrChange", metadata.author, metadata.timestamp, &previous)?;
    if unmodeled_run_properties_differ(original, edited) {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    let mut current_run = edited.clone();
    preserve_unmodeled_run_properties(&mut current_run, original);
    let mut current = run_xml(&current_run)?;
    let properties = direct_word_element_spans(&current, "rPr")?;
    if let Some(span) = properties.first() {
        let updated = append_word_child(&current[span.clone()], "rPr", &change)?;
        current.replace_range(span.clone(), &updated);
    } else {
        let open = current
            .find('>')
            .ok_or_else(|| Error::Other("run XML has no start".to_owned()))?
            + 1;
        current = format!(
            "{}<w:rPr>{change}</w:rPr>{}",
            &current[..open],
            &current[open..]
        );
    }
    Ok(current)
}

fn compare_field_xml(
    original: &rdocx_oxml::text::Field,
    edited: &rdocx_oxml::text::Field,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let original_xml = field_xml(original)?;
    let edited_xml = field_xml(edited)?;
    if original == edited {
        return Ok(original_xml);
    }
    let same_instruction = original.effective_instruction() == edited.effective_instruction();
    if same_instruction && original.is_complex() == edited.is_complex() {
        if original.is_complex() {
            let (before, old_result, after) = complex_field_result(&original_xml)?;
            let (_, new_result, _) = complex_field_result(&edited_xml)?;
            return Ok(format!(
                "{before}{}{after}",
                tracked_field_result(&old_result, &new_result, metadata)?
            ));
        }
        let original_inner = element_inner_range(&original_xml, "fldSimple")?;
        let edited_inner = element_inner_range(&edited_xml, "fldSimple")?;
        let old_result = original_xml[original_inner.clone()].to_owned();
        let new_result = edited_xml[edited_inner].to_owned();
        let mut tracked = original_xml;
        tracked.replace_range(
            original_inner,
            &tracked_field_result(&old_result, &new_result, metadata)?,
        );
        return Ok(tracked);
    }

    let deleted = metadata.ids.revision(
        "del",
        metadata.author,
        metadata.timestamp,
        &deleted_text_xml(&original_xml),
    )?;
    let inserted =
        metadata
            .ids
            .revision("ins", metadata.author, metadata.timestamp, &edited_xml)?;
    Ok(format!("{deleted}{inserted}"))
}

fn element_inner_range(xml: &str, local: &str) -> Result<Range<usize>> {
    let start = xml
        .find('>')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("{local} source has no start tag")))?;
    let end = xml
        .rfind(&format!("</w:{local}>"))
        .ok_or_else(|| Error::Other(format!("{local} source has no end tag")))?;
    Ok(start..end)
}

fn field_xml(field: &rdocx_oxml::text::Field) -> Result<String> {
    paragraph_owned_run_xml(&CT_R {
        properties: None,
        content: vec![RunContent::Field(field.clone())],
        extra_xml: Vec::new(),
        extra_xml_positions: Vec::new(),
        alt_drawings: Vec::new(),
    })
}

fn tracked_field_result(
    original: &str,
    edited: &str,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let deleted = metadata.ids.revision(
        "del",
        metadata.author,
        metadata.timestamp,
        &deleted_text_xml(original),
    )?;
    let inserted = metadata
        .ids
        .revision("ins", metadata.author, metadata.timestamp, edited)?;
    Ok(format!("{deleted}{inserted}"))
}

fn deleted_text_xml(xml: &str) -> String {
    xml.replace("<w:t", "<w:delText")
        .replace("</w:t>", "</w:delText>")
}

fn complex_field_result(xml: &str) -> Result<(String, String, String)> {
    let separate = xml
        .find("fldCharType=\"separate\"")
        .or_else(|| xml.find("fldCharType='separate'"))
        .ok_or_else(|| Error::Other("complex field source has no separate boundary".to_owned()))?;
    let (_, result_start) = containing_run(xml, separate)?;
    let end_marker = xml[result_start..]
        .find("fldCharType=\"end\"")
        .or_else(|| xml[result_start..].find("fldCharType='end'"))
        .map(|offset| result_start + offset)
        .ok_or_else(|| Error::Other("complex field source has no end boundary".to_owned()))?;
    let (result_end, _) = containing_run(xml, end_marker)?;
    Ok((
        xml[..result_start].to_owned(),
        xml[result_start..result_end].to_owned(),
        xml[result_end..].to_owned(),
    ))
}

fn containing_run(xml: &str, at: usize) -> Result<(usize, usize)> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<(usize, bool)>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                stack.push((before, element.local_name().as_ref() == b"r"));
            }
            Event::End(_) => {
                let Some((start, run)) = stack.pop() else {
                    return Err(Error::Other("complex field XML is unbalanced".to_owned()));
                };
                if run && start <= at && at < after {
                    return Ok((start, after));
                }
            }
            Event::Empty(_) => {}
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Err(Error::Other(
        "complex field boundary has no owning run".to_owned(),
    ))
}

fn modeled_run_properties(run: &CT_R) -> Option<rdocx_oxml::properties::CT_RPr> {
    run.properties.clone().map(|mut properties| {
        properties.revision_markers.clear();
        properties.change = None;
        properties.revision_xml.clear();
        properties.revision_xml_positions.clear();
        properties.language_extra_attributes.clear();
        properties
    })
}

fn unmodeled_run_properties_differ(original: &CT_R, edited: &CT_R) -> bool {
    let original = original.properties.as_ref();
    let edited = edited.properties.as_ref();
    original
        .map(|properties| properties.language_extra_attributes.as_slice())
        .unwrap_or_default()
        != edited
            .map(|properties| properties.language_extra_attributes.as_slice())
            .unwrap_or_default()
        || original
            .map(|properties| properties.revision_xml.as_slice())
            .unwrap_or_default()
            != edited
                .map(|properties| properties.revision_xml.as_slice())
                .unwrap_or_default()
        || original
            .map(|properties| properties.revision_xml_positions.as_slice())
            .unwrap_or_default()
            != edited
                .map(|properties| properties.revision_xml_positions.as_slice())
                .unwrap_or_default()
}

fn preserve_unmodeled_run_properties(current: &mut CT_R, original: &CT_R) {
    let current_properties = current
        .properties
        .get_or_insert_with(rdocx_oxml::properties::CT_RPr::default);
    if let Some(original_properties) = original.properties.as_ref() {
        current_properties.language_extra_attributes =
            original_properties.language_extra_attributes.clone();
        current_properties.revision_xml = original_properties.revision_xml.clone();
        current_properties.revision_xml_positions =
            original_properties.revision_xml_positions.clone();
    } else {
        current_properties.language_extra_attributes.clear();
        current_properties.revision_xml.clear();
        current_properties.revision_xml_positions.clear();
    }
}

fn run_property_xml(properties: Option<&rdocx_oxml::properties::CT_RPr>) -> Result<String> {
    let mut bytes = Vec::new();
    if let Some(properties) = properties {
        properties.to_xml(&mut Writer::new(&mut bytes))?;
    }
    if bytes.is_empty() {
        Ok("<w:rPr/>".to_owned())
    } else {
        String::from_utf8(bytes).map_err(utf8_error)
    }
}

fn formatting_diagnostic(diagnostics: &mut Vec<ComparisonDiagnostic>, location: String) {
    diagnostics.push(ComparisonDiagnostic {
        location,
        message: "formatting differs and the original formatting was retained".to_owned(),
    });
}

fn body_location(location: &str, content: &BodyContent, index: usize) -> String {
    match content {
        BodyContent::Paragraph(_) => format!("{location}/paragraph[{index}]"),
        BodyContent::Table(_) => format!("{location}/table[{index}]"),
        BodyContent::ContentControl(_) => format!("{location}/content-control[{index}]"),
        BodyContent::RawXml(_) => format!("{location}/raw[{index}]"),
    }
}

fn align(original: &[String], edited: &[String]) -> Vec<(Option<usize>, Option<usize>)> {
    let mut lcs = vec![vec![0usize; edited.len() + 1]; original.len() + 1];
    for i in (0..original.len()).rev() {
        for j in (0..edited.len()).rev() {
            lcs[i][j] = if original[i] == edited[j] {
                1 + lcs[i + 1][j + 1]
            } else {
                lcs[i + 1][j].max(lcs[i][j + 1])
            };
        }
    }
    let mut matches = Vec::new();
    let (mut i, mut j) = (0usize, 0usize);
    while i < original.len() && j < edited.len() {
        if original[i] == edited[j] {
            matches.push((i, j));
            i += 1;
            j += 1;
        } else if lcs[i + 1][j] >= lcs[i][j + 1] {
            i += 1;
        } else {
            j += 1;
        }
    }
    let mut result = Vec::new();
    let (mut left, mut right) = (0usize, 0usize);
    for (matched_left, matched_right) in matches
        .into_iter()
        .chain(std::iter::once((original.len(), edited.len())))
    {
        let paired = (matched_left - left).min(matched_right - right);
        for offset in 0..paired {
            result.push((Some(left + offset), Some(right + offset)));
        }
        for index in left + paired..matched_left {
            result.push((Some(index), None));
        }
        for index in right + paired..matched_right {
            result.push((None, Some(index)));
        }
        if matched_left < original.len() {
            result.push((Some(matched_left), Some(matched_right)));
        }
        left = matched_left.saturating_add(1);
        right = matched_right.saturating_add(1);
    }
    result
}

fn expand_body_alignment(
    aligned: Vec<(Option<usize>, Option<usize>)>,
    original: &[BodyContent],
    edited: &[BodyContent],
) -> Vec<(Option<usize>, Option<usize>)> {
    let mut expanded = Vec::with_capacity(aligned.len());
    for (left, right) in aligned {
        match (left, right) {
            (Some(i), Some(j))
                if matches!(original[i], BodyContent::Paragraph(_))
                    && matches!(edited[j], BodyContent::Table(_)) =>
            {
                expanded.push((None, Some(j)));
                expanded.push((Some(i), None));
            }
            (Some(i), Some(j))
                if matches!(original[i], BodyContent::Table(_))
                    && matches!(edited[j], BodyContent::Paragraph(_)) =>
            {
                expanded.push((Some(i), None));
                expanded.push((None, Some(j)));
            }
            pair => expanded.push(pair),
        }
    }
    expanded
}

fn expand_control_alignment(
    aligned: Vec<(Option<usize>, Option<usize>)>,
    original: &[&SdtContent],
    edited: &[&SdtContent],
) -> Vec<(Option<usize>, Option<usize>)> {
    let mut expanded = Vec::with_capacity(aligned.len());
    for (left, right) in aligned {
        match (left, right) {
            (Some(i), Some(j))
                if matches!(original[i], SdtContent::Paragraph(_))
                    && matches!(edited[j], SdtContent::Table(_)) =>
            {
                expanded.push((None, Some(j)));
                expanded.push((Some(i), None));
            }
            (Some(i), Some(j))
                if matches!(original[i], SdtContent::Table(_))
                    && matches!(edited[j], SdtContent::Paragraph(_)) =>
            {
                expanded.push((Some(i), None));
                expanded.push((None, Some(j)));
            }
            pair => expanded.push(pair),
        }
    }
    expanded
}

fn body_signature(content: &BodyContent) -> String {
    match content {
        BodyContent::Paragraph(paragraph) => format!("p:{}", paragraph_signature(paragraph)),
        BodyContent::Table(table) => format!("t:{}", table_signature(table)),
        BodyContent::ContentControl(control) => format!("s:{}", control_signature(control)),
        BodyContent::RawXml(raw) => format!("x:{raw:?}"),
    }
}

fn paragraph_signature(paragraph: &CT_P) -> String {
    let numbering = paragraph_numbering(paragraph);
    format!(
        "{numbering:?}:{:?}:{:?}:{:?}:{:?}:{:?}:{:?}",
        paragraph.runs.iter().map(run_signature).collect::<Vec<_>>(),
        paragraph.hyperlinks,
        paragraph.comment_ranges,
        paragraph.bookmark_markers,
        paragraph.extra_xml,
        paragraph
            .content_controls
            .iter()
            .map(|(at, raw_before, markers_before, control)| (
                at,
                raw_before,
                markers_before,
                control_signature(control),
            ))
            .collect::<Vec<_>>(),
    )
}

fn paragraph_numbering(paragraph: &CT_P) -> Option<(Option<u32>, Option<u32>)> {
    paragraph.properties.as_ref().and_then(|properties| {
        (properties.num_id.is_some() || properties.num_ilvl.is_some())
            .then_some((properties.num_id, properties.num_ilvl))
    })
}

fn run_signature(run: &CT_R) -> String {
    format!(
        "{:?}:{:?}:{:?}",
        run.content
            .iter()
            .map(run_content_signature)
            .collect::<Vec<_>>(),
        run.extra_xml,
        run.extra_xml_positions
    )
}

fn run_content_signature(content: &RunContent) -> String {
    match content {
        RunContent::Field(_) => "field-owner".to_owned(),
        content => format!("{content:?}"),
    }
}

fn table_signature(table: &CT_Tbl) -> String {
    format!(
        "{:?}:{:?}:{:?}:{:?}",
        table.grid,
        table.rows.iter().map(row_signature).collect::<Vec<_>>(),
        table.extra_xml,
        table
            .content_controls
            .iter()
            .map(|(at, raw_before, control)| (at, raw_before, control_signature(control)))
            .collect::<Vec<_>>()
    )
}

fn row_signature(row: &CT_Row) -> String {
    format!(
        "{:?}:{:?}:{:?}",
        row.cells.iter().map(cell_signature).collect::<Vec<_>>(),
        row.extra_xml,
        row.content_controls
            .iter()
            .map(|(at, raw_before, control)| (at, raw_before, control_signature(control)))
            .collect::<Vec<_>>()
    )
}

fn cell_signature(cell: &rdocx_oxml::table::CT_Tc) -> String {
    format!(
        "{:?}:{:?}",
        cell.content
            .iter()
            .map(|content| match content {
                CellContent::Paragraph(paragraph) => paragraph_signature(paragraph),
                CellContent::Table(table) => table_signature(table),
                CellContent::ContentControl(control) => control_signature(control),
            })
            .collect::<Vec<_>>(),
        cell.extra_xml
    )
}

fn control_signature(control: &CT_Sdt) -> String {
    let properties = control_property_signature(control);
    format!(
        "{:?}:{:?}",
        properties,
        control
            .content
            .iter()
            .filter_map(|content| match content {
                SdtContent::Paragraph(paragraph) => Some(paragraph_signature(paragraph)),
                SdtContent::Table(table) => Some(table_signature(table)),
                SdtContent::Row(row) => Some(row_signature(row)),
                SdtContent::Cell(cell) => Some(cell_signature(cell)),
                SdtContent::Run(run) => Some(run_signature(run)),
                SdtContent::ContentControl(control) => Some(control_signature(control)),
                SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace) => None,
                SdtContent::RawXml(raw) => Some(format!("{raw:?}")),
            })
            .collect::<Vec<_>>()
    )
}

fn control_property_signature(control: &CT_Sdt) -> ControlPropertySignature<'_> {
    control.properties.as_ref().map(|properties| {
        (
            properties.alias.as_deref(),
            properties.tag.as_deref(),
            properties.id,
            properties.control_type,
            properties.data_binding.as_ref(),
        )
    })
}

fn modeled_control_content(control: &CT_Sdt) -> Vec<&SdtContent> {
    control
        .content
        .iter()
        .filter(|content| {
            !matches!(content, SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace))
        })
        .collect()
}

fn control_whitespace_slots(control: &CT_Sdt) -> Result<Vec<String>> {
    let modeled_count = modeled_control_content(control).len();
    let mut slots = vec![String::new(); modeled_count + 1];
    let mut modeled_before = 0usize;
    for content in &control.content {
        match content {
            SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace) => {
                slots[modeled_before].push_str(std::str::from_utf8(raw).map_err(utf8_error)?);
            }
            _ => modeled_before += 1,
        }
    }
    Ok(slots)
}

fn normalized_body(document: &CT_Document) -> Vec<String> {
    document.body.content.iter().map(body_signature).collect()
}

fn normalized_body_with_options(
    document: &CT_Document,
    options: &ComparisonOptions,
) -> Vec<String> {
    if options == &ComparisonOptions::default() {
        return normalized_body(document);
    }
    document
        .body
        .content
        .iter()
        .map(|content| body_signature_with_options(content, options))
        .collect()
}

fn body_signature_with_options(content: &BodyContent, options: &ComparisonOptions) -> String {
    match content {
        BodyContent::Paragraph(paragraph) => {
            format!("p:{}", paragraph_signature_with_options(paragraph, options))
        }
        BodyContent::Table(table) => {
            format!("t:{}", table_signature_with_options(table, options))
        }
        BodyContent::ContentControl(control) => {
            format!("s:{}", control_signature_with_options(control, options))
        }
        BodyContent::RawXml(raw) => format!("x:{raw:?}"),
    }
}

fn paragraph_signature_with_options(paragraph: &CT_P, options: &ComparisonOptions) -> String {
    let numbering = (!options.ignore_formatting)
        .then(|| paragraph_numbering(paragraph))
        .flatten();
    let runs = if !uses_attributed_run_path(options) {
        paragraph
            .runs
            .iter()
            .map(|run| run_signature_with_options(run, options))
            .filter(|signature| !signature.is_empty())
            .collect::<Vec<_>>()
    } else {
        attributed_run_units(&paragraph.runs, options)
            .iter()
            .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
            .map(attributed_unit_signature)
            .collect::<Vec<_>>()
    };
    let comment_ranges = (!options.ignore_comments).then_some(&paragraph.comment_ranges);
    let hyperlinks = paragraph
        .hyperlinks
        .iter()
        .map(|link| {
            (
                &link.rel_id,
                &link.anchor,
                &link.tooltip,
                &link.doc_location,
                &link.extra_attributes,
                hyperlink_raw_boundaries(paragraph, link, options),
                link.preserved_raw_before,
                policy_run_boundary(paragraph, link.run_start, options),
                policy_run_boundary(paragraph, link.run_end, options),
            )
        })
        .collect::<Vec<_>>();
    format!(
        "{numbering:?}:{runs:?}:{:?}:{comment_ranges:?}:{:?}:{:?}:{:?}",
        hyperlinks,
        paragraph.bookmark_markers,
        paragraph.extra_xml,
        paragraph
            .content_controls
            .iter()
            .map(|(at, raw_before, markers_before, control)| (
                policy_run_boundary(paragraph, *at, options),
                raw_before,
                markers_before,
                control_signature_with_options(control, options),
            ))
            .collect::<Vec<_>>(),
    )
}

fn run_signature_with_options(run: &CT_R, options: &ComparisonOptions) -> String {
    if options == &ComparisonOptions::default() {
        return run_signature(run);
    }
    let content = run
        .content
        .iter()
        .filter_map(|content| {
            if options.ignore_comments && matches!(content, RunContent::CommentReference { .. }) {
                return None;
            }
            if options.ignore_fields && matches!(content, RunContent::Field(_)) {
                return Some("ignored:field".to_owned());
            }
            match content {
                RunContent::Text(text) | RunContent::DeletedText(text)
                    if options.ignore_whitespace =>
                {
                    let visible = text
                        .text
                        .chars()
                        .filter(|character| !character.is_whitespace())
                        .collect::<String>();
                    (!visible.is_empty()).then_some(format!("text:{visible:?}"))
                }
                _ => Some(run_content_signature(content)),
            }
        })
        .collect::<Vec<_>>();
    if content.is_empty() && run.extra_xml.is_empty() {
        String::new()
    } else {
        format!(
            "{content:?}:{:?}:{:?}",
            run.extra_xml, run.extra_xml_positions
        )
    }
}

fn table_signature_with_options(table: &CT_Tbl, options: &ComparisonOptions) -> String {
    format!(
        "{:?}:{:?}:{:?}:{:?}",
        table.grid,
        table
            .rows
            .iter()
            .map(|row| row_signature_with_options(row, options))
            .collect::<Vec<_>>(),
        table.extra_xml,
        table
            .content_controls
            .iter()
            .map(|(at, raw_before, control)| (
                at,
                raw_before,
                control_signature_with_options(control, options),
            ))
            .collect::<Vec<_>>()
    )
}

fn row_signature_with_options(row: &CT_Row, options: &ComparisonOptions) -> String {
    format!(
        "{:?}:{:?}:{:?}",
        row.cells
            .iter()
            .map(|cell| cell_signature_with_options(cell, options))
            .collect::<Vec<_>>(),
        row.extra_xml,
        row.content_controls
            .iter()
            .map(|(at, raw_before, control)| (
                at,
                raw_before,
                control_signature_with_options(control, options),
            ))
            .collect::<Vec<_>>()
    )
}

fn cell_signature_with_options(cell: &CT_Tc, options: &ComparisonOptions) -> String {
    format!(
        "{:?}:{:?}",
        cell.content
            .iter()
            .map(|content| match content {
                CellContent::Paragraph(paragraph) => {
                    paragraph_signature_with_options(paragraph, options)
                }
                CellContent::Table(table) => table_signature_with_options(table, options),
                CellContent::ContentControl(control) => {
                    control_signature_with_options(control, options)
                }
            })
            .collect::<Vec<_>>(),
        cell.extra_xml
    )
}

fn control_signature_with_options(control: &CT_Sdt, options: &ComparisonOptions) -> String {
    let mut content = Vec::new();
    for child in &control.content {
        match child {
            SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace) => {}
            SdtContent::Run(run) if uses_attributed_run_path(options) => {
                content.extend(
                    attributed_run_units(std::slice::from_ref(run), options)
                        .into_iter()
                        .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
                        .map(|unit| format!("u:{}", attributed_unit_signature(&unit))),
                );
            }
            child => content.push(control_content_signature_with_options(child, options)),
        }
    }
    format!("{:?}:{:?}", control_property_signature(control), content)
}

fn control_content_signature_with_options(
    content: &SdtContent,
    options: &ComparisonOptions,
) -> String {
    match content {
        SdtContent::Paragraph(paragraph) => {
            format!("p:{}", paragraph_signature_with_options(paragraph, options))
        }
        SdtContent::Table(table) => {
            format!("t:{}", table_signature_with_options(table, options))
        }
        SdtContent::Row(row) => format!("r:{}", row_signature_with_options(row, options)),
        SdtContent::Cell(cell) => format!("c:{}", cell_signature_with_options(cell, options)),
        SdtContent::Run(run) => format!("u:{}", run_signature_with_options(run, options)),
        SdtContent::ContentControl(control) => {
            format!("s:{}", control_signature_with_options(control, options))
        }
        SdtContent::RawXml(raw) => format!("x:{raw:?}"),
    }
}

fn paragraph_formatting(paragraph: &CT_P) -> Option<CT_PPr> {
    paragraph.properties.clone().map(|mut properties| {
        properties.num_id = None;
        properties.num_ilvl = None;
        properties.numbering_revision = None;
        properties.numbering_revision_xml.clear();
        properties.change = None;
        properties.revision_xml.clear();
        properties
    })
}

fn row_formatting(row: &CT_Row) -> Option<CT_TrPr> {
    row.properties.clone().map(|mut properties| {
        properties.revision_markers.clear();
        properties.revision_xml.clear();
        properties
    })
}

fn body_content_xml(content: &BodyContent) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => paragraph_xml(paragraph),
        BodyContent::Table(table) => table_xml(table),
        BodyContent::ContentControl(control) => control_xml(control),
        BodyContent::RawXml(raw) => String::from_utf8(raw.clone()).map_err(utf8_error),
    }
}

fn paragraph_xml(paragraph: &CT_P) -> Result<String> {
    let mut bytes = Vec::new();
    paragraph.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn property_xml(properties: &CT_PPr) -> Result<String> {
    let mut bytes = Vec::new();
    properties.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn run_xml(run: &CT_R) -> Result<String> {
    let mut bytes = Vec::new();
    run.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn paragraph_owned_run_xml(run: &CT_R) -> Result<String> {
    if !run_is_field(run) {
        return run_xml(run);
    }
    let mut paragraph = CT_P::new();
    paragraph.runs.push(run.clone());
    let xml = paragraph_xml(&paragraph)?;
    let start = xml
        .find('>')
        .ok_or_else(|| Error::Other("serialized paragraph has no start".to_owned()))?
        + 1;
    let end = xml
        .rfind("</w:p>")
        .ok_or_else(|| Error::Other("serialized paragraph has no end".to_owned()))?;
    Ok(xml[start..end].to_owned())
}

fn policy_inserted_run_xml(
    edited: &CT_R,
    original: Option<&CT_R>,
    options: &ComparisonOptions,
) -> Result<String> {
    if !options.ignore_formatting {
        return paragraph_owned_run_xml(edited);
    }
    let mut inserted = edited.clone();
    if let Some(original) = original {
        inserted.properties = original.properties.clone();
    }
    paragraph_owned_run_xml(&inserted)
}

fn run_is_field(run: &CT_R) -> bool {
    matches!(run.content.as_slice(), [RunContent::Field(_)])
}

fn run_field(run: &CT_R) -> Option<&rdocx_oxml::text::Field> {
    match run.content.as_slice() {
        [RunContent::Field(field)] => Some(field),
        _ => None,
    }
}

fn validate_field_alignment(
    aligned: &[(Option<usize>, Option<usize>)],
    original: &[CT_R],
    edited: &[CT_R],
    original_signatures: &[String],
    edited_signatures: &[String],
    location: &str,
) -> Result<()> {
    for (left, right) in aligned {
        let left_field = left.is_some_and(|index| run_is_field(&original[index]));
        let right_field = right.is_some_and(|index| run_is_field(&edited[index]));
        let unchanged_pair = match (*left, *right) {
            (Some(i), Some(j)) => original_signatures[i] == edited_signatures[j],
            _ => false,
        };
        if (left_field || right_field) && !unchanged_pair {
            return Err(Error::Other(format!(
                "comparison cannot revise a modeled field at {location}"
            )));
        }
    }
    Ok(())
}

fn deleted_run_xml(run: &CT_R) -> Result<String> {
    let mut deleted = run.clone();
    for content in &mut deleted.content {
        if let RunContent::Text(text) = content {
            *content = RunContent::DeletedText(text.clone());
        }
    }
    run_xml(&deleted)
}

fn table_xml(table: &CT_Tbl) -> Result<String> {
    let mut bytes = Vec::new();
    table.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn row_xml(row: &CT_Row) -> Result<String> {
    let mut bytes = Vec::new();
    row.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn cell_xml(cell: &CT_Tc) -> Result<String> {
    let mut bytes = Vec::new();
    cell.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn control_xml(control: &CT_Sdt) -> Result<String> {
    let document = CT_Document {
        body: rdocx_oxml::document::CT_Body {
            content: vec![BodyContent::ContentControl(control.clone())],
            sect_pr: None,
        },
        ..CT_Document::new()
    };
    let xml = document.to_xml()?;
    extract_body_inner(&xml).map(str::to_owned)
}

fn section_property_xml(section: &rdocx_oxml::document::CT_SectPr) -> Result<String> {
    let mut bytes = Vec::new();
    section.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn replace_body_inner(document: &[u8], body: &str) -> Result<String> {
    let xml = std::str::from_utf8(document).map_err(utf8_error)?;
    let start = xml
        .find("<w:body")
        .and_then(|at| xml[at..].find('>').map(|offset| at + offset + 1))
        .ok_or_else(|| Error::Other("serialized document has no w:body start".to_owned()))?;
    let end = xml
        .rfind("</w:body>")
        .ok_or_else(|| Error::Other("serialized document has no w:body end".to_owned()))?;
    Ok(format!("{}{}{}", &xml[..start], body, &xml[end..]))
}

fn extract_body_inner(document: &[u8]) -> Result<&str> {
    let xml = std::str::from_utf8(document).map_err(utf8_error)?;
    let start = xml
        .find("<w:body")
        .and_then(|at| xml[at..].find('>').map(|offset| at + offset + 1))
        .ok_or_else(|| Error::Other("serialized document has no w:body start".to_owned()))?;
    let end = xml
        .rfind("</w:body>")
        .ok_or_else(|| Error::Other("serialized document has no w:body end".to_owned()))?;
    Ok(&xml[start..end])
}

fn replace_element_inner(xml: &str, name: &str, inner: &str) -> Result<String> {
    let opening = format!("<{name}");
    let closing = format!("</{name}>");
    let start = xml
        .find(&opening)
        .and_then(|at| xml[at..].find('>').map(|offset| at + offset + 1))
        .ok_or_else(|| Error::Other(format!("serialized XML has no {name} start")))?;
    let end = xml
        .rfind(&closing)
        .ok_or_else(|| Error::Other(format!("serialized XML has no {name} end")))?;
    Ok(format!("{}{}{}", &xml[..start], inner, &xml[end..]))
}

fn inject_before_close(xml: &str, closing: &str, addition: &str) -> Result<String> {
    let at = xml
        .rfind(closing)
        .ok_or_else(|| Error::Other(format!("serialized XML has no {closing}")))?;
    Ok(format!("{}{}{}", &xml[..at], addition, &xml[at..]))
}

fn append_word_child(xml: &str, local: &str, addition: &str) -> Result<String> {
    let closing = format!("</w:{local}>");
    if xml.ends_with(&closing) {
        return inject_before_close(xml, &closing, addition);
    }
    let Some(slash) = xml.rfind("/>") else {
        return Err(Error::Other(format!(
            "serialized w:{local} has no closing boundary"
        )));
    };
    Ok(format!(
        "{}>{addition}</w:{local}>{}",
        &xml[..slash],
        &xml[slash + 2..]
    ))
}

fn direct_word_element_spans(xml: &str, local: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let expected = format!("w:{local}");
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let target = depth == 1 && element.name().as_ref() == expected.as_bytes();
                open.push(target.then_some(before));
                depth += 1;
            }
            Event::Empty(element) => {
                if depth == 1 && element.name().as_ref() == expected.as_bytes() {
                    spans.push(before..after);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open.pop() {
                    spans.push(start..after);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn paragraph_close_start(xml: &str) -> Result<usize> {
    xml.rfind("</w:p>")
        .ok_or_else(|| Error::Other("serialized paragraph has no end".to_owned()))
}

fn paragraph_run_spans(xml: &str) -> Result<Vec<Range<usize>>> {
    let open = format!(r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:w="{W_NS}">"#);
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut stack = Vec::<(Vec<u8>, Option<usize>)>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                let parent = stack.last().map(|(local, _)| local.as_slice());
                let target = is_word
                    && matches!(local.as_slice(), b"r" | b"fldSimple")
                    && parent.is_some_and(|value| value == b"p" || value == b"hyperlink");
                stack.push((local, target.then_some(before)));
            }
            Event::Empty(element) => {
                let local = element.local_name();
                let parent = stack.last().map(|(local, _)| local.as_slice());
                if is_word
                    && matches!(local.as_ref(), b"r" | b"fldSimple")
                    && parent.is_some_and(|value| value == b"p" || value == b"hyperlink")
                {
                    spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::End(_) => {
                if let Some((_, Some(start))) = stack.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    spans.sort_by_key(|span| span.start);
    Ok(spans)
}

fn replace_paragraph_run_elements(xml: &str, replacements: &[String]) -> Result<String> {
    let spans = paragraph_run_spans(xml)?;
    if spans.len() != replacements.len() {
        return Err(Error::Other(format!(
            "comparison expected {} paragraph run owners, found {}",
            replacements.len(),
            spans.len()
        )));
    }
    let mut output = xml.to_owned();
    for (span, replacement) in spans.into_iter().zip(replacements).rev() {
        output.replace_range(span, replacement);
    }
    Ok(output)
}

fn replace_direct_word_elements(xml: &str, local: &str, replacements: &[String]) -> Result<String> {
    let spans = direct_word_element_spans(xml, local)?;
    if spans.len() != replacements.len() {
        return Err(Error::Other(format!(
            "comparison expected {} direct w:{local} elements, found {}",
            replacements.len(),
            spans.len()
        )));
    }
    let mut output = xml.to_owned();
    for (span, replacement) in spans.into_iter().zip(replacements).rev() {
        output.replace_range(span, replacement);
    }
    Ok(output)
}

fn cell_child_spans(
    cell: &CT_Tc,
    xml: &str,
    location: &str,
) -> Result<Vec<(String, Range<usize>)>> {
    let candidates = direct_modeled_child_spans(xml)?;
    let mut spans = Vec::with_capacity(cell.content.len());
    let mut candidate_index = 0usize;
    for (index, child) in cell.content.iter().enumerate() {
        for (_, raw) in cell.extra_xml.iter().filter(|(at, _)| *at == index) {
            let raw = std::str::from_utf8(raw).map_err(utf8_error)?;
            if candidates
                .get(candidate_index)
                .is_some_and(|(_, span)| &xml[span.clone()] == raw)
            {
                candidate_index += 1;
            }
        }
        let expected_name = match child {
            CellContent::Paragraph(_) => "w:p",
            CellContent::Table(_) => "w:tbl",
            CellContent::ContentControl(_) => "w:sdt",
        };
        let (name, span) = candidates.get(candidate_index).ok_or_else(|| {
            Error::Other(format!(
                "comparison could not correlate serialized cell content at {location}"
            ))
        })?;
        if name != expected_name {
            return Err(Error::Other(format!(
                "comparison found {name} instead of {expected_name} at {location}"
            )));
        }
        spans.push((name.clone(), span.clone()));
        candidate_index += 1;
    }
    Ok(spans)
}

fn direct_modeled_child_spans(xml: &str) -> Result<Vec<(String, Range<usize>)>> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open: Vec<Option<(String, usize)>> = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let name = std::str::from_utf8(element.name().as_ref())
                    .map_err(utf8_error)?
                    .to_owned();
                let modeled = depth == 1 && matches!(name.as_str(), "w:p" | "w:tbl" | "w:sdt");
                open.push(modeled.then_some((name, before)));
                depth += 1;
            }
            Event::Empty(element) => {
                let name = std::str::from_utf8(element.name().as_ref())
                    .map_err(utf8_error)?
                    .to_owned();
                if depth == 1 && matches!(name.as_str(), "w:p" | "w:tbl" | "w:sdt") {
                    spans.push((name, before..after));
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some((name, start))) = open.pop() {
                    spans.push((name, start..after));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn replace_ranges(
    xml: &str,
    spans: &[(String, Range<usize>)],
    replacements: &[String],
) -> Result<String> {
    if spans.len() != replacements.len() {
        return Err(Error::Other(
            "comparison replacement count does not match modeled children".to_owned(),
        ));
    }
    let mut output = xml.to_owned();
    for ((_, span), replacement) in spans.iter().zip(replacements).rev() {
        output.replace_range(span.clone(), replacement);
    }
    Ok(output)
}

fn word_ids(xml: &[u8]) -> Result<HashSet<i32>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut ids = HashSet::new();
    let mut buffer = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if local.as_ref() == b"id"
                        && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
                        && let Ok(value) = std::str::from_utf8(&attribute.value)
                        && let Ok(id) = value.parse::<i32>()
                    {
                        ids.insert(id);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(ids)
}

fn utf8_error(error: impl std::fmt::Display) -> Error {
    Error::Other(format!("comparison XML is not UTF-8: {error}"))
}

#[cfg(test)]
mod tests {
    use super::{
        ComparisonGranularity, ComparisonOptions, FAIL_AFTER_COMPARISON_STAGING,
        attributed_run_units, word_fragments,
    };
    use crate::Document;
    use rdocx_oxml::text::{BreakType, CT_R, CT_Text, RunContent};

    #[test]
    fn comparison_defaults_preserve_run_granularity() {
        let mut legacy = Document::new();
        legacy.add_paragraph("before");
        let mut explicit = legacy.clone_for_staging();
        let mut edited = Document::new();
        edited.add_paragraph("after");
        let legacy_diagnostics = legacy
            .compare(&edited, "Ada", "2026-09-04T09:00:00Z")
            .unwrap();
        let explicit_diagnostics = explicit
            .compare_with_options(
                &edited,
                "Ada",
                "2026-09-04T09:00:00Z",
                &ComparisonOptions::default(),
            )
            .unwrap();
        assert_eq!(legacy_diagnostics, explicit_diagnostics);
        assert_eq!(legacy.to_bytes().unwrap(), explicit.to_bytes().unwrap());
        assert_eq!(
            ComparisonOptions::default().granularity,
            ComparisonGranularity::Run
        );
    }

    #[test]
    fn word_and_character_granularity_split_only_text_content() {
        assert_eq!(
            word_fragments("élan_東京  \t—!?"),
            ["élan_東京", "  \t", "—!?"].map(str::to_owned)
        );
        let run = CT_R {
            properties: None,
            content: vec![
                RunContent::Text(CT_Text::new("A😀")),
                RunContent::Tab,
                RunContent::Break(BreakType::Page),
            ],
            extra_xml: vec![b"<x:raw xmlns:x=\"urn:f235\"/>".to_vec()],
            extra_xml_positions: vec![3],
            alt_drawings: Vec::new(),
        };
        let units = attributed_run_units(
            &[run],
            &ComparisonOptions {
                granularity: ComparisonGranularity::Character,
                ..Default::default()
            },
        );
        assert_eq!(units.len(), 4);
        assert!(matches!(&units[0].run.content[..], [RunContent::Text(text)] if text.text == "A"));
        assert!(matches!(&units[1].run.content[..], [RunContent::Text(text)] if text.text == "😀"));
        assert!(matches!(&units[2].run.content[..], [RunContent::Tab]));
        assert!(matches!(
            &units[3].run.content[..],
            [RunContent::Break(BreakType::Page)]
        ));
        assert_eq!(
            units
                .iter()
                .map(|unit| unit.run.extra_xml.len())
                .sum::<usize>(),
            1
        );
    }

    #[test]
    fn staged_comparison_postcondition_failure_preserves_bytes_and_layout_cache() {
        let mut original = Document::new();
        original.add_paragraph("original");
        let mut edited = Document::new();
        edited.add_paragraph("edited");
        let before = original.to_bytes().unwrap();
        let layout = original.layout().unwrap();
        FAIL_AFTER_COMPARISON_STAGING.with(|fail| fail.set(true));
        let error = original
            .compare_with_options(
                &edited,
                "Ada",
                "2026-09-04T09:00:00Z",
                &ComparisonOptions::default(),
            )
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("staged comparison postcondition")
        );
        assert_eq!(original.to_bytes().unwrap(), before);
        assert!(std::sync::Arc::ptr_eq(&layout, &original.layout().unwrap()));
    }
}
