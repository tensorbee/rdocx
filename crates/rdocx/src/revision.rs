//! Native tracked-revision inspection and atomic resolution.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
pub use rdocx_oxml::RevisionKind;
use rdocx_oxml::text::CT_P;
use rdocx_oxml::{CT_Document, CT_Revision};
use sha2::{Digest, Sha256};

use crate::{Document, Error, ParagraphRef, Result, StoryId};

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type NamespaceScope = HashMap<String, String>;
type NamespaceDeclarations = Vec<(String, String)>;
type ResolvedStoryLeaves = (Vec<u8>, usize, Vec<Option<Range<usize>>>);

/// Whether a tracked revision takes its text out of the accepted view, as a
/// deletion or a move away does. The exporters leave that text out.
pub(crate) fn revision_removes_text(revision: &CT_Revision) -> bool {
    matches!(
        revision.kind(),
        RevisionKind::Deletion | RevisionKind::MoveFrom
    )
}

/// The content of a tracked insertion or move in, which the exporters write in
/// place. None for an empty one, and for a deletion, a move away or a property
/// change.
pub(crate) fn accepted_revision_content(revision: &CT_Revision) -> Option<&CT_P> {
    matches!(
        revision.kind(),
        RevisionKind::Insertion | RevisionKind::MoveTo
    )
    .then(|| revision.content_paragraph())
    .flatten()
}

/// Whether the preserved paragraph child `paragraph.extra_xml[index]` is the
/// source of one of the typed `paragraph.revisions`, which the exporters read
/// through the revision instead.
pub(crate) fn raw_is_typed_revision(paragraph: &CT_P, index: usize) -> bool {
    let run_index = paragraph.extra_xml[index].0;
    let slot = paragraph.extra_xml[..index]
        .iter()
        .filter(|(at, _)| *at == run_index)
        .count();
    paragraph
        .revisions
        .iter()
        .any(|(at, revision_slot, _)| *at == run_index && *revision_slot == slot)
}

/// An immutable view of one tracked revision in the main document.
#[derive(Debug, Clone, Copy)]
pub struct RevisionRef<'a> {
    pub(crate) inner: &'a CT_Revision,
}

impl RevisionRef<'_> {
    pub fn id(&self) -> i32 {
        self.inner.id()
    }

    pub fn author(&self) -> &str {
        self.inner.author()
    }

    pub fn timestamp(&self) -> Option<&str> {
        self.inner.timestamp()
    }

    pub fn kind(&self) -> RevisionKind {
        self.inner.kind()
    }

    /// Return the insertion content as a paragraph reader projection.
    ///
    /// This retains runs, hyperlinks, fields, and preserved XML in their
    /// original inline order. Non-insertion revisions return `None`.
    pub fn insertion_paragraph(&self) -> Option<ParagraphRef<'_>> {
        if !matches!(
            self.inner.kind(),
            RevisionKind::Insertion | RevisionKind::MoveTo
        ) {
            return None;
        }
        self.inner
            .content_paragraph()
            .map(|inner| ParagraphRef { inner })
    }
}

/// One owned tracked revision and the Word story that holds it.
///
/// [`Document::story_revisions`] returns one snapshot for each revision
/// element that accepting or rejecting every revision resolves.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryRevision {
    story: StoryId,
    id: i32,
    author: String,
    timestamp: Option<String>,
    kind: RevisionKind,
    text: String,
}

fn comment_comparison_record(
    document: &Document,
) -> Result<Option<(serde_json::Value, RevisionMetadata)>> {
    let Some(raw) = document
        .package
        .get_part(crate::comparison::COMMENT_COMPARISON_PART)
    else {
        return Ok(None);
    };
    use base64::Engine;
    let encoded = std::str::from_utf8(raw)
        .map_err(|error| Error::Other(error.to_string()))?
        .strip_prefix(crate::comparison::COMMENT_COMPARISON_OPEN)
        .and_then(|xml| xml.strip_suffix(crate::comparison::COMMENT_COMPARISON_CLOSE))
        .ok_or_else(|| Error::Other("invalid comment comparison XML".to_owned()))?;
    let data = base64::engine::general_purpose::STANDARD
        .decode(encoded)
        .map_err(|error| Error::Other(error.to_string()))?;
    let record: serde_json::Value =
        serde_json::from_slice(&data).map_err(|error| Error::Other(error.to_string()))?;
    let metadata = RevisionMetadata {
        kind: RevisionKind::Insertion,
        id: record
            .get("id")
            .and_then(serde_json::Value::as_i64)
            .and_then(|id| i32::try_from(id).ok())
            .ok_or_else(|| Error::Other("invalid comment revision id".to_owned()))?,
        author: record
            .get("author")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| Error::Other("invalid comment revision author".to_owned()))?
            .to_owned(),
        timestamp: Some(
            record
                .get("timestamp")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| Error::Other("invalid comment revision date".to_owned()))?
                .to_owned(),
        ),
    };
    Ok(Some((record, metadata)))
}

pub(crate) fn comment_comparison_revision(
    document: &Document,
    story: StoryId,
) -> Result<Option<StoryRevision>> {
    Ok(
        comment_comparison_record(document)?.map(|(_, metadata)| StoryRevision {
            story,
            id: metadata.id,
            author: metadata.author,
            timestamp: metadata.timestamp,
            kind: metadata.kind,
            text: String::new(),
        }),
    )
}

impl StoryRevision {
    /// Stored visible text of this revision, or its current property owner.
    ///
    /// This owned snapshot does not evaluate fields or change with later edits.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return the story that holds this revision.
    ///
    /// Table cells fold into the story that holds the table. The identity is
    /// one of those returned by [`Document::stories`] for the same document.
    pub fn story(&self) -> &StoryId {
        &self.story
    }

    pub fn id(&self) -> i32 {
        self.id
    }

    pub fn author(&self) -> &str {
        &self.author
    }

    pub fn timestamp(&self) -> Option<&str> {
        self.timestamp.as_deref()
    }

    pub fn kind(&self) -> RevisionKind {
        self.kind
    }
}

#[derive(Clone, Copy)]
pub(crate) enum Resolution {
    Accept,
    Reject,
}

#[derive(Clone)]
pub(crate) enum RevisionScope<'a> {
    All,
    Author(&'a str),
    Id(i32),
    Ids(&'a HashSet<i32>),
    DateRange { start: Instant, end: Instant },
}

#[derive(Clone, PartialEq, Eq)]
pub(crate) struct RevisionMetadata {
    pub(crate) kind: RevisionKind,
    pub(crate) id: i32,
    pub(crate) author: String,
    pub(crate) timestamp: Option<String>,
}

pub(crate) struct XmlElement {
    pub(crate) start: usize,
    pub(crate) open_end: usize,
    pub(crate) close_start: usize,
    pub(crate) end: usize,
    pub(crate) name: String,
    pub(crate) local: String,
    pub(crate) word: bool,
    pub(crate) modeled: bool,
    pub(crate) empty: bool,
    pub(crate) parent: Option<usize>,
    pub(crate) children: Vec<usize>,
    pub(crate) revision: Option<RevisionMetadata>,
    pub(crate) namespace_declarations: NamespaceDeclarations,
    field_marker: Option<String>,
    snapshot_field_marker: Option<String>,
    snapshot_scalar: String,
}

/// Read-projection facts share the existing tree, never a second parsed model.
struct SnapshotFacts {
    nodes: Vec<SnapshotNode>,
    fields: Vec<SnapshotField>,
    #[cfg(test)]
    preparation_visits: usize,
    #[cfg(test)]
    owner_visits: std::cell::Cell<usize>,
}

#[derive(Clone, Copy)]
struct SnapshotNode {
    story: usize,
    domain: usize,
    domain_end: usize,
    canonical_w: bool,
    control_parent: Option<usize>,
    admitted: bool,
    wrappers: bool,
    hidden: bool,
    row_removed: bool,
    seed: Option<usize>,
}

#[derive(Clone, Copy)]
struct SnapshotField {
    previous: Option<usize>,
    domain: usize,
    end: usize,
    instructions: usize,
    result: bool,
}

pub(crate) struct XmlTree<'a> {
    source: &'a [u8],
    pub(crate) elements: Vec<XmlElement>,
    pub(crate) root: usize,
    note_admission: Option<Vec<bool>>,
}

struct RenderState<'a> {
    resolution: Resolution,
    scope: RevisionScope<'a>,
    resolved: HashSet<usize>,
    /// True while rendering content that the resolution removes, such as a
    /// removed row. Its paragraph marks need no merge partner.
    discarded: bool,
    leaf_elements: Vec<usize>,
    leaf_ranges: Vec<Option<Range<usize>>>,
    leaf_updates: Vec<usize>,
    retiring_simple_fields: HashSet<usize>,
}

impl RenderState<'_> {
    fn shift_leaves(&mut self, from: usize, offset: usize) {
        for &index in &self.leaf_updates[from..] {
            if let Some(range) = &mut self.leaf_ranges[index] {
                range.start += offset;
                range.end += offset;
            }
        }
    }
}

/// Operation-private source transport for the comparison Field publication.
/// Ordinary public revision resolution never creates or carries this observer.
pub(crate) struct NotePublication {
    pub(crate) part_name: String,
    pub(crate) revision_id: i32,
    pub(crate) roots: Vec<Option<Range<usize>>>,
    pub(crate) output: Option<Vec<u8>>,
    source_hash: Option<[u8; 32]>,
}

impl NotePublication {
    pub(crate) fn new(part_name: &str, revision_id: i32, roots: &[Range<usize>]) -> Self {
        Self {
            part_name: part_name.to_owned(),
            revision_id,
            roots: roots.iter().cloned().map(Some).collect(),
            output: None,
            source_hash: None,
        }
    }

    fn validate_source(&self, source: &[u8]) -> Result<()> {
        if self.source_hash != Some(Sha256::digest(source).into()) {
            return Err(Error::Other(
                "note publication transport lost its actual source phase".into(),
            ));
        }
        Ok(())
    }

    pub(crate) fn remove_ranges(
        &mut self,
        part: &str,
        source: &[u8],
        ranges: &[Range<usize>],
    ) -> Result<()> {
        if self.part_name != part {
            return Ok(());
        }
        self.validate_source(source)?;
        let mut previous = 0;
        for range in ranges {
            if range.start >= range.end || range.end > source.len() || range.start < previous {
                return Err(Error::Other(
                    "note publication edits overlap or leave their actual source".into(),
                ));
            }
            previous = range.end;
        }
        for range in self.roots.iter().flatten() {
            if range.start >= range.end || range.end > source.len() {
                return Err(Error::Other(
                    "note publication root is outside its actual source".into(),
                ));
            }
            for edit in ranges {
                if edit.start < range.end
                    && edit.end > range.start
                    && !(edit.start <= range.start && edit.end >= range.end)
                {
                    return Err(Error::Other(
                        "note publication root partially overlaps a retired source owner".into(),
                    ));
                }
            }
        }
        for root in &mut self.roots {
            let Some(range) = root else {
                continue;
            };
            if range.start >= range.end || range.end > source.len() {
                return Err(Error::Other(
                    "note publication root is outside its actual source".into(),
                ));
            }
            let mut removed = false;
            let mut shift = 0;
            for edit in ranges {
                if edit.start <= range.start && edit.end >= range.end {
                    removed = true;
                    break;
                }
                if edit.start < range.end && edit.end > range.start {
                    return Err(Error::Other(
                        "note publication root partially overlaps a retired source owner".into(),
                    ));
                }
                if edit.end <= range.start {
                    shift += edit.len();
                }
            }
            *root = if removed {
                None
            } else {
                Some(range.start - shift..range.end - shift)
            };
        }
        let mut hash = Sha256::new();
        let mut cursor = 0;
        for range in ranges {
            hash.update(&source[cursor..range.start]);
            cursor = range.end;
        }
        hash.update(&source[cursor..]);
        self.source_hash = Some(hash.finalize().into());
        Ok(())
    }
}

impl Document {
    /// Accept every modeled revision in the main document and related stories.
    pub fn accept_all(&mut self) -> Result<usize> {
        self.resolve_revisions(Resolution::Accept, RevisionScope::All)
    }

    /// Reject every modeled revision in the main document and related stories.
    pub fn reject_all(&mut self) -> Result<usize> {
        self.resolve_revisions(Resolution::Reject, RevisionScope::All)
    }

    /// Accept every modeled revision written by `author`.
    pub fn accept_revisions_by_author(&mut self, author: &str) -> Result<usize> {
        self.resolve_revisions(Resolution::Accept, RevisionScope::Author(author))
    }

    /// Reject every modeled revision written by `author`.
    pub fn reject_revisions_by_author(&mut self, author: &str) -> Result<usize> {
        self.resolve_revisions(Resolution::Reject, RevisionScope::Author(author))
    }

    /// Accept every modeled revision in the inclusive RFC 3339 instant range.
    pub fn accept_revisions_in_date_range(&mut self, start: &str, end: &str) -> Result<usize> {
        let scope = date_scope(start, end)?;
        self.resolve_revisions(Resolution::Accept, scope)
    }

    /// Reject every modeled revision in the inclusive RFC 3339 instant range.
    pub fn reject_revisions_in_date_range(&mut self, start: &str, end: &str) -> Result<usize> {
        let scope = date_scope(start, end)?;
        self.resolve_revisions(Resolution::Reject, scope)
    }

    /// Accept every modeled revision element carrying `id`.
    pub fn accept_revision_id(&mut self, id: i32) -> Result<usize> {
        self.resolve_revisions(Resolution::Accept, RevisionScope::Id(id))
    }

    /// Reject every modeled revision element carrying `id`.
    pub fn reject_revision_id(&mut self, id: i32) -> Result<usize> {
        self.resolve_revisions(Resolution::Reject, RevisionScope::Id(id))
    }

    /// Accept every modeled revision element whose `w:id` is in `ids`.
    pub(crate) fn accept_revision_ids(&mut self, ids: &HashSet<i32>) -> Result<usize> {
        self.resolve_revisions(Resolution::Accept, RevisionScope::Ids(ids))
    }

    /// Reject every modeled revision element whose `w:id` is in `ids`.
    pub(crate) fn reject_revision_ids(&mut self, ids: &HashSet<i32>) -> Result<usize> {
        self.resolve_revisions(Resolution::Reject, RevisionScope::Ids(ids))
    }

    fn resolve_revisions(
        &mut self,
        resolution: Resolution,
        scope: RevisionScope<'_>,
    ) -> Result<usize> {
        self.resolve_revisions_with_note_publication(resolution, scope, None)
    }

    pub(crate) fn resolve_revisions_with_note_publication(
        &mut self,
        resolution: Resolution,
        scope: RevisionScope<'_>,
        mut publication: Option<&mut NotePublication>,
    ) -> Result<usize> {
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        let footnote_custody = crate::comparison::validate_footnote_custody(&candidate, &scope)?;
        if let Some(trace) = publication.as_deref_mut() {
            if candidate.footnotes_part_name.as_deref() != Some(trace.part_name.as_str()) {
                return Err(Error::Other(
                    "publication trace has no positively discovered note part".into(),
                ));
            }
            let bytes = candidate
                .package
                .get_part(&trace.part_name)
                .ok_or_else(|| Error::Other("publication note part is missing".into()))?;
            let roots = trace
                .roots
                .iter()
                .cloned()
                .collect::<Option<Vec<_>>>()
                .ok_or_else(|| Error::Other("publication starts with a missing root".into()))?;
            let mut tree = XmlTree::parse(bytes)?;
            tree.qualify_note_revisions()?;
            publication_root_indices(&tree, resolution, &scope, trace.revision_id, &roots)?;
            trace.source_hash = Some(Sha256::digest(bytes).into());
        }
        let source = candidate
            .package
            .get_part(&candidate.doc_part_name)
            .ok_or_else(|| {
                Error::Other(format!(
                    "missing revision story part {}",
                    candidate.doc_part_name
                ))
            })?
            .to_vec();
        let tree = XmlTree::parse(&source)?;
        let mut state = RenderState {
            resolution,
            scope: scope.clone(),
            resolved: HashSet::new(),
            discarded: false,
            leaf_elements: Vec::new(),
            leaf_ranges: Vec::new(),
            leaf_updates: Vec::new(),
            retiring_simple_fields: HashSet::new(),
        };
        state.retiring_simple_fields = tree.carried_simple_field_retirements(&state)?;
        let mut output = Vec::with_capacity(source.len());
        output.extend_from_slice(&source[..tree.elements[tree.root].start]);
        output.extend_from_slice(&tree.render(tree.root, &mut state, false)?);
        output.extend_from_slice(&source[tree.elements[tree.root].end..]);

        let mut resolved = state.resolved.len();
        if resolved > 0 {
            let staged = CT_Document::from_xml(&output)?;
            staged.to_xml()?;
            candidate.document = staged;
            candidate.package.set_part(&candidate.doc_part_name, output);
        }

        let removed_note_revisions = crate::comparison::remove_selected_footnote_definitions(
            &mut candidate,
            &footnote_custody,
            matches!(resolution, Resolution::Accept),
            publication.as_deref_mut(),
        )?;
        resolved = resolved
            .checked_add(removed_note_revisions)
            .ok_or_else(|| Error::Other("resolved revision count overflowed".into()))?;
        let story_parts = crate::comparison::revision_story_parts(&candidate)?;
        for (_, part_name) in story_parts {
            let part = candidate
                .package
                .get_part(&part_name)
                .ok_or_else(|| Error::Other(format!("missing revision story part {part_name}")))?
                .to_vec();
            let (updated, count) = if let Some(trace) = publication
                .as_deref_mut()
                .filter(|trace| trace.part_name == part_name)
            {
                trace.validate_source(&part)?;
                let active = trace
                    .roots
                    .iter()
                    .enumerate()
                    .filter_map(|(index, root)| root.clone().map(|root| (index, root)))
                    .collect::<Vec<_>>();
                let (updated, count, mapped) = if active.is_empty() {
                    let (updated, count) = resolve_story_xml(&part, resolution, scope.clone())?;
                    (updated, count, Vec::new())
                } else {
                    resolve_story_xml_with_publication_roots(
                        &part,
                        resolution,
                        scope.clone(),
                        trace.revision_id,
                        &active
                            .iter()
                            .map(|(_, root)| root.clone())
                            .collect::<Vec<_>>(),
                    )?
                };
                if mapped.len() != active.len() {
                    return Err(Error::Other(
                        "publication renderer changed its root inventory".into(),
                    ));
                }
                for ((index, _), root) in active.into_iter().zip(mapped) {
                    trace.roots[index] = root;
                }
                trace.source_hash = Some(Sha256::digest(&updated).into());
                (updated, count)
            } else {
                resolve_story_xml(&part, resolution, scope.clone())?
            };
            if count > 0 {
                candidate.package.set_part(&part_name, updated);
                resolved = resolved
                    .checked_add(count)
                    .ok_or_else(|| Error::Other("resolved revision count overflowed".to_owned()))?;
            }
        }
        if let Some((record, metadata)) = comment_comparison_record(&candidate)?
            && scope.matches(&metadata)
        {
            if matches!(resolution, Resolution::Reject) {
                if let Some(ids) = record
                    .get("added_reply_ids")
                    .and_then(serde_json::Value::as_array)
                {
                    for id in ids.iter().rev() {
                        let id = id
                            .as_i64()
                            .and_then(|id| i32::try_from(id).ok())
                            .ok_or_else(|| {
                                Error::Other("invalid comparison reply identity".into())
                            })?;
                        candidate.remove_comment(id)?;
                    }
                }
                crate::comparison::apply_comment_snapshot(
                    &mut candidate,
                    record.get("original").ok_or_else(|| {
                        Error::Other("missing original comment snapshot".to_owned())
                    })?,
                )?;
            }
            candidate
                .package
                .remove_part(crate::comparison::COMMENT_COMPARISON_PART);
            candidate
                .package
                .content_types
                .remove_override(crate::comparison::COMMENT_COMPARISON_PART);
            if let Some(relationships) = candidate
                .package
                .get_part_rels_mut(&candidate.doc_part_name)
            {
                relationships.items.retain(|relationship| {
                    relationship.rel_type != crate::comparison::COMMENT_COMPARISON_REL
                        || oxml_opc::OpcPackage::resolve_rel_target(
                            &candidate.doc_part_name,
                            &relationship.target,
                        ) != crate::comparison::COMMENT_COMPARISON_PART
                });
            }
            resolved = resolved
                .checked_add(1)
                .ok_or_else(|| Error::Other("resolved revision count overflowed".to_owned()))?;
        }
        if resolved == 0 {
            return Ok(0);
        }
        crate::comparison::settle_footnote_custody(
            &mut candidate,
            &footnote_custody,
            matches!(resolution, Resolution::Accept),
            publication.as_deref_mut(),
        )?;
        candidate.reconcile_comment_removal(self)?;
        candidate = crate::comparison::reopen_staged(candidate)?;
        if let Some(trace) = publication {
            if let Some(bytes) = candidate.package.get_part(&trace.part_name) {
                trace.validate_source(bytes)?;
                trace.output = Some(bytes.to_vec());
            } else if trace.roots.iter().any(Option::is_some) {
                return Err(Error::Other(
                    "publication part disappeared with retained roots".into(),
                ));
            }
        }
        self.commit_staged_mutation(candidate);
        Ok(resolved)
    }
}

pub(crate) fn resolve_story_xml(
    source: &[u8],
    resolution: Resolution,
    scope: RevisionScope<'_>,
) -> Result<(Vec<u8>, usize)> {
    let (output, count, _) = resolve_story_xml_with_leaf_ranges(source, resolution, scope, &[])?;
    Ok((output, count))
}

/// Transport positively associated whole text leaves through the actual resolver.
/// Indices are local source occurrences. No output correspondence is discovered.
pub(crate) fn resolve_story_xml_with_leaf_ranges(
    source: &[u8],
    resolution: Resolution,
    scope: RevisionScope<'_>,
    leaves: &[Range<usize>],
) -> Result<ResolvedStoryLeaves> {
    let mut tree = XmlTree::parse(source)?;
    tree.qualify_note_revisions()?;
    let mut selected = Vec::with_capacity(leaves.len());
    let mut previous_end = 0;
    for range in leaves {
        if range.start >= range.end || range.end > source.len() || range.start < previous_end {
            return Err(Error::Other(
                "invalid or overlapping comparison leaf ranges".into(),
            ));
        }
        let index = tree
            .elements
            .iter()
            .position(|element| element.start == range.start && element.end == range.end)
            .ok_or_else(|| {
                Error::Other("comparison leaf range is not a whole source element".into())
            })?;
        let element = &tree.elements[index];
        if !element.word
            || !matches!(element.local.as_str(), "t" | "delText")
            || !element.children.is_empty()
            || !element.parent.is_some_and(|parent| {
                tree.elements[parent].word && tree.elements[parent].local == "r"
            })
        {
            return Err(Error::Other(
                "comparison range has no qualified text-leaf owner".into(),
            ));
        }
        selected.push(index);
        previous_end = range.end;
    }
    render_with_source_ranges(&tree, resolution, scope, selected)
}

/// Transport a positively proved complete field publication, using the actual
/// renderer's ordinary-element hooks after descendant and namespace changes.
pub(crate) fn resolve_story_xml_with_publication_roots(
    source: &[u8],
    resolution: Resolution,
    scope: RevisionScope<'_>,
    publication: i32,
    roots: &[Range<usize>],
) -> Result<ResolvedStoryLeaves> {
    let mut tree = XmlTree::parse(source)?;
    tree.qualify_note_revisions()?;
    let selected = publication_root_indices(&tree, resolution, &scope, publication, roots)?;
    render_with_source_ranges(&tree, resolution, scope, selected)
}

fn publication_root_indices(
    tree: &XmlTree<'_>,
    resolution: Resolution,
    scope: &RevisionScope<'_>,
    publication: i32,
    roots: &[Range<usize>],
) -> Result<Vec<usize>> {
    let source = tree.source;
    let mut selected = Vec::with_capacity(roots.len());
    let mut previous = 0usize;
    let mut parent = None;
    for range in roots {
        if range.start >= range.end || range.end > source.len() || range.start < previous {
            return Err(Error::Other(
                "invalid or overlapping publication roots".into(),
            ));
        }
        let index = tree
            .elements
            .iter()
            .position(|element| element.start == range.start && element.end == range.end)
            .ok_or_else(|| {
                Error::Other("publication root is not a complete actual element".into())
            })?;
        let element = &tree.elements[index];
        if !element.word
            || !matches!(
                element.local.as_str(),
                "r" | "sdt"
                    | "customXml"
                    | "smartTag"
                    | "dir"
                    | "bdo"
                    | "proofErr"
                    | "permStart"
                    | "permEnd"
                    | "bookmarkStart"
                    | "bookmarkEnd"
                    | "commentRangeStart"
                    | "commentRangeEnd"
                    | "moveFromRangeStart"
                    | "moveFromRangeEnd"
                    | "moveToRangeStart"
                    | "moveToRangeEnd"
                    | "customXmlInsRangeStart"
                    | "customXmlInsRangeEnd"
                    | "customXmlDelRangeStart"
                    | "customXmlDelRangeEnd"
                    | "customXmlMoveFromRangeStart"
                    | "customXmlMoveFromRangeEnd"
                    | "customXmlMoveToRangeStart"
                    | "customXmlMoveToRangeEnd"
            )
        {
            return Err(Error::Other(
                "publication root has no qualified ordinary run-content owner".into(),
            ));
        }
        let actual_parent = element
            .parent
            .ok_or_else(|| Error::Other("publication root lacks a revision owner".into()))?;
        if parent.is_some_and(|parent| parent != actual_parent) {
            return Err(Error::Other(
                "publication roots cross actual revision owners".into(),
            ));
        }
        let metadata = tree.elements[actual_parent]
            .revision
            .as_ref()
            .ok_or_else(|| Error::Other("publication parent is not a modeled revision".into()))?;
        if metadata.id != publication
            || !scope.matches(metadata)
            || !keeps_content(resolution, metadata.kind)
            || !matches!(
                metadata.kind,
                RevisionKind::Insertion | RevisionKind::Deletion
            )
        {
            return Err(Error::Other(
                "publication root belongs to another or discarded revision".into(),
            ));
        }
        parent = Some(actual_parent);
        selected.push(index);
        previous = range.end;
    }
    let parent =
        parent.ok_or_else(|| Error::Other("field publication has no actual roots".into()))?;
    if selected != tree.elements[parent].children {
        return Err(Error::Other(
            "publication root coverage omits or duplicates an actual member".into(),
        ));
    }
    let trivia = |bytes: &[u8]| {
        bytes
            .iter()
            .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    };
    let mut cursor = tree.elements[parent].open_end;
    for index in &selected {
        let element = &tree.elements[*index];
        if !trivia(&source[cursor..element.start]) {
            return Err(Error::Other(
                "publication root gaps contain independent source".into(),
            ));
        }
        cursor = element.end;
    }
    if !trivia(&source[cursor..tree.elements[parent].close_start]) {
        return Err(Error::Other(
            "publication has independent trailing source".into(),
        ));
    }
    fn marker(tree: &XmlTree<'_>, index: usize) -> Result<String> {
        let element = &tree.elements[index];
        let scope = tree
            .source_scope(index)
            .into_iter()
            .collect::<HashMap<_, _>>();
        let mut reader = Reader::from_reader(&tree.source[element.start..element.end]);
        let (Event::Start(root) | Event::Empty(root)) = reader.read_event().map_err(xml_error)?
        else {
            return Err(Error::Other("field marker is missing".into()));
        };
        let mut kind = None;
        for attribute in root.attributes() {
            let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
            let key = std::str::from_utf8(attribute.key.as_ref())
                .map_err(|error| Error::Other(error.to_string()))?;
            if let Some((prefix, local)) = key.split_once(':')
                && scope.get(prefix).is_some_and(|uri| uri == WORD_NS)
                && local == "fldCharType"
            {
                if kind.is_some() {
                    return Err(Error::Other("duplicate qualified field marker kind".into()));
                }
                kind = Some(
                    attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, root.decoder())
                        .map_err(xml_error)?
                        .into_owned(),
                );
            }
        }
        kind.ok_or_else(|| Error::Other("field marker lacks a qualified kind".into()))
    }
    fn edge(tree: &XmlTree<'_>, run: usize, marker: usize, begin: bool) -> bool {
        let element = &tree.elements[run];
        let properties = element
            .children
            .iter()
            .filter(|index| tree.elements[**index].word && tree.elements[**index].local == "rPr")
            .copied()
            .collect::<Vec<_>>();
        if properties.len() > 1
            || properties
                .first()
                .is_some_and(|property| element.children.first() != Some(property))
        {
            return false;
        }
        let material = &element.children[usize::from(!properties.is_empty())..];
        if (begin && material.first() != Some(&marker))
            || (!begin && material.last() != Some(&marker))
        {
            return false;
        }
        let boundary = if begin {
            properties
                .first()
                .map_or(element.open_end, |index| tree.elements[*index].end)
                ..tree.elements[marker].start
        } else {
            tree.elements[marker].end..element.close_start
        };
        tree.source[boundary]
            .iter()
            .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    }
    let mut depth = 0usize;
    let mut separate = false;
    for (slot, run) in selected.iter().copied().enumerate() {
        if tree.elements[run].local != "r" {
            continue;
        }
        for child in tree.elements[run].children.iter().copied() {
            if !tree.elements[child].word || tree.elements[child].local != "fldChar" {
                continue;
            }
            match marker(tree, child)?.as_str() {
                "begin" => {
                    if depth == 0 && (slot != 0 || !edge(tree, run, child, true)) {
                        return Err(Error::Other(
                            "publication begins with a shared or partial field run".into(),
                        ));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| Error::Other("publication field depth overflow".into()))?;
                }
                "separate" if depth == 1 => {
                    if separate {
                        return Err(Error::Other(
                            "publication has an ambiguous separate run".into(),
                        ));
                    }
                    separate = true;
                }
                "separate" if depth > 1 => {}
                "end" if depth > 0 => {
                    if depth == 1 && (slot + 1 != selected.len() || !edge(tree, run, child, false))
                    {
                        return Err(Error::Other(
                            "publication ends with a shared or partial field run".into(),
                        ));
                    }
                    depth -= 1;
                }
                _ => {
                    return Err(Error::Other(
                        "publication field markers are unbalanced".into(),
                    ));
                }
            }
        }
    }
    if depth != 0 || !separate {
        return Err(Error::Other(
            "publication lacks one complete actual complex field".into(),
        ));
    }
    Ok(selected)
}

fn render_with_source_ranges(
    tree: &XmlTree<'_>,
    resolution: Resolution,
    scope: RevisionScope<'_>,
    selected: Vec<usize>,
) -> Result<ResolvedStoryLeaves> {
    let mut state = RenderState {
        resolution,
        scope,
        resolved: HashSet::new(),
        discarded: false,
        leaf_ranges: vec![None; selected.len()],
        leaf_elements: selected,
        leaf_updates: Vec::new(),
        retiring_simple_fields: HashSet::new(),
    };
    state.retiring_simple_fields = tree.carried_simple_field_retirements(&state)?;
    let mut output = Vec::with_capacity(tree.source.len());
    output.extend_from_slice(&tree.source[..tree.elements[tree.root].start]);
    let root = tree.render(tree.root, &mut state, false)?;
    state.shift_leaves(0, output.len());
    output.extend_from_slice(&root);
    output.extend_from_slice(&tree.source[tree.elements[tree.root].end..]);
    Ok((output, state.resolved.len(), state.leaf_ranges))
}

pub(crate) fn modeled_revision_count(source: &[u8]) -> Result<usize> {
    let mut tree = XmlTree::parse(source)?;
    tree.qualify_note_revisions()?;
    Ok(tree
        .elements
        .iter()
        .filter(|element| element.revision.is_some())
        .count())
}

/// The modeled revision elements of one part in document order: the byte
/// range of each element and its `w:id`. These are the elements that
/// [`modeled_revision_count`] counts.
pub(crate) fn modeled_revision_spans(source: &[u8]) -> Result<Vec<(Range<usize>, i32)>> {
    let mut tree = XmlTree::parse(source)?;
    tree.qualify_note_revisions()?;
    Ok(tree
        .elements
        .iter()
        .filter_map(|element| Some((element.start..element.end, element.revision.as_ref()?.id)))
        .collect())
}

/// List the modeled revisions of one story part in document order.
///
/// These are the elements that resolving the part counts. Each one belongs to
/// the innermost span in `owners` that holds its start tag.
pub(crate) fn story_part_revisions(
    part_name: &str,
    source: &[u8],
    owners: &[(Range<usize>, StoryId)],
) -> Result<Vec<StoryRevision>> {
    let mut tree = XmlTree::parse(source)?;
    tree.qualify_note_revisions()?;
    let facts = tree.snapshot_facts();
    tree.elements
        .iter()
        .enumerate()
        .filter_map(|(index, element)| Some((index, element.start, element.revision.as_ref()?)))
        .map(|(index, start, metadata)| {
            let (_, story) = owners
                .iter()
                .filter(|(span, _)| span.contains(&start))
                .max_by_key(|(span, _)| span.start)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "tracked revision at byte {start} of {part_name} has no story owner"
                    ))
                })?;
            Ok(StoryRevision {
                story: story.clone(),
                id: metadata.id,
                author: metadata.author.clone(),
                timestamp: metadata.timestamp.clone(),
                kind: metadata.kind,
                text: tree.snapshot_text(index, &facts)?,
            })
        })
        .collect()
}

fn date_scope(start: &str, end: &str) -> Result<RevisionScope<'static>> {
    let start = Instant::parse(start)?;
    let end = Instant::parse(end)?;
    if start > end {
        return Err(Error::Other(
            "revision date range starts after it ends".to_owned(),
        ));
    }
    Ok(RevisionScope::DateRange { start, end })
}

pub(crate) fn validate_revision_timestamp(value: &str) -> Result<()> {
    Instant::parse(value).map(|_| ())
}

impl RevisionScope<'_> {
    pub(crate) fn matches(&self, metadata: &RevisionMetadata) -> bool {
        match self {
            Self::All => true,
            Self::Author(author) => metadata.author == *author,
            Self::Id(id) => metadata.id == *id,
            Self::Ids(ids) => ids.contains(&metadata.id),
            Self::DateRange { start, end } => metadata
                .timestamp
                .as_deref()
                .and_then(|value| Instant::parse(value).ok())
                .is_some_and(|instant| *start <= instant && instant <= *end),
        }
    }
}

/// Track actual paragraph-content owners without moving a simple field or
/// hyperlink into the narrower run-revision grammar. The adaptor supplies only
/// the caller's paragraph root and canonical writer binding, not a new resolver root.
pub(crate) fn tracked_paragraph_children(
    source: &str,
    deleted: bool,
    inside_hyperlink: bool,
    wrap: &mut impl FnMut(&str) -> Result<String>,
) -> Result<String> {
    struct Frame {
        owner: usize,
        next: usize,
        cursor: usize,
        groups: Vec<(String, bool)>,
        hyperlink: bool,
        output_start: usize,
    }
    fn flush(
        groups: &mut Vec<(String, bool)>,
        output: &mut String,
        wrap: &mut impl FnMut(&str) -> Result<String>,
    ) -> Result<()> {
        if groups.len() > 1
            && let Some((markers, false)) = groups.pop_if(|(_, content)| !*content)
            && let Some((group, _)) = groups.last_mut()
        {
            group.push_str(&markers);
        }
        for (group, _) in groups.drain(..) {
            output.push_str(&wrap(&group)?);
        }
        Ok(())
    }
    let opening = format!("<w:document xmlns:w=\"{WORD_NS}\"><w:body>");
    let input = format!("{opening}{source}</w:body></w:document>");
    let tree = XmlTree::parse(input.as_bytes())?;
    let body = tree.direct_word_child(tree.root, "body")?;
    let [paragraph] = tree.elements[body].children.as_slice() else {
        return Err(Error::Other(
            "whole-paragraph publication needs one actual paragraph".into(),
        ));
    };
    if !tree.elements[*paragraph].word || tree.elements[*paragraph].local != "p" {
        return Err(Error::Other(
            "whole-paragraph publication root is not Word p".into(),
        ));
    }
    let publication_frames = tree.paragraph_publication_frames(body);
    // Source-only paragraph grammar. This does not promote preserved lower
    // hyperlink/control payloads into typed Field or borrowed revision owners.
    let mut inline_source = vec![false; tree.elements.len()];
    let mut hyperlink_source = vec![false; tree.elements.len()];
    for (index, element) in tree.elements.iter().enumerate() {
        inline_source[index] = index == *paragraph
            || element.word
                && element.parent.is_some_and(|parent| {
                    inline_source[parent]
                        && match tree.elements[parent].local.as_str() {
                            "p" | "fldSimple" | "hyperlink" | "sdtContent" => matches!(
                                element.local.as_str(),
                                "r" | "fldSimple" | "hyperlink" | "sdt"
                            ),
                            "sdt" => element.local == "sdtContent",
                            _ => false,
                        }
                });
        hyperlink_source[index] = inline_source[index]
            && (element.local == "hyperlink"
                || element
                    .parent
                    .is_some_and(|parent| hyperlink_source[parent]));
    }
    let mut frames = vec![Frame {
        owner: *paragraph,
        next: 0,
        cursor: tree.elements[*paragraph].open_end,
        groups: Vec::new(),
        hyperlink: inside_hyperlink,
        output_start: 0,
    }];
    let mut output = String::new();
    while let Some(frame) = frames.last_mut() {
        let owner = &tree.elements[frame.owner];
        let Some(child_index) = owner.children.get(frame.next).copied() else {
            if frame.cursor < owner.close_start {
                frame
                    .groups
                    .push((input[frame.cursor..owner.close_start].to_owned(), false));
            }
            flush(&mut frame.groups, &mut output, wrap)?;
            if frame.owner != *paragraph {
                if owner.local == "hyperlink" && output.len() == frame.output_start {
                    return Err(Error::Other(
                        "comparison cannot track a whole paragraph's empty hyperlink".into(),
                    ));
                }
                output.push_str(&input[owner.close_start..owner.end]);
            }
            frames.pop();
            continue;
        };
        frame.next += 1;
        let child = &tree.elements[child_index];
        if frame.cursor < child.start {
            frame
                .groups
                .push((input[frame.cursor..child.start].to_owned(), false));
        }
        frame.cursor = child.end;
        if publication_frames[child_index] {
            if child.local == "hyperlink" && frame.hyperlink {
                return Err(Error::Other(
                    "comparison cannot track nested hyperlinks in a whole paragraph".into(),
                ));
            }
            flush(&mut frame.groups, &mut output, wrap)?;
            if child.empty {
                if child.local == "hyperlink" {
                    return Err(Error::Other(
                        "comparison cannot track a whole paragraph's empty hyperlink".into(),
                    ));
                }
                output.push_str(&input[child.start..child.end]);
                continue;
            }
            output.push_str(&input[child.start..child.open_end]);
            let mut next = 0;
            let mut cursor = child.open_end;
            if child.local == "fldSimple" {
                let data = child
                    .children
                    .iter()
                    .copied()
                    .filter(|index| {
                        tree.elements[*index].word && tree.elements[*index].local == "fldData"
                    })
                    .collect::<Vec<_>>();
                if data.len() > 1
                    || data
                        .first()
                        .is_some_and(|index| child.children.first() != Some(index))
                {
                    return Err(Error::Other(
                        "whole-field source has ambiguous or late fldData".into(),
                    ));
                }
                if let Some(data) = data.first() {
                    cursor = tree.elements[*data].end;
                    output.push_str(&input[child.open_end..cursor]);
                    next = 1;
                }
            }
            let hyperlink = frame.hyperlink || child.local == "hyperlink";
            let output_start = output.len();
            frames.push(Frame {
                owner: child_index,
                next,
                cursor,
                groups: Vec::new(),
                hyperlink,
                output_start,
            });
            continue;
        }
        let end = tree
            .elements
            .partition_point(|element| element.start < child.end);
        let mut renames = Vec::new();
        for index in child_index..end {
            let element = &tree.elements[index];
            if inline_source[index]
                && element.local == "hyperlink"
                && (frame.hyperlink
                    || element
                        .parent
                        .is_some_and(|parent| hyperlink_source[parent]))
            {
                return Err(Error::Other(
                    "comparison cannot track nested hyperlinks in a whole paragraph".into(),
                ));
            }
            if deleted
                && element.word
                && matches!(element.local.as_str(), "t" | "instrText")
                && element.parent.is_some_and(|parent| {
                    inline_source[parent]
                        && tree.elements[parent].modeled
                        && tree.elements[parent].word
                        && tree.elements[parent].local == "r"
                })
            {
                renames.push(index);
            }
        }
        let mut raw = input[child.start..child.end].to_owned();
        for index in renames.into_iter().rev() {
            let leaf = &tree.elements[index];
            let local = if leaf.local == "t" {
                "delText"
            } else {
                "delInstrText"
            };
            let mut renamed =
                rename_element(&tree.source[leaf.start..leaf.open_end], &leaf.name, local);
            if !leaf.empty {
                renamed.extend_from_slice(&tree.source[leaf.open_end..leaf.close_start]);
                renamed.extend_from_slice(&rename_element(
                    &tree.source[leaf.close_start..leaf.end],
                    &leaf.name,
                    local,
                ));
            }
            raw.replace_range(
                leaf.start - child.start..leaf.end - child.start,
                std::str::from_utf8(&renamed).map_err(|error| Error::Other(error.to_string()))?,
            );
        }
        let mut destination = tree.source_scope(frame.owner);
        destination.retain(|(prefix, _)| prefix != "w");
        destination.push(("w".to_owned(), WORD_NS.to_owned()));
        let declarations = tree.relocation_namespaces(child_index, &destination);
        // Descendant edits leave this atomic owner's opening unchanged. Close
        // its namespace scope on that tag before appending the edited remainder.
        let opening_len = child.open_end - child.start;
        let mut closed = inject_namespace_declarations(
            &raw.as_bytes()[..opening_len],
            &child.namespace_declarations,
            &declarations,
        );
        closed.extend_from_slice(&raw.as_bytes()[opening_len..]);
        raw = String::from_utf8(closed).map_err(|error| Error::Other(error.to_string()))?;
        let marker = child.word
            && matches!(
                child.local.as_str(),
                "bookmarkStart"
                    | "bookmarkEnd"
                    | "commentRangeStart"
                    | "commentRangeEnd"
                    | "proofErr"
            );
        if let Some((group, content)) = frame.groups.last_mut()
            && !*content
        {
            group.push_str(&raw);
            *content = !marker;
        } else {
            frame.groups.push((raw, !marker));
        }
    }
    Ok(output)
}

#[derive(Default)]
struct CarriedField {
    markers: [Option<usize>; 3],
    edge_runs: [usize; 2],
    parent: Option<usize>,
    domain: usize,
    domain_end: usize,
    simple_fields: Vec<usize>,
    valid: bool,
}

type CarriedFieldSources = (Vec<CarriedField>, Vec<(usize, Option<usize>)>);

/// The coarse source events used to grow a comparison's paragraph range.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum FieldEvent {
    Begin,
    End,
    Owner,
}

pub(crate) fn paragraph_field_events(paragraph: &CT_P) -> Result<Vec<FieldEvent>> {
    let mut source = format!("<w:document xmlns:w=\"{WORD_NS}\"><w:body>").into_bytes();
    paragraph.to_xml(&mut quick_xml::Writer::new(&mut source))?;
    source.extend_from_slice(b"</w:body></w:document>");
    let tree = XmlTree::parse(&source)?;
    let body = tree.direct_word_child(tree.root, "body")?;
    let paragraph = tree.direct_word_child(body, "p")?;
    Ok(tree.paragraph_field_events(paragraph))
}

/// The actual paragraph source must establish carriage before its publisher may
/// leave simple/hyperlink shells outside the selected run wrappers.
pub(crate) fn validate_carried_paragraph_fields(paragraphs: &[&CT_P]) -> Result<()> {
    let mut source = format!("<w:document xmlns:w=\"{WORD_NS}\"><w:body>").into_bytes();
    for paragraph in paragraphs {
        paragraph.to_xml(&mut quick_xml::Writer::new(&mut source))?;
    }
    source.extend_from_slice(b"</w:body></w:document>");
    let tree = XmlTree::parse(&source)?;
    let body = tree.direct_word_child(tree.root, "body")?;
    let publication_frames = tree.paragraph_publication_frames(body);
    let (fields, owners) = tree.carried_field_sources()?;
    for (index, owner) in owners {
        if publication_frames[index]
            && owner.is_none_or(|index| {
                !fields[index].valid || fields[index].markers.iter().any(Option::is_none)
            })
        {
            return Err(Error::Other(
                "whole-paragraph field owner lacks a qualified complete source carrier".into(),
            ));
        }
    }
    Ok(())
}

enum SnapshotFieldTransition {
    Keep,
    Pop(Option<usize>),
    Push(SnapshotField),
}

fn snapshot_field_transition_from(
    current: Option<SnapshotField>,
    head: Option<usize>,
    node: SnapshotNode,
    marker: &str,
) -> SnapshotFieldTransition {
    match marker {
        "begin" => SnapshotFieldTransition::Push(SnapshotField {
            previous: head,
            domain: node.domain,
            end: node.domain_end,
            instructions: current.map_or(0, |field| field.instructions) + 1,
            result: false,
        }),
        "separate" => match current {
            Some(field) if field.domain == node.domain && !field.result => {
                SnapshotFieldTransition::Push(SnapshotField {
                    result: true,
                    instructions: field.instructions - 1,
                    ..field
                })
            }
            _ => SnapshotFieldTransition::Keep,
        },
        "end" => match current {
            Some(field) if field.domain == node.domain => {
                SnapshotFieldTransition::Pop(field.previous)
            }
            _ => SnapshotFieldTransition::Keep,
        },
        _ => SnapshotFieldTransition::Keep,
    }
}

impl<'a> XmlTree<'a> {
    // Physical roots come from the same admitted story inventory as discovery.
    // Fold ordinary modeled ancestry once. Only a positively admitted child
    // story may restart that ancestry across a drawing's opaque envelope.
    pub(crate) fn admitted_story_elements(&self, root_kind: crate::StoryKind) -> Result<Vec<bool>> {
        let roots = crate::document::scan_story_owners_with_special(self.source, root_kind, true)?
            .into_iter()
            .map(|owner| (owner.full.start, owner.full.end))
            .collect::<HashSet<_>>();
        let mut admitted = Vec::with_capacity(self.elements.len());
        for element in &self.elements {
            let actual_root = roots.contains(&(element.start, element.end));
            let parent_admitted = element.parent.is_some_and(|parent| admitted[parent]);
            admitted.push(element.word && element.modeled && (actual_root || parent_admitted));
        }
        Ok(admitted)
    }

    // This removes metadata only. It neither promotes invalid/special note
    // owners to normal nor changes the inherited generic textbox parser policy.
    pub(crate) fn qualify_note_revisions(&mut self) -> Result<()> {
        let root = &self.elements[self.root];
        if !root.word || root.local != "footnotes" {
            return Ok(());
        }
        let admitted = self.admitted_story_elements(crate::StoryKind::Footnote)?;
        for (element, admitted) in self.elements.iter_mut().zip(&admitted) {
            if !admitted {
                element.revision = None;
            }
        }
        self.note_admission = Some(admitted);
        Ok(())
    }

    // The revision parser models structural containers, while scalar run
    // content remains raw. Only recognized scalar leaves of an admitted run
    // share its semantic authority. Opaque ancestors do not transmit it.
    fn semantic_element(&self, index: usize) -> bool {
        self.note_admission.as_ref().is_none_or(|admitted| {
            let element = &self.elements[index];
            admitted[index]
                || element.word
                    && element.children.is_empty()
                    && matches!(
                        element.local.as_str(),
                        "t" | "delText" | "instrText" | "delInstrText"
                    )
                    && element.parent.is_some_and(|parent| {
                        admitted[parent]
                            && self.elements[parent].word
                            && self.elements[parent].local == "r"
                    })
        })
    }

    /// Build source context once for the whole snapshot batch. Persistent
    /// scalar field heads share prefixes without copying stacks or source text.
    fn snapshot_facts(&self) -> SnapshotFacts {
        let (stories, ends) = self.paragraph_source_domains();
        let mut facts = SnapshotFacts {
            nodes: Vec::with_capacity(self.elements.len()),
            fields: Vec::new(),
            #[cfg(test)]
            preparation_visits: 0,
            #[cfg(test)]
            owner_visits: std::cell::Cell::new(0),
        };
        for (index, element) in self.elements.iter().enumerate() {
            #[cfg(test)]
            {
                facts.preparation_visits += 1;
            }
            let parent = element.parent.map(|parent| facts.nodes[parent]);
            let story = stories[index].unwrap_or(index);
            let mut node = SnapshotNode {
                story,
                domain: parent.map_or(story, |parent| parent.domain),
                domain_end: parent.map_or(element.end, |parent| parent.domain_end),
                canonical_w: parent.is_some_and(|parent| parent.canonical_w),
                control_parent: element.parent.and_then(|parent| {
                    if matches!(self.elements[parent].local.as_str(), "sdt" | "sdtContent") {
                        facts.nodes[parent].control_parent
                    } else {
                        Some(parent)
                    }
                }),
                admitted: false,
                wrappers: true,
                hidden: parent.is_some_and(|parent| parent.hidden),
                row_removed: element.word
                    && element.local == "tr"
                    && self.snapshot_row_removed(index),
                seed: None,
            };
            for (prefix, uri) in &element.namespace_declarations {
                if prefix == "w" {
                    node.canonical_w = uri == WORD_NS;
                }
            }
            if element.word && element.local == "p" {
                node.domain = story;
                node.domain_end = ends.get(&story).copied().unwrap_or(element.end);
            } else if element.word
                && matches!(
                    element.local.as_str(),
                    "fldSimple" | "hyperlink" | "sdtContent"
                )
            {
                node.domain = index;
                node.domain_end = element.end;
            }
            node.hidden |= node.row_removed
                || element.revision.as_ref().is_some_and(|metadata| {
                    matches!(
                        metadata.kind,
                        RevisionKind::Deletion | RevisionKind::MoveFrom
                    )
                });
            // Story roots may independently model textbox contents. An outer
            // owner's traversal still cannot cross the opaque drawing path.
            if element.word
                && element.modeled
                && matches!(
                    element.local.as_str(),
                    "body" | "hdr" | "ftr" | "comment" | "footnote" | "endnote" | "txbxContent"
                )
            {
                node.admitted = true;
            } else if let Some(parent) = element.parent
                && facts.nodes[parent].admitted
                && let Some(wrappers) = self.snapshot_child(parent, index, &facts, node.canonical_w)
            {
                node.admitted = true;
                node.wrappers = wrappers;
            }
            facts.nodes.push(node);
        }
        let mut heads = HashMap::<usize, Option<usize>>::new();
        let mut undo = Vec::<(usize, Option<usize>)>::new();
        let mut branches = Vec::<(usize, usize)>::new();
        for (index, element) in self.elements.iter().enumerate() {
            #[cfg(test)]
            {
                facts.preparation_visits += 1;
            }
            while branches
                .last()
                .is_some_and(|(end, _)| *end <= element.start)
            {
                let (_, checkpoint) = branches.pop().expect("branch exists");
                while undo.len() > checkpoint {
                    let (story, previous) = undo.pop().expect("undo exists");
                    heads.insert(story, previous);
                }
            }
            let node = facts.nodes[index];
            if !node.admitted {
                continue;
            }
            let mut head = heads.get(&node.story).copied().flatten();
            while let Some(active) = head {
                if facts.fields[active].end > element.start {
                    break;
                }
                head = facts.fields[active].previous;
            }
            let previous = heads.get(&node.story).copied().flatten();
            if head != previous {
                if !branches.is_empty() {
                    undo.push((node.story, previous));
                }
                heads.insert(node.story, head);
            }
            facts.nodes[index].seed = head;
            if element.revision.as_ref().is_some_and(|metadata| {
                matches!(
                    metadata.kind,
                    RevisionKind::Deletion | RevisionKind::MoveFrom
                )
            }) {
                branches.push((element.end, undo.len()));
            }
            if let Some(marker) = element.snapshot_field_marker.as_deref() {
                let next = snapshot_field_transition_from(
                    head.map(|index| facts.fields[index]),
                    head,
                    node,
                    marker,
                );
                let next = match next {
                    SnapshotFieldTransition::Keep => head,
                    SnapshotFieldTransition::Pop(previous) => previous,
                    SnapshotFieldTransition::Push(field) => {
                        let index = facts.fields.len();
                        facts.fields.push(field);
                        Some(index)
                    }
                };
                if next != head {
                    if !branches.is_empty() {
                        undo.push((node.story, head));
                    }
                    heads.insert(node.story, next);
                }
            }
        }
        facts
    }

    /// The same concrete child grammar drives context and owner projection.
    fn snapshot_child(
        &self,
        parent: usize,
        child: usize,
        facts: &SnapshotFacts,
        canonical_w: bool,
    ) -> Option<bool> {
        let element = &self.elements[parent];
        let next = &self.elements[child];
        if !next.word {
            return None;
        }
        let local = next.local.as_str();
        let wrappers = facts.nodes[parent].wrappers;
        let role = if element.local == "sdtContent" {
            facts.nodes[parent]
                .control_parent
                .map(|index| self.elements[index].local.as_str())
                .unwrap_or("")
        } else {
            element.local.as_str()
        };
        if element.local == "sdtContent"
            && matches!(
                role,
                "p" | "hyperlink"
                    | "ins"
                    | "del"
                    | "moveFrom"
                    | "moveTo"
                    | "smartTag"
                    | "customXml"
                    | "fldSimple"
            )
        {
            return matches!(local, "r" | "sdt" | "ins" | "del" | "moveFrom" | "moveTo")
                .then_some(false);
        }
        let admitted = match role {
            "tbl" => matches!(local, "tr" | "sdt"),
            "tr" => matches!(local, "tc" | "sdt"),
            "body" | "hdr" | "ftr" | "comment" | "footnote" | "endnote" | "txbxContent" | "tc" => {
                matches!(local, "p" | "tbl" | "sdt")
            }
            "sdt" => local == "sdtContent",
            "p" | "smartTag" | "customXml" | "fldSimple" => {
                matches!(
                    local,
                    "r" | "hyperlink" | "sdt" | "ins" | "del" | "moveFrom" | "moveTo"
                ) || wrappers
                    && canonical_w
                    && !next.empty
                    && matches!(local, "smartTag" | "customXml" | "fldSimple")
            }
            "hyperlink" | "ins" | "del" | "moveFrom" | "moveTo" => matches!(
                local,
                "r" | "hyperlink" | "sdt" | "ins" | "del" | "moveFrom" | "moveTo"
            ),
            "r" => matches!(local, "t" | "delText" | "tab" | "br" | "cr" | "fldChar"),
            _ => false,
        };
        admitted.then_some(if matches!(local, "p" | "tc" | "tbl") {
            true
        } else {
            wrappers
                && !matches!(
                    local,
                    "sdt" | "sdtContent" | "hyperlink" | "ins" | "del" | "moveFrom" | "moveTo"
                )
        })
    }

    fn snapshot_text(&self, revision: usize, facts: &SnapshotFacts) -> Result<String> {
        let kind = self.elements[revision]
            .revision
            .as_ref()
            .expect("revision index")
            .kind;
        let current = matches!(
            kind,
            RevisionKind::RunPropertyChange
                | RevisionKind::ParagraphPropertyChange
                | RevisionKind::TablePropertyChange
        );
        if kind == RevisionKind::SectionPropertyChange {
            return Ok(String::new());
        }
        let owner = if current {
            let expected = match kind {
                RevisionKind::RunPropertyChange => ("rPr", "r"),
                RevisionKind::ParagraphPropertyChange => ("pPr", "p"),
                RevisionKind::TablePropertyChange => ("tblPr", "tbl"),
                _ => unreachable!("property kind"),
            };
            let Some(properties) = self.elements[revision].parent else {
                return Ok(String::new());
            };
            let Some(owner) = self.elements[properties].parent else {
                return Ok(String::new());
            };
            if !self.elements[properties].word
                || self.elements[properties].local != expected.0
                || !self.elements[owner].word
                || self.elements[owner].local != expected.1
            {
                return Ok(String::new());
            }
            owner
        } else {
            revision
        };
        if !facts.nodes[owner].admitted || current && facts.nodes[owner].hidden {
            return Ok(String::new());
        }
        let table = self.elements[owner].local == "tbl";
        let mut output = String::new();
        let mut pending = vec![(owner, false, 0usize)];
        let mut heads = HashMap::<usize, Option<usize>>::new();
        let mut local_fields = Vec::<SnapshotField>::new();
        let base = facts.fields.len();
        while let Some((index, exiting, table_depth)) = pending.pop() {
            #[cfg(test)]
            {
                facts.owner_visits.set(facts.owner_visits.get() + 1);
            }
            let element = &self.elements[index];
            let node = facts.nodes[index];
            if exiting {
                if table && element.local == "p" {
                    output.push('\t');
                } else if table && element.local == "tr" && table_depth == 1 {
                    output.push('\n');
                }
                continue;
            }
            if current
                && (node.row_removed
                    || index != owner
                        && element.revision.as_ref().is_some_and(|metadata| {
                            matches!(
                                metadata.kind,
                                RevisionKind::Deletion | RevisionKind::MoveFrom
                            )
                        }))
            {
                continue;
            }
            let head = heads.entry(node.story).or_insert(node.seed);
            while let Some(active) = *head {
                let field = if active < base {
                    facts.fields[active]
                } else {
                    local_fields[active - base]
                };
                if field.end > element.start {
                    break;
                }
                *head = field.previous;
            }
            if let Some(marker) = element.snapshot_field_marker.as_deref() {
                let current = head.map(|active| {
                    if active < base {
                        facts.fields[active]
                    } else {
                        local_fields[active - base]
                    }
                });
                let transition = snapshot_field_transition_from(current, *head, node, marker);
                match transition {
                    SnapshotFieldTransition::Keep => {}
                    SnapshotFieldTransition::Pop(previous) => *head = previous,
                    SnapshotFieldTransition::Push(field) => {
                        *head = Some(base + local_fields.len());
                        local_fields.push(field);
                    }
                }
                continue;
            }
            if matches!(
                element.local.as_str(),
                "t" | "delText" | "tab" | "br" | "cr"
            ) {
                let visible = head.is_none_or(|active| {
                    (if active < base {
                        facts.fields[active]
                    } else {
                        local_fields[active - base]
                    })
                    .instructions
                        == 0
                });
                if visible {
                    match element.local.as_str() {
                        "tab" => output.push('\t'),
                        "br" | "cr" => output.push('\n'),
                        _ => output.push_str(&element.snapshot_scalar),
                    }
                }
                continue;
            }
            let depth = table_depth + usize::from(element.local == "tbl");
            pending.push((index, true, depth));
            for child in element.children.iter().rev().copied() {
                if self
                    .snapshot_child(index, child, facts, facts.nodes[child].canonical_w)
                    .is_some()
                {
                    pending.push((child, false, depth));
                }
            }
        }
        Ok(output)
    }

    fn snapshot_row_removed(&self, row: usize) -> bool {
        self.elements[row]
            .children
            .iter()
            .copied()
            .filter(|child| self.elements[*child].word && self.elements[*child].local == "trPr")
            .any(|properties| {
                self.elements[properties].children.iter().any(|child| {
                    self.elements[*child]
                        .revision
                        .as_ref()
                        .is_some_and(|metadata| {
                            matches!(
                                metadata.kind,
                                RevisionKind::Deletion | RevisionKind::MoveFrom
                            )
                        })
                })
            })
    }

    /// Exactly the shells left outside run revisions by paragraph publication.
    /// Every other child, including an SDT or opaque drawing, stays atomic.
    fn paragraph_publication_frames(&self, body: usize) -> Vec<bool> {
        let mut frames = vec![false; self.elements.len()];
        for (index, element) in self.elements.iter().enumerate() {
            frames[index] = element.modeled
                && element.word
                && match element.local.as_str() {
                    "p" => element.parent == Some(body),
                    "fldSimple" | "hyperlink" => {
                        element.parent.is_some_and(|parent| frames[parent])
                    }
                    _ => false,
                };
        }
        frames
    }

    /// Consecutive paragraphs share a structural domain. A nested paragraph,
    /// including an independently modeled textbox, starts its own domain.
    fn paragraph_source_domains(&self) -> (Vec<Option<usize>>, HashMap<usize, usize>) {
        let count = self.elements.len();
        let mut stories = vec![None; count];
        let mut next_domain = count;
        let mut domain_ends = HashMap::new();
        for parent in &self.elements {
            let mut domain = None;
            for child in &parent.children {
                let element = &self.elements[*child];
                if element.modeled && element.word && element.local == "p" {
                    let current = *domain.get_or_insert_with(|| {
                        next_domain += 1;
                        next_domain
                    });
                    stories[*child] = Some(current);
                    domain_ends.insert(current, element.end);
                } else {
                    domain = None;
                }
            }
        }
        for (index, element) in self.elements.iter().enumerate() {
            if stories[index].is_none() {
                stories[index] = element.parent.and_then(|parent| stories[parent]);
            }
        }
        (stories, domain_ends)
    }

    fn field_marker_run(&self, index: usize) -> Option<usize> {
        let element = &self.elements[index];
        if !element.word || element.local != "fldChar" {
            return None;
        }
        element.parent.filter(|parent| {
            let parent = &self.elements[*parent];
            parent.modeled && parent.word && parent.local == "r"
        })
    }

    fn paragraph_field_events(&self, paragraph: usize) -> Vec<FieldEvent> {
        let (stories, _) = self.paragraph_source_domains();
        let mut events = Vec::new();
        for (index, element) in self.elements.iter().enumerate() {
            if stories[index] != stories[paragraph] {
                continue;
            }
            if element.parent == Some(paragraph)
                && element.modeled
                && element.word
                && matches!(element.local.as_str(), "fldSimple" | "hyperlink")
            {
                events.push(FieldEvent::Owner);
            } else if self.field_marker_run(index).is_some() {
                match element.field_marker.as_deref() {
                    Some("begin") => events.push(FieldEvent::Begin),
                    Some("end") => events.push(FieldEvent::End),
                    _ => {}
                }
            }
        }
        events
    }

    /// Inventory actual qualified marker occurrences once. Inline owners can
    /// establish marker domains inside the same structural paragraph domain.
    fn carried_field_sources(&self) -> Result<CarriedFieldSources> {
        let count = self.elements.len();
        let (stories, mut domain_ends) = self.paragraph_source_domains();
        let mut domains = vec![None; count];
        let mut fields = Vec::<CarriedField>::new();
        let mut story_fields = HashMap::<usize, Vec<usize>>::new();
        let mut owners = Vec::new();
        for (index, element) in self.elements.iter().enumerate() {
            domains[index] = if element.modeled && element.word && element.local == "p" {
                stories[index]
            } else {
                element.parent.and_then(|parent| domains[parent])
            };
            let Some(story) = stories[index] else {
                continue;
            };
            let active = story_fields.entry(story).or_default();
            while active
                .last()
                .is_some_and(|field| fields[*field].domain_end <= element.start)
            {
                fields[active.pop().expect("active field")].valid = false;
            }
            if element.modeled
                && element.word
                && matches!(element.local.as_str(), "fldSimple" | "hyperlink")
            {
                let owner = active.last().copied();
                owners.push((index, owner));
                if element.local == "fldSimple"
                    && let Some(owner) = owner
                {
                    fields[owner].simple_fields.push(index);
                }
                domains[index] = Some(index);
                domain_ends.insert(index, element.end);
            } else if element.modeled
                && element.word
                && element.local == "sdtContent"
                && inline_field_control_owner(&self.elements, index)
            {
                domains[index] = Some(index);
                domain_ends.insert(index, element.end);
            }
            let Some(run) = self.field_marker_run(index) else {
                continue;
            };
            let Some(domain) = domains[run] else {
                continue;
            };
            match element.field_marker.as_deref() {
                Some("begin") => {
                    let field = fields.len();
                    fields.push(CarriedField {
                        markers: [Some(index), None, None],
                        edge_runs: [run, run],
                        parent: active.last().copied(),
                        domain,
                        domain_end: domain_ends[&domain],
                        simple_fields: Vec::new(),
                        valid: true,
                    });
                    active.push(field);
                }
                Some("separate") => {
                    if let Some(field) = active.last().copied()
                        && fields[field].domain == domain
                        && fields[field].markers[1].replace(index).is_some()
                    {
                        fields[field].valid = false;
                    }
                }
                Some("end") => {
                    if let Some(field) = active.last().copied() {
                        if fields[field].domain == domain {
                            fields[field].markers[2] = Some(index);
                            fields[field].edge_runs[1] = run;
                            active.pop();
                        } else {
                            for field in active.drain(..) {
                                fields[field].valid = false;
                            }
                        }
                    }
                }
                _ => {
                    for field in active.drain(..) {
                        fields[field].valid = false;
                    }
                }
            }
        }
        for active in story_fields.values() {
            for field in active {
                fields[*field].valid = false;
            }
        }
        for field in &mut fields {
            field.valid &= field.markers.iter().all(Option::is_some);
        }
        for index in 0..fields.len() {
            if let Some(parent) = fields[index].parent
                && !fields[parent].valid
            {
                fields[index].valid = false;
            }
        }
        Ok((fields, owners))
    }

    fn carried_simple_field_retirements(&self, state: &RenderState<'_>) -> Result<HashSet<usize>> {
        let (fields, _) = self.carried_field_sources()?;
        let mut has_simple = fields
            .iter()
            .map(|field| !field.simple_fields.is_empty())
            .collect::<Vec<_>>();
        for index in (0..fields.len()).rev() {
            if has_simple[index]
                && let Some(parent) = fields[index].parent
            {
                has_simple[parent] = true;
            }
        }
        let mut watched = vec![false; fields.len()];
        for (index, field) in fields.iter().enumerate() {
            watched[index] = field.valid
                && field.markers.iter().all(Option::is_some)
                && (has_simple[index] || field.parent.is_some_and(|parent| watched[parent]));
        }
        let mut removed = vec![false; self.elements.len()];
        let mut unclaimed_gaps = vec![false; self.elements.len() + 1];
        for element in &self.elements {
            if element.modeled
                && element.word
                && matches!(
                    element.local.as_str(),
                    "p" | "fldSimple" | "hyperlink" | "sdtContent"
                )
            {
                let mut cursor = element.open_end;
                for child in &element.children {
                    let current = &self.elements[*child];
                    unclaimed_gaps[*child] = !self.source[cursor..current.start]
                        .iter()
                        .all(u8::is_ascii_whitespace);
                    cursor = current.end;
                }
                let end = self
                    .elements
                    .partition_point(|child| child.start < element.end);
                unclaimed_gaps[end] |= !self.source[cursor..element.close_start]
                    .iter()
                    .all(u8::is_ascii_whitespace);
            }
        }
        let mut bad = Vec::with_capacity(self.elements.len() + 1);
        bad.push(0usize);
        for (index, element) in self.elements.iter().enumerate() {
            removed[index] = element.parent.is_some_and(|parent| removed[parent])
                || element.revision.as_ref().is_some_and(|metadata| {
                    state.scope.matches(metadata)
                        && removes_content(state.resolution, metadata.kind)
                })
                || self
                    .selected_owner_markers(index, state)
                    .iter()
                    .any(|marker| {
                        removes_content(
                            state.resolution,
                            self.elements[*marker]
                                .revision
                                .as_ref()
                                .expect("actual marker")
                                .kind,
                        )
                    });
            let unselected_revision = element
                .revision
                .as_ref()
                .is_some_and(|metadata| !state.scope.matches(metadata));
            let live_run =
                element.modeled && element.word && element.local == "r" && !removed[index];
            let unclaimed = !removed[index]
                && !element.modeled
                && element.parent.is_some_and(|parent| {
                    let parent = &self.elements[parent];
                    parent.modeled
                        && parent.word
                        && matches!(
                            parent.local.as_str(),
                            "p" | "fldSimple" | "hyperlink" | "sdtContent"
                        )
                        && !(element.word
                            && element.local == "fldData"
                            && parent.local == "fldSimple")
                });
            bad.push(
                bad[index]
                    .checked_add(usize::from(
                        unselected_revision
                            || live_run
                            || unclaimed
                            || unclaimed_gaps[index] && !removed[index],
                    ))
                    .ok_or_else(|| Error::Other("carrier source budget overflow".into()))?,
            );
        }
        // One bottom-up source-material proof per actual XML owner. Prefix
        // budgets separately cover the physical cross-paragraph carrier interval.
        let mut surviving_material = bad
            .windows(2)
            .map(|pair| pair[0] != pair[1])
            .collect::<Vec<_>>();
        for index in (0..self.elements.len()).rev() {
            if surviving_material[index]
                && let Some(parent) = self.elements[index].parent
            {
                surviving_material[parent] = true;
            }
        }
        let mut retiring = HashSet::new();
        let mut removes_field = vec![false; fields.len()];
        for (index, field) in fields.iter().enumerate() {
            if !watched[index] {
                continue;
            }
            let markers = field
                .markers
                .map(|marker| marker.expect("proved complete carrier"));
            let survival = markers.map(|marker| !removed[marker]);
            if survival == [true; 3] {
                continue;
            }
            let end = self
                .elements
                .partition_point(|element| element.start < self.elements[field.edge_runs[1]].end);
            if survival != [false; 3]
                || bad[end] != bad[field.edge_runs[0]]
                || field
                    .simple_fields
                    .iter()
                    .any(|simple| surviving_material[*simple])
            {
                return Err(Error::Other(format!(
                    "selected revision would break complete field carriage at byte {} or leave surviving payload",
                    self.elements[markers[0]].start
                )));
            }
            removes_field[index] = true;
        }
        for (index, field) in fields.iter().enumerate() {
            if field.parent.is_some_and(|parent| removes_field[parent]) {
                removes_field[index] = true;
            }
            if removes_field[index] {
                retiring.extend(field.simple_fields.iter().copied());
            }
        }
        Ok(retiring)
    }

    pub(crate) fn parse(source: &'a [u8]) -> Result<Self> {
        Self::parse_in_scope(source, &NamespaceScope::new())
    }

    fn parse_in_scope(source: &'a [u8], inherited: &NamespaceScope) -> Result<Self> {
        let mut reader = Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        let mut elements: Vec<XmlElement> = Vec::new();
        let mut stack = Vec::new();
        let mut scopes = vec![inherited.clone()];
        let mut root = None;
        let mut buffer = Vec::new();

        loop {
            let before = reader.buffer_position() as usize;
            let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
            let after = reader.buffer_position() as usize;
            match event {
                Event::Start(start) => {
                    let (scope, declarations) =
                        element_scope(&start, scopes.last().expect("scope exists"))?;
                    let index = push_element(
                        &mut elements,
                        &stack,
                        &start,
                        &scope,
                        declarations,
                        before,
                        after,
                        false,
                    )?;
                    root.get_or_insert(index);
                    stack.push(index);
                    scopes.push(scope);
                }
                Event::Empty(start) => {
                    let (scope, declarations) =
                        element_scope(&start, scopes.last().expect("scope exists"))?;
                    let index = push_element(
                        &mut elements,
                        &stack,
                        &start,
                        &scope,
                        declarations,
                        before,
                        after,
                        true,
                    )?;
                    root.get_or_insert(index);
                }
                Event::End(_) => {
                    let index = stack.pop().ok_or_else(|| {
                        Error::Other("XML end element has no matching start".to_owned())
                    })?;
                    scopes.pop();
                    elements[index].close_start = before;
                    elements[index].end = after;
                }
                Event::Text(ref text) => {
                    if let Some(index) = stack.last().copied()
                        && elements[index].word
                        && matches!(elements[index].local.as_str(), "t" | "delText")
                    {
                        elements[index]
                            .snapshot_scalar
                            .push_str(&oxml_core::xml_text::decode_plain(text));
                    }
                }
                Event::GeneralRef(ref entity) => {
                    if let Some(index) = stack.last().copied()
                        && elements[index].word
                        && matches!(elements[index].local.as_str(), "t" | "delText")
                    {
                        elements[index]
                            .snapshot_scalar
                            .push_str(&oxml_core::xml_text::resolve_entity(entity));
                    }
                }
                Event::CData(ref data) => {
                    if let Some(index) = stack.last().copied()
                        && elements[index].word
                        && matches!(elements[index].local.as_str(), "t" | "delText")
                        && let Ok(text) = data.decode()
                    {
                        elements[index].snapshot_scalar.push_str(&text);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }

        if !stack.is_empty() {
            return Err(Error::Other("XML has unclosed elements".to_owned()));
        }
        Ok(Self {
            source,
            elements,
            root: root.ok_or_else(|| Error::Other("XML has no root element".to_owned()))?,
            note_admission: None,
        })
    }

    /// Read a selected custody attribute in this tree's actual ancestor scope.
    pub(crate) fn word_attribute(&self, index: usize, wanted: &str) -> Result<Option<String>> {
        let mut ancestors = Vec::new();
        let mut current = Some(index);
        while let Some(at) = current {
            ancestors.push(at);
            current = self.elements[at].parent;
        }
        let mut scope = NamespaceScope::new();
        for at in ancestors.into_iter().rev() {
            scope.extend(self.elements[at].namespace_declarations.iter().cloned());
        }
        let element = &self.elements[index];
        let mut reader = Reader::from_reader(&self.source[element.start..element.open_end]);
        let (Event::Start(start) | Event::Empty(start)) = reader
            .read_event()
            .map_err(|error| Error::Other(error.to_string()))?
        else {
            return Err(Error::Other("note custody element has no start".into()));
        };
        for attribute in start.attributes() {
            let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
            let key = std::str::from_utf8(attribute.key.as_ref())
                .map_err(|error| Error::Other(error.to_string()))?;
            let Some((prefix, local)) = key.split_once(':') else {
                continue;
            };
            if local == wanted && scope.get(prefix).map(String::as_str) == Some(WORD_NS) {
                return Ok(Some(
                    attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
                        .map_err(|error| Error::Other(error.to_string()))?
                        .into_owned(),
                ));
            }
        }
        Ok(None)
    }

    fn direct_word_child(&self, index: usize, local: &str) -> Result<usize> {
        self.elements[index]
            .children
            .iter()
            .copied()
            .find(|child| self.elements[*child].word && self.elements[*child].local == local)
            .ok_or_else(|| Error::Other(format!("rendered properties have no direct Word {local}")))
    }

    pub(crate) fn source_scope(&self, index: usize) -> NamespaceDeclarations {
        let mut ancestors = Vec::new();
        let mut current = Some(index);
        while let Some(index) = current {
            ancestors.push(index);
            current = self.elements[index].parent;
        }
        ancestors
            .into_iter()
            .rev()
            .fold(Vec::new(), |scope, index| {
                merged_namespaces(&scope, &self.elements[index].namespace_declarations)
            })
    }

    /// Relocation preserves the source context, including prefixes used only
    /// in opaque values. A child-local declaration wins over inherited context.
    /// Identical source and destination bindings need no extra declaration.
    fn relocation_namespaces(
        &self,
        index: usize,
        destination: &[(String, String)],
    ) -> NamespaceDeclarations {
        let element = &self.elements[index];
        let source = element
            .parent
            .map(|parent| self.source_scope(parent))
            .unwrap_or_default();
        let mut changed = source
            .into_iter()
            .filter(|(prefix, value)| {
                prefix != "xml"
                    && !element
                        .namespace_declarations
                        .iter()
                        .any(|(name, _)| name == prefix)
                    && !(prefix.is_empty()
                        && value.is_empty()
                        && !destination.iter().any(|(name, _)| name.is_empty()))
                    && !destination
                        .iter()
                        .any(|(name, target)| name == prefix && target == value)
            })
            .collect::<Vec<_>>();
        if !self
            .source_scope(index)
            .iter()
            .any(|(prefix, _)| prefix.is_empty())
            && destination
                .iter()
                .any(|(prefix, value)| prefix.is_empty() && !value.is_empty())
        {
            changed.push((String::new(), String::new()));
        }
        changed.sort();
        changed
    }

    fn render(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
    ) -> Result<Vec<u8>> {
        self.render_with_namespaces(
            index,
            state,
            convert_deleted_text,
            &self.elements[index]
                .parent
                .map(|parent| self.source_scope(parent))
                .unwrap_or_default(),
        )
    }

    fn render_with_namespaces(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[index];
        let semantic = self.semantic_element(index);
        if semantic && state.retiring_simple_fields.contains(&index) {
            self.validate_selected_descendants(index, state)?;
            return Ok(Vec::new());
        }
        if let Some(metadata) = &element.revision
            && semantic
            && state.scope.matches(metadata)
        {
            state.resolved.insert(index);
            return self.render_selected_revision(index, metadata.kind, state, destination);
        }

        if semantic && matches!(state.resolution, Resolution::Reject) {
            let changes = self.selected_property_changes(index, state)?;
            if let Some((_, prior)) = changes.first().copied() {
                for (selected, _) in &changes {
                    state.resolved.insert(*selected);
                    self.validate_selected_descendants(*selected, state)?;
                }

                if !state.discarded
                    && state.leaf_elements.iter().any(|leaf| {
                        self.elements[*leaf].start >= element.start
                            && self.elements[*leaf].end <= element.end
                    })
                {
                    return Err(Error::Other(
                        "comparison leaf transport does not admit property replacement".into(),
                    ));
                }
                return self.render_rejected_property_owner(
                    index,
                    prior,
                    &changes,
                    state,
                    destination,
                );
            }
        }

        let markers = if semantic {
            self.selected_owner_markers(index, state)
        } else {
            Vec::new()
        };
        if markers.iter().any(|marker| {
            let kind = self.elements[*marker]
                .revision
                .as_ref()
                .expect("selected marker is a revision")
                .kind;
            removes_content(state.resolution, kind)
        }) {
            if element.local == "numPr" {
                return self.render_rejected_numbering_owner(index, &markers, state, destination);
            }
            self.validate_selected_descendants(index, state)?;
            return Ok(Vec::new());
        }

        self.render_ordinary(index, state, convert_deleted_text, destination)
    }

    fn render_selected_revision(
        &self,
        index: usize,
        kind: RevisionKind,
        state: &mut RenderState<'_>,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        if is_property_change(kind) {
            let prior = self.prior_property(index, kind)?;
            if matches!(state.resolution, Resolution::Accept) {
                return Ok(Vec::new());
            }

            return self.render_with_namespaces(prior, state, false, destination);
        }

        if self.is_contextual_marker(index) {
            return Ok(Vec::new());
        }

        if keeps_content(state.resolution, kind) {
            self.render_inner_with_namespaces(
                index,
                state,
                matches!(state.resolution, Resolution::Reject)
                    && matches!(kind, RevisionKind::Deletion | RevisionKind::MoveFrom),
                destination,
            )
        } else {
            self.validate_selected_descendants(index, state)?;
            Ok(Vec::new())
        }
    }

    fn render_ordinary(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let leaf_mark = state.leaf_updates.len();
        let element = &self.elements[index];
        let declarations = self.relocation_namespaces(index, destination);
        let current_scope = merged_namespaces(
            &merged_namespaces(destination, &declarations),
            &element.namespace_declarations,
        );
        // Rejected deleted text reads as text again, and a deleted field
        // instruction as an instruction, only in its admitted source role.
        let semantic = self.semantic_element(index);
        let restored_local = match element.local.as_str() {
            "delText" if semantic && convert_deleted_text && element.word => Some("t"),
            "delInstrText" if semantic && convert_deleted_text && element.word => Some("instrText"),
            _ => None,
        };
        if element.empty {
            let raw = &self.source[element.start..element.end];
            let raw =
                inject_namespace_declarations(raw, &element.namespace_declarations, &declarations);
            let output = match restored_local {
                Some(local) => rename_element(&raw, &element.name, local),
                None => raw,
            };
            if !state.discarded
                && let Some(leaf) = state
                    .leaf_elements
                    .iter()
                    .position(|selected| *selected == index)
            {
                state.leaf_ranges[leaf] = Some(0..output.len());
                state.leaf_updates.push(leaf);
            }
            return Ok(output);
        }

        let resolves_paragraph_property_change = semantic
            && element.word
            && element.local == "pPr"
            && !self.selected_property_changes(index, state)?.is_empty();
        let removes_empty_property_shell = semantic
            && element.word
            && matches!(element.local.as_str(), "pPr" | "rPr" | "trPr" | "numPr")
            && (resolves_paragraph_property_change
                || self.subtree_contains_selected_contextual_marker(index, state));
        let inner =
            self.render_inner_with_namespaces(index, state, convert_deleted_text, &current_scope)?;
        if semantic
            && element.word
            && element.local == "tbl"
            && self.owned_rows_all_remove(index, state)
        {
            return Ok(Vec::new());
        }
        if semantic
            && element.word
            && element.local == "hyperlink"
            && inner.iter().all(u8::is_ascii_whitespace)
            && self.owned_content_all_removes(index, state)
        {
            return Ok(Vec::new());
        }
        if removes_empty_property_shell
            && inner.iter().all(u8::is_ascii_whitespace)
            && !self.element_has_attributes(index)?
            && element.namespace_declarations.is_empty()
        {
            return Ok(Vec::new());
        }

        let mut output = Vec::with_capacity(element.end - element.start);
        let open = inject_namespace_declarations(
            &self.source[element.start..element.open_end],
            &element.namespace_declarations,
            &declarations,
        );
        if let Some(local) = restored_local {
            output.extend_from_slice(&rename_element(&open, &element.name, local));
        } else {
            output.extend_from_slice(&open);
        }
        state.shift_leaves(leaf_mark, output.len());
        output.extend_from_slice(&inner);
        let close = &self.source[element.close_start..element.end];
        if let Some(local) = restored_local {
            output.extend_from_slice(&rename_element(close, &element.name, local));
        } else {
            output.extend_from_slice(close);
        }
        if !state.discarded
            && let Some(leaf) = state
                .leaf_elements
                .iter()
                .position(|selected| *selected == index)
        {
            state.leaf_ranges[leaf] = Some(0..output.len());
            state.leaf_updates.push(leaf);
        }
        Ok(output)
    }

    fn render_inner_with_namespaces(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[index];
        let mut output = Vec::new();
        let mut cursor = element.open_end;
        let mut child_index = 0;
        while child_index < element.children.len() {
            let child = element.children[child_index];
            let child_element = &self.elements[child];
            output.extend_from_slice(&self.source[cursor..child_element.start]);
            if self.semantic_element(child)
                && child_element.word
                && child_element.local == "p"
                && !state.discarded
                && self.paragraph_mark_removes(child, state)
            {
                if !state.leaf_elements.is_empty() {
                    return Err(Error::Other(
                        "comparison leaf transport does not admit paragraph merging".into(),
                    ));
                }
                let mut paragraphs = vec![child];
                // A table whose every row is removed is crossed, so the
                // paragraphs on either side of it merge. Rendering it resolves
                // its row markers and yields no bytes.
                let mut siblings = vec![child];
                let mut next_index = child_index + 1;
                // A table that stays leaves the paragraphs nothing to merge
                // into, so they go, which needs them to keep no content.
                let mut before_table = false;
                loop {
                    let next = element.children.get(next_index).copied().ok_or_else(|| {
                        Error::Other("cannot remove the final paragraph mark".to_owned())
                    })?;
                    let next_element = &self.elements[next];
                    if self.semantic_element(next)
                        && next_element.word
                        && next_element.local == "tbl"
                    {
                        if !self.owned_rows_all_remove(next, state) {
                            before_table = true;
                            break;
                        }
                        siblings.push(next);
                        next_index += 1;
                        continue;
                    }
                    if !self.semantic_element(next)
                        || !next_element.word
                        || next_element.local != "p"
                    {
                        return Err(Error::Other(
                            "paragraph mark removal requires an adjacent paragraph".to_owned(),
                        ));
                    }
                    siblings.push(next);
                    next_index += 1;
                    paragraphs.push(next);
                    if !self.paragraph_mark_removes(next, state) {
                        break;
                    }
                }
                for pair in siblings.windows(2) {
                    output.extend_from_slice(
                        &self.source[self.elements[pair[0]].end..self.elements[pair[1]].start],
                    );
                    if self.elements[pair[1]].local == "tbl" {
                        self.render(pair[1], state, false)?;
                    }
                }
                if before_table {
                    self.validate_removed_paragraphs(&paragraphs, state)?;
                } else {
                    output.extend_from_slice(&self.render_merged_paragraphs(
                        &paragraphs,
                        state,
                        destination,
                    )?);
                }
                cursor = self.elements[*siblings.last().expect("paragraph chain exists")].end;
                child_index = next_index;
                continue;
            }
            let leaf_mark = state.leaf_updates.len();
            let rendered =
                self.render_with_namespaces(child, state, convert_deleted_text, destination)?;
            state.shift_leaves(leaf_mark, output.len());
            output.extend_from_slice(&rendered);
            cursor = child_element.end;
            child_index += 1;
        }
        output.extend_from_slice(&self.source[cursor..element.close_start]);
        Ok(output)
    }

    fn selected_property_changes(
        &self,
        index: usize,
        state: &RenderState<'_>,
    ) -> Result<Vec<(usize, usize)>> {
        let element = &self.elements[index];
        let expected = match element.local.as_str() {
            "rPr" => RevisionKind::RunPropertyChange,
            "pPr" => RevisionKind::ParagraphPropertyChange,
            "tblPr" => RevisionKind::TablePropertyChange,
            "sectPr" => RevisionKind::SectionPropertyChange,
            _ => return Ok(Vec::new()),
        };
        element
            .children
            .iter()
            .filter_map(|child| {
                let metadata = self.elements[*child].revision.as_ref()?;
                (metadata.kind == expected && state.scope.matches(metadata)).then_some(*child)
            })
            .map(|change| {
                self.prior_property(change, expected)
                    .map(|prior| (change, prior))
            })
            .collect()
    }

    fn render_rejected_property_owner(
        &self,
        owner: usize,
        prior: usize,
        selected: &[(usize, usize)],
        state: &mut RenderState<'_>,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let mut prior_xml = self.render_with_namespaces(prior, state, false, destination)?;
        let prior_scope = rendered_namespace_scope(&prior_xml, destination, None)?;
        if let Some(markers) =
            self.legacy_mark_markers(owner, prior, state, &prior_xml, destination)?
            && !markers.is_empty()
        {
            prior_xml = insert_first_children(&prior_xml, "rPr", &markers, destination)?;
        }
        let selected = selected
            .iter()
            .map(|(change, _)| *change)
            .collect::<HashSet<_>>();
        let retained = self.render_retained_owner_children(
            owner,
            Some(prior),
            &selected,
            state,
            &prior_scope,
        )?;
        let prior_element = &self.elements[prior];
        if prior_element.word
            && matches!(prior_element.local.as_str(), "pPr" | "tblPr")
            && retained.is_empty()
            && !self.element_has_attributes(prior)?
            && (prior_element.empty
                || self.source[prior_element.open_end..prior_element.close_start]
                    .iter()
                    .all(u8::is_ascii_whitespace))
        {
            return Ok(Vec::new());
        }
        if prior_element.word && prior_element.local == "rPr" && !retained.is_empty() {
            let tree = XmlTree::parse_in_scope(&prior_xml, &prior_scope.iter().cloned().collect())?;
            // Retained markers precede any still-unselected prior rPrChange.
            if let Ok(change) = tree.direct_word_child(tree.root, "rPrChange") {
                let at = tree.elements[change].start;
                let mut output = prior_xml[..at].to_vec();
                output.extend_from_slice(&retained);
                output.extend_from_slice(&prior_xml[at..]);
                return Ok(output);
            }
        }
        append_children_to_element(&prior_xml, &retained)
    }

    fn element_has_attributes(&self, index: usize) -> Result<bool> {
        let element = &self.elements[index];
        let mut reader = Reader::from_reader(&self.source[element.start..element.open_end]);
        let mut buffer = Vec::new();
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(start) | Event::Empty(start) => Ok(start.attributes().next().is_some()),
            _ => Err(Error::Other(
                "property owner XML has no start element".to_owned(),
            )),
        }
    }

    fn render_rejected_numbering_owner(
        &self,
        owner: usize,
        selected: &[usize],
        state: &mut RenderState<'_>,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let selected = selected.iter().copied().collect::<HashSet<_>>();
        for marker in &selected {
            state.resolved.insert(*marker);
            self.validate_selected_descendants(*marker, state)?;
        }
        let element = &self.elements[owner];
        let declarations = self.relocation_namespaces(owner, destination);
        let current_scope = merged_namespaces(
            &merged_namespaces(destination, &declarations),
            &element.namespace_declarations,
        );
        let retained =
            self.render_retained_owner_children(owner, None, &selected, state, &current_scope)?;
        if retained.is_empty() {
            return Ok(Vec::new());
        }
        let mut output = inject_namespace_declarations(
            &self.source[element.start..element.open_end],
            &element.namespace_declarations,
            &declarations,
        );
        output.extend_from_slice(&retained);
        output.extend_from_slice(&self.source[element.close_start..element.end]);
        Ok(output)
    }

    /// The prior mark `w:rPr` of an older rdocx `w:pPrChange`, and the
    /// paragraph-mark markers of the current mark that the rejection leaves
    /// in place, which go first in the prior mark.
    ///
    /// Those redlines put the prior mark formatting in the prior properties,
    /// where it replaces the current mark formatting. A mark formatting change
    /// left in place has no owner then, so it is refused.
    fn legacy_mark_markers(
        &self,
        owner: usize,
        prior: usize,
        state: &mut RenderState<'_>,
        prior_xml: &[u8],
        destination: &[(String, String)],
    ) -> Result<Option<Vec<u8>>> {
        let mark = |properties: usize| {
            self.elements[properties]
                .children
                .iter()
                .copied()
                .find(|child| self.elements[*child].word && self.elements[*child].local == "rPr")
        };
        let owner_element = &self.elements[owner];
        if !owner_element.word || owner_element.local != "pPr" {
            return Ok(None);
        }
        let (Some(current), Some(_)) = (mark(owner), mark(prior)) else {
            return Ok(None);
        };
        let namespaces = rendered_namespace_scope(prior_xml, destination, Some("rPr"))?;
        let mut markers = Vec::new();
        for child in &self.elements[current].children {
            let rendered = self.render_with_namespaces(*child, state, false, &namespaces)?;
            match &self.elements[*child].revision {
                Some(metadata)
                    if matches!(
                        metadata.kind,
                        RevisionKind::Insertion
                            | RevisionKind::Deletion
                            | RevisionKind::MoveFrom
                            | RevisionKind::MoveTo
                    ) =>
                {
                    markers.extend_from_slice(&rendered);
                }
                Some(metadata)
                    if metadata.kind == RevisionKind::RunPropertyChange
                        && !state.scope.matches(metadata) =>
                {
                    return Err(Error::Other(
                        "rejecting a paragraph property change whose prior properties hold the \
                         mark formatting cannot keep a mark formatting change"
                            .to_owned(),
                    ));
                }
                _ => {}
            }
        }
        Ok(Some(markers))
    }

    /// Render the children of a property owner that its rejected prior value
    /// does not replace.
    ///
    /// A `w:pPrChange` holds only the prior base paragraph properties, so the
    /// paragraph mark `w:rPr` and the `w:sectPr` stay and resolve their own
    /// revisions. Older rdocx redlines put the prior mark formatting in the
    /// prior properties, and that `w:rPr` replaces the current one, see
    /// [`Self::legacy_mark_markers`].
    fn render_retained_owner_children(
        &self,
        owner: usize,
        prior: Option<usize>,
        selected: &HashSet<usize>,
        state: &mut RenderState<'_>,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[owner];
        // A complete prior rPr replaces run or paragraph-mark formatting,
        // including opaque children. Independently scoped revision markers
        // remain current children and resolve through the normal selection.
        let replaces_run_properties = prior.is_some() && element.word && element.local == "rPr";
        let mut output = Vec::new();
        for child in &element.children {
            let child_element = &self.elements[*child];
            if replaces_run_properties
                && child_element.revision.is_none()
                && !retained_revision_local(&element.local, &child_element.local)
            {
                continue;
            }
            if let Some(prior) = prior
                && element.word
                && element.local == "pPr"
                && child_element.word
                && matches!(child_element.local.as_str(), "rPr" | "sectPr")
            {
                let replaced = child_element.local == "rPr"
                    && self.elements[prior].children.iter().any(|replacement| {
                        self.elements[*replacement].word
                            && self.elements[*replacement].local == "rPr"
                    });
                if !replaced {
                    output.extend_from_slice(&self.render_with_namespaces(
                        *child,
                        state,
                        false,
                        destination,
                    )?);
                }
                continue;
            }
            if !selected.contains(child)
                && (child_element.revision.is_some()
                    || !child_element.word
                    || retained_revision_local(&element.local, &child_element.local)
                    || self.subtree_contains_revision(*child))
            {
                output.extend_from_slice(&self.render_with_namespaces(
                    *child,
                    state,
                    false,
                    destination,
                )?);
            }
        }
        Ok(output)
    }

    fn subtree_contains_revision(&self, index: usize) -> bool {
        self.elements[index].children.iter().any(|child| {
            self.elements[*child].revision.is_some() || self.subtree_contains_revision(*child)
        })
    }

    fn subtree_contains_selected_contextual_marker(
        &self,
        index: usize,
        state: &RenderState<'_>,
    ) -> bool {
        self.elements[index].children.iter().any(|child| {
            self.elements[*child]
                .revision
                .as_ref()
                .is_some_and(|metadata| {
                    matches!(
                        metadata.kind,
                        RevisionKind::Insertion
                            | RevisionKind::Deletion
                            | RevisionKind::MoveFrom
                            | RevisionKind::MoveTo
                    ) && state.scope.matches(metadata)
                })
                || self.subtree_contains_selected_contextual_marker(*child, state)
        })
    }

    fn owned_rows_all_remove(&self, index: usize, state: &RenderState<'_>) -> bool {
        let mut rows = Vec::new();
        for child in &self.elements[index].children {
            self.collect_owned_rows(*child, &mut rows);
        }
        !rows.is_empty()
            && rows.iter().all(|row| {
                self.selected_owner_markers(*row, state)
                    .iter()
                    .any(|marker| {
                        let kind = self.elements[*marker]
                            .revision
                            .as_ref()
                            .expect("row marker is a revision")
                            .kind;
                        removes_content(state.resolution, kind)
                    })
            })
    }

    /// Whether the resolution removes every child of a hyperlink, each a
    /// revision wrapper, which would leave the hyperlink empty. Word reads a
    /// hyperlink as a field, and an empty one shows nothing.
    fn owned_content_all_removes(&self, index: usize, state: &RenderState<'_>) -> bool {
        let children = &self.elements[index].children;
        !children.is_empty()
            && children.iter().all(|child| {
                state.retiring_simple_fields.contains(child)
                    || self.elements[*child]
                        .revision
                        .as_ref()
                        .is_some_and(|metadata| {
                            state.scope.matches(metadata)
                                && removes_content(state.resolution, metadata.kind)
                        })
            })
    }

    fn collect_owned_rows(&self, index: usize, rows: &mut Vec<usize>) {
        let element = &self.elements[index];
        if element.word && element.local == "tbl" {
            return;
        }
        if element.word && element.local == "tr" {
            rows.push(index);
            return;
        }
        for child in &element.children {
            self.collect_owned_rows(*child, rows);
        }
    }

    fn prior_property(&self, change: usize, kind: RevisionKind) -> Result<usize> {
        let expected = match kind {
            RevisionKind::RunPropertyChange => "rPr",
            RevisionKind::ParagraphPropertyChange => "pPr",
            RevisionKind::TablePropertyChange => "tblPr",
            RevisionKind::SectionPropertyChange => "sectPr",
            _ => {
                return Err(Error::Other(
                    "selected revision is not a property change".to_owned(),
                ));
            }
        };
        let children = &self.elements[change].children;
        if children.len() != 1 {
            return Err(Error::Other(format!(
                "selected property revision requires exactly one prior w:{expected} value"
            )));
        }
        children
            .first()
            .copied()
            .filter(|prior| self.elements[*prior].word && self.elements[*prior].local == expected)
            .ok_or_else(|| {
                Error::Other(format!(
                    "selected property revision requires a prior w:{expected} value"
                ))
            })
    }

    fn selected_owner_markers(&self, index: usize, state: &RenderState<'_>) -> Vec<usize> {
        let owner = &self.elements[index];
        let property_local = match owner.local.as_str() {
            "r" => "rPr",
            "tr" => "trPr",
            "numPr" => "numPr",
            _ => return Vec::new(),
        };
        let property = if owner.local == "numPr" {
            index
        } else {
            let Some(property) = owner.children.iter().find(|child| {
                self.elements[**child].word && self.elements[**child].local == property_local
            }) else {
                return Vec::new();
            };
            *property
        };
        self.elements[property]
            .children
            .iter()
            .filter_map(|child| {
                let metadata = self.elements[*child].revision.as_ref()?;
                (matches!(
                    metadata.kind,
                    RevisionKind::Insertion
                        | RevisionKind::Deletion
                        | RevisionKind::MoveFrom
                        | RevisionKind::MoveTo
                ) && state.scope.matches(metadata))
                .then_some(*child)
            })
            .collect()
    }

    fn is_contextual_marker(&self, index: usize) -> bool {
        self.elements[index].parent.is_some_and(|parent| {
            matches!(
                self.elements[parent].local.as_str(),
                "rPr" | "trPr" | "numPr"
            )
        })
    }

    fn validate_selected_descendants(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
    ) -> Result<()> {
        let discarded = std::mem::replace(&mut state.discarded, true);
        let validated = self.elements[index]
            .children
            .iter()
            .try_for_each(|child| self.render(*child, state, false).map(drop));
        state.discarded = discarded;
        validated
    }

    fn paragraph_mark_removes(&self, index: usize, state: &RenderState<'_>) -> bool {
        self.paragraph_markers(index, state).iter().any(|marker| {
            let kind = self.elements[*marker]
                .revision
                .as_ref()
                .expect("paragraph marker is a revision")
                .kind;
            removes_content(state.resolution, kind)
        })
    }

    fn paragraph_markers(&self, index: usize, state: &RenderState<'_>) -> Vec<usize> {
        let paragraph = &self.elements[index];
        let Some(properties) = paragraph
            .children
            .iter()
            .find(|child| self.elements[**child].word && self.elements[**child].local == "pPr")
        else {
            return Vec::new();
        };
        let Some(run_properties) = self.elements[*properties]
            .children
            .iter()
            .find(|child| self.elements[**child].word && self.elements[**child].local == "rPr")
        else {
            return Vec::new();
        };
        self.elements[*run_properties]
            .children
            .iter()
            .filter_map(|child| {
                let metadata = self.elements[*child].revision.as_ref()?;
                (matches!(
                    metadata.kind,
                    RevisionKind::Insertion
                        | RevisionKind::Deletion
                        | RevisionKind::MoveFrom
                        | RevisionKind::MoveTo
                ) && state.scope.matches(metadata))
                .then_some(*child)
            })
            .collect()
    }

    /// Resolve paragraphs whose marks are removed before a table that stays,
    /// and which go. Nothing of them but proofing marks may remain.
    fn validate_removed_paragraphs(
        &self,
        paragraphs: &[usize],
        state: &mut RenderState<'_>,
    ) -> Result<()> {
        for paragraph in paragraphs {
            let owner = &self.elements[*paragraph];
            let mut cursor = owner.open_end;
            for child in &owner.children {
                let element = &self.elements[*child];
                if !self.source[cursor..element.start]
                    .iter()
                    .all(u8::is_ascii_whitespace)
                {
                    return Err(Error::Other(
                        "paragraph mark removal before a table requires a paragraph with no content left".to_owned(),
                    ));
                }
                cursor = element.end;
            }
            if !self.source[cursor..owner.close_start]
                .iter()
                .all(u8::is_ascii_whitespace)
            {
                return Err(Error::Other(
                    "paragraph mark removal before a table requires a paragraph with no content left".to_owned(),
                ));
            }
            for child in &owner.children {
                let child_element = &self.elements[*child];
                if child_element.word && child_element.local == "pPr" {
                    self.validate_selected_descendants(*child, state)?;
                } else if !self.render(*child, state, false)?.is_empty()
                    && !(child_element.word && child_element.local == "proofErr")
                {
                    return Err(Error::Other(
                        "paragraph mark removal before a table requires a paragraph with no \
                         content left"
                            .to_owned(),
                    ));
                }
            }
        }
        Ok(())
    }

    fn render_merged_paragraphs(
        &self,
        paragraphs: &[usize],
        state: &mut RenderState<'_>,
        destination: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let current = paragraphs[0];
        let next = *paragraphs.last().expect("paragraph chain exists");
        let current_element = &self.elements[current];
        let next_element = &self.elements[next];
        let declarations = self.relocation_namespaces(current, destination);
        let current_scope = merged_namespaces(
            &merged_namespaces(destination, &declarations),
            &current_element.namespace_declarations,
        );
        let mut output = inject_namespace_declarations(
            &self.source[current_element.start..current_element.open_end],
            &current_element.namespace_declarations,
            &declarations,
        );
        if let Some(properties) = next_element
            .children
            .iter()
            .find(|child| self.elements[**child].word && self.elements[**child].local == "pPr")
        {
            output.extend_from_slice(&self.render_with_namespaces(
                *properties,
                state,
                false,
                &current_scope,
            )?);
        }
        for paragraph in paragraphs {
            for child in &self.elements[*paragraph].children {
                let child_element = &self.elements[*child];
                if child_element.word && child_element.local == "pPr" {
                    if *paragraph != next {
                        self.validate_selected_descendants(*child, state)?;
                    }
                } else {
                    output.extend_from_slice(&self.render_with_namespaces(
                        *child,
                        state,
                        false,
                        &current_scope,
                    )?);
                }
            }
        }
        output.extend_from_slice(&self.source[current_element.close_start..current_element.end]);
        Ok(output)
    }
}

fn rendered_namespace_scope(
    xml: &[u8],
    destination: &[(String, String)],
    local: Option<&str>,
) -> Result<NamespaceDeclarations> {
    if xml.is_empty() {
        return Ok(destination.to_vec());
    }
    let tree = XmlTree::parse_in_scope(xml, &destination.iter().cloned().collect())?;
    let index = match local {
        Some(local) => tree.direct_word_child(tree.root, local)?,
        None => tree.root,
    };
    Ok(merged_namespaces(destination, &tree.source_scope(index)))
}

fn push_element(
    elements: &mut Vec<XmlElement>,
    stack: &[usize],
    start: &BytesStart<'_>,
    scope: &NamespaceScope,
    namespace_declarations: NamespaceDeclarations,
    offset: usize,
    end: usize,
    empty: bool,
) -> Result<usize> {
    let name = std::str::from_utf8(start.name().as_ref())
        .map_err(|error| Error::Other(error.to_string()))?
        .to_owned();
    let local = std::str::from_utf8(start.local_name().as_ref())
        .map_err(|error| Error::Other(error.to_string()))?
        .to_owned();
    let word = element_namespace(&name, scope) == Some(WORD_NS);
    let parent = stack.last().copied();
    let mut modeled = (word && local == "txbxContent")
        || parent.map_or(
            word && matches!(
                local.as_str(),
                "document" | "hdr" | "ftr" | "comments" | "footnotes" | "endnotes"
            ),
            |parent| modeled_child(&elements[parent], word, &local),
        );
    if word && local == "fldSimple" {
        modeled = parent.is_some_and(|index| {
            let parent = &elements[index];
            parent.modeled
                && parent.word
                && match parent.local.as_str() {
                    "p" | "fldSimple" | "hyperlink" => true,
                    "sdtContent" => inline_field_control_owner(elements, index),
                    _ => false,
                }
        });
    }
    let revision = modeled
        .then(|| revision_metadata(start, &local, scope))
        .transpose()?
        .flatten();
    if modeled && is_revision_local(&local) && revision.is_none() {
        modeled = false;
    }
    let mut snapshot_field_marker = None;
    if word
        && local == "fldChar"
        && parent.is_some_and(|parent| elements[parent].word && elements[parent].local == "r")
    {
        for attribute in start.attributes() {
            let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
            let key = std::str::from_utf8(attribute.key.as_ref())
                .map_err(|error| Error::Other(error.to_string()))?;
            if let Some((prefix, "fldCharType")) = key.split_once(':')
                && scope.get(prefix).is_some_and(|uri| uri == WORD_NS)
            {
                if snapshot_field_marker.is_some() {
                    snapshot_field_marker = Some(String::new());
                    break;
                }
                snapshot_field_marker = match attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
                {
                    Ok(value) => Some(value.into_owned()),
                    Err(error) if parent.is_some_and(|parent| elements[parent].modeled) => {
                        return Err(xml_error(error));
                    }
                    // Read-only observations do not add resolver errors for an
                    // otherwise opaque run. Its existing admission stays intact.
                    Err(_) => None,
                };
            }
        }
    }
    let index = elements.len();
    elements.push(XmlElement {
        start: offset,
        open_end: end,
        close_start: end,
        end,
        name,
        local,
        word,
        modeled,
        empty,
        parent,
        children: Vec::new(),
        revision,
        namespace_declarations,
        field_marker: parent
            .filter(|parent| elements[*parent].modeled)
            .and(snapshot_field_marker.clone()),
        snapshot_field_marker,
        snapshot_scalar: String::new(),
    });
    if let Some(parent) = parent {
        elements[parent].children.push(index);
    }
    Ok(index)
}

fn element_scope(
    start: &BytesStart<'_>,
    inherited: &NamespaceScope,
) -> Result<(NamespaceScope, NamespaceDeclarations)> {
    let mut scope = inherited.clone();
    let mut declarations = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Other(error.to_string()))?;
        if key == "xmlns" || key.starts_with("xmlns:") {
            let prefix = key.strip_prefix("xmlns:").unwrap_or("").to_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
                .map_err(xml_error)?
                .into_owned();
            scope.insert(prefix.clone(), value.clone());
            declarations.push((prefix, value));
        }
    }
    Ok((scope, declarations))
}

fn inline_field_control_owner(elements: &[XmlElement], mut index: usize) -> bool {
    loop {
        let element = &elements[index];
        if !element.modeled || !element.word {
            return false;
        }
        match element.local.as_str() {
            "p" | "fldSimple" | "hyperlink" => return true,
            "sdt" | "sdtContent" | "ins" | "del" | "moveFrom" | "moveTo" => {}
            _ => return false,
        }
        let Some(parent) = element.parent else {
            return false;
        };
        index = parent;
    }
}

fn modeled_child(parent: &XmlElement, word: bool, local: &str) -> bool {
    if !parent.modeled || !word {
        return false;
    }
    match parent.local.as_str() {
        "document" => local == "body",
        "body" => matches!(local, "p" | "tbl" | "sdt" | "sectPr"),
        "hdr" | "ftr" | "comment" | "footnote" | "endnote" | "txbxContent" => {
            matches!(local, "p" | "tbl" | "sdt")
        }
        "comments" => local == "comment",
        "footnotes" => local == "footnote",
        "endnotes" => local == "endnote",
        "p" => matches!(
            local,
            "pPr" | "r" | "hyperlink" | "sdt" | "ins" | "del" | "moveFrom" | "moveTo"
        ),
        "hyperlink" => matches!(local, "r" | "ins" | "del" | "moveFrom" | "moveTo"),
        "fldSimple" => matches!(
            local,
            "r" | "fldSimple" | "hyperlink" | "sdt" | "ins" | "del" | "moveFrom" | "moveTo"
        ),
        "pPr" => matches!(local, "rPr" | "numPr" | "sectPr" | "pPrChange"),
        "r" => local == "rPr",
        "rPr" => matches!(local, "ins" | "del" | "moveFrom" | "moveTo" | "rPrChange"),
        "tbl" => matches!(local, "tblPr" | "tblGrid" | "tr" | "sdt"),
        "tblPr" => local == "tblPrChange",
        "tr" => matches!(local, "trPr" | "tc" | "sdt"),
        "trPr" => matches!(local, "ins" | "del" | "moveFrom" | "moveTo"),
        "tc" => matches!(local, "tcPr" | "p" | "tbl" | "sdt"),
        "sdt" => matches!(local, "sdtPr" | "sdtContent"),
        "sdtContent" => matches!(
            local,
            "p" | "tbl"
                | "tr"
                | "tc"
                | "r"
                | "sdt"
                | "ins"
                | "del"
                | "moveFrom"
                | "moveTo"
                | "rPrChange"
                | "pPrChange"
                | "tblPrChange"
                | "sectPrChange"
        ),
        "sectPr" => local == "sectPrChange",
        "numPr" => local == "ins",
        "ins" | "del" | "moveFrom" | "moveTo" => matches!(
            local,
            "r" | "hyperlink" | "ins" | "del" | "moveFrom" | "moveTo"
        ),
        "rPrChange" | "pPrChange" | "tblPrChange" | "sectPrChange" => false,
        _ => false,
    }
}

fn merged_namespaces(
    inherited: &[(String, String)],
    local: &[(String, String)],
) -> Vec<(String, String)> {
    let mut merged = inherited.to_vec();
    for (prefix, value) in local {
        if let Some((_, existing)) = merged.iter_mut().find(|(name, _)| name == prefix) {
            *existing = value.clone();
        } else {
            merged.push((prefix.clone(), value.clone()));
        }
    }
    merged
}

/// Insert children at the actual direct Word property owner, keeping its QName.
fn insert_first_children(
    xml: &[u8],
    local: &str,
    children: &[u8],
    destination: &[(String, String)],
) -> Result<Vec<u8>> {
    let tree = XmlTree::parse_in_scope(xml, &destination.iter().cloned().collect())?;
    let index = tree.direct_word_child(tree.root, local)?;
    let element = &tree.elements[index];
    let mut output = Vec::with_capacity(xml.len() + children.len() + element.name.len() + 3);
    if element.empty {
        output.extend_from_slice(&xml[..element.open_end - 2]);
        output.push(b'>');
        output.extend_from_slice(children);
        output.extend_from_slice(format!("</{}>", element.name).as_bytes());
    } else {
        output.extend_from_slice(&xml[..element.open_end]);
        output.extend_from_slice(children);
    }
    output.extend_from_slice(&xml[element.open_end..]);
    Ok(output)
}

fn append_children_to_element(element_xml: &[u8], children: &[u8]) -> Result<Vec<u8>> {
    if children.is_empty() {
        return Ok(element_xml.to_vec());
    }

    let mut reader = Reader::from_reader(element_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(_) => {
                depth += 1;
            }
            Event::Empty(element) if depth == 0 => {
                let end = reader.buffer_position() as usize;
                let slash = element_xml[..end]
                    .iter()
                    .rposition(|byte| *byte == b'/')
                    .ok_or_else(|| {
                        Error::Other("empty XML element has no closing slash".to_owned())
                    })?;
                let name = element.name();
                let mut output = element_xml[..slash].to_vec();
                output.push(b'>');
                output.extend_from_slice(children);
                output.extend_from_slice(b"</");
                output.extend_from_slice(name.as_ref());
                output.push(b'>');
                output.extend_from_slice(&element_xml[end..]);
                return Ok(output);
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    let mut output = element_xml[..before].to_vec();
                    output.extend_from_slice(children);
                    output.extend_from_slice(&element_xml[before..]);
                    return Ok(output);
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "property XML ended before its root element closed".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn retained_revision_local(owner: &str, child: &str) -> bool {
    match owner {
        "pPr" => child == "pPrChange",
        "rPr" => matches!(child, "rPrChange" | "ins" | "del"),
        "tblPr" => child == "tblPrChange",
        "sectPr" => child == "sectPrChange",
        "numPr" => child == "ins",
        _ => false,
    }
}

fn inject_namespace_declarations(
    open: &[u8],
    local: &[(String, String)],
    promoted: &[(String, String)],
) -> Vec<u8> {
    let missing = promoted
        .iter()
        .filter(|(prefix, _)| !local.iter().any(|(name, _)| name == prefix))
        .collect::<Vec<_>>();
    if missing.is_empty() {
        return open.to_vec();
    }

    let insertion = open
        .iter()
        .rposition(|byte| *byte == b'>')
        .map(|position| {
            if position > 0 && open[position - 1] == b'/' {
                position - 1
            } else {
                position
            }
        })
        .unwrap_or(open.len());
    let mut output = Vec::with_capacity(open.len() + missing.len() * 24);
    output.extend_from_slice(&open[..insertion]);
    for (prefix, value) in missing {
        if prefix.is_empty() {
            output.extend_from_slice(b" xmlns=\"");
        } else {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
        }
        write_xml_attribute_value(&mut output, value);
        output.push(b'"');
    }
    output.extend_from_slice(&open[insertion..]);
    output
}

fn write_xml_attribute_value(output: &mut Vec<u8>, value: &str) {
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            _ => output.push(byte),
        }
    }
}

fn element_namespace<'a>(name: &str, scope: &'a NamespaceScope) -> Option<&'a str> {
    let prefix = name.split_once(':').map(|(prefix, _)| prefix).unwrap_or("");
    scope.get(prefix).map(String::as_str)
}

fn revision_metadata(
    start: &BytesStart<'_>,
    local: &str,
    scope: &NamespaceScope,
) -> Result<Option<RevisionMetadata>> {
    let kind = match local {
        "ins" => RevisionKind::Insertion,
        "del" => RevisionKind::Deletion,
        "moveFrom" => RevisionKind::MoveFrom,
        "moveTo" => RevisionKind::MoveTo,
        "rPrChange" => RevisionKind::RunPropertyChange,
        "pPrChange" => RevisionKind::ParagraphPropertyChange,
        "tblPrChange" => RevisionKind::TablePropertyChange,
        "sectPrChange" => RevisionKind::SectionPropertyChange,
        _ => return Ok(None),
    };
    let mut id = None;
    let mut author = None;
    let mut timestamp = None;
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Other(error.to_string()))?;
        let Some((prefix, local)) = key.split_once(':') else {
            continue;
        };
        if scope.get(prefix).map(String::as_str) != Some(WORD_NS) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
            .map_err(xml_error)?
            .into_owned();
        match local {
            "id" => id = value.parse().ok(),
            "author" => author = Some(value),
            "date" => timestamp = Some(value),
            _ => {}
        }
    }
    Ok(id.zip(author).map(|(id, author)| RevisionMetadata {
        kind,
        id,
        author,
        timestamp,
    }))
}

fn is_revision_local(local: &str) -> bool {
    matches!(
        local,
        "ins"
            | "del"
            | "moveFrom"
            | "moveTo"
            | "rPrChange"
            | "pPrChange"
            | "tblPrChange"
            | "sectPrChange"
    )
}

fn is_property_change(kind: RevisionKind) -> bool {
    matches!(
        kind,
        RevisionKind::RunPropertyChange
            | RevisionKind::ParagraphPropertyChange
            | RevisionKind::TablePropertyChange
            | RevisionKind::SectionPropertyChange
    )
}

fn keeps_content(resolution: Resolution, kind: RevisionKind) -> bool {
    !removes_content(resolution, kind)
}

fn removes_content(resolution: Resolution, kind: RevisionKind) -> bool {
    match resolution {
        Resolution::Accept => matches!(kind, RevisionKind::Deletion | RevisionKind::MoveFrom),
        Resolution::Reject => matches!(kind, RevisionKind::Insertion | RevisionKind::MoveTo),
    }
}

fn rename_element(raw: &[u8], qualified_name: &str, local: &str) -> Vec<u8> {
    let replacement = qualified_name
        .split_once(':')
        .map(|(prefix, _)| format!("{prefix}:{local}"))
        .unwrap_or_else(|| local.to_owned());
    let mut output = raw.to_vec();
    if let Some(position) = raw
        .windows(qualified_name.len())
        .position(|window| window == qualified_name.as_bytes())
    {
        output.splice(
            position..position + qualified_name.len(),
            replacement.as_bytes().iter().copied(),
        );
    }
    output
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::Other(format!("revision XML transformation failed: {error}"))
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Instant {
    seconds: i64,
    fraction: String,
}

impl Instant {
    fn parse(value: &str) -> Result<Self> {
        parse_rfc3339(value)
            .ok_or_else(|| Error::Other(format!("invalid RFC 3339 revision timestamp: {value}")))
    }
}

pub(crate) fn parse_rfc3339(value: &str) -> Option<Instant> {
    let separator = value
        .as_bytes()
        .iter()
        .position(|byte| matches!(byte, b'T' | b't'))?;
    if value.as_bytes()[separator + 1..]
        .iter()
        .any(|byte| matches!(byte, b'T' | b't'))
    {
        return None;
    }
    let date = &value[..separator];
    let time_and_zone = &value[separator + 1..];
    let date = date.as_bytes();
    if date.len() != 10
        || date[4] != b'-'
        || date[7] != b'-'
        || !date[..4].iter().all(u8::is_ascii_digit)
        || !date[5..7].iter().all(u8::is_ascii_digit)
        || !date[8..].iter().all(u8::is_ascii_digit)
    {
        return None;
    }
    let year: i64 = std::str::from_utf8(&date[..4]).ok()?.parse().ok()?;
    let month: u32 = std::str::from_utf8(&date[5..7]).ok()?.parse().ok()?;
    let day: u32 = std::str::from_utf8(&date[8..]).ok()?.parse().ok()?;
    if !(1..=12).contains(&month) || day == 0 || day > days_in_month(year, month) {
        return None;
    }

    let (time, offset_seconds) = if let Some(time) = time_and_zone
        .strip_suffix('Z')
        .or_else(|| time_and_zone.strip_suffix('z'))
    {
        (time, 0i64)
    } else {
        if time_and_zone.len() < 6 {
            return None;
        }
        let position = time_and_zone.len() - 6;
        let sign = if time_and_zone.as_bytes()[position] == b'+' {
            1i64
        } else if time_and_zone.as_bytes()[position] == b'-' {
            -1i64
        } else {
            return None;
        };
        let offset = &time_and_zone[position + 1..];
        if offset.len() != 5
            || offset.as_bytes()[2] != b':'
            || !offset[..2].bytes().all(|byte| byte.is_ascii_digit())
            || !offset[3..].bytes().all(|byte| byte.is_ascii_digit())
        {
            return None;
        }
        let (hours, minutes) = offset.split_once(':')?;
        let hours: i64 = hours.parse().ok()?;
        let minutes: i64 = minutes.parse().ok()?;
        if hours > 23 || minutes > 59 {
            return None;
        }
        (
            &time_and_zone[..position],
            sign * (hours * 3600 + minutes * 60),
        )
    };

    let (clock, fraction) = time
        .split_once('.')
        .map_or((time, ""), |(clock, fraction)| (clock, fraction));
    let clock = clock.as_bytes();
    if clock.len() != 8
        || clock[2] != b':'
        || clock[5] != b':'
        || !clock[..2].iter().all(u8::is_ascii_digit)
        || !clock[3..5].iter().all(u8::is_ascii_digit)
        || !clock[6..].iter().all(u8::is_ascii_digit)
        || (time.contains('.')
            && (fraction.is_empty() || !fraction.bytes().all(|byte| byte.is_ascii_digit())))
    {
        return None;
    }
    let hour: i64 = std::str::from_utf8(&clock[..2]).ok()?.parse().ok()?;
    let minute: i64 = std::str::from_utf8(&clock[3..5]).ok()?.parse().ok()?;
    let second: i64 = std::str::from_utf8(&clock[6..]).ok()?.parse().ok()?;
    if hour > 23 || minute > 59 || second > 59 {
        return None;
    }
    let days = days_from_civil(year, month, day);
    let seconds = days
        .checked_mul(86_400)?
        .checked_add(hour.checked_mul(3600)?)?
        .checked_add(minute.checked_mul(60)?)?
        .checked_add(second)?
        .checked_sub(offset_seconds)?;
    Some(Instant {
        seconds,
        fraction: fraction.trim_end_matches('0').to_owned(),
    })
}

impl Ord for Instant {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        self.seconds.cmp(&other.seconds).then_with(|| {
            let length = self.fraction.len().max(other.fraction.len());
            (0..length)
                .map(|index| {
                    self.fraction
                        .as_bytes()
                        .get(index)
                        .copied()
                        .unwrap_or(b'0')
                        .cmp(
                            &other
                                .fraction
                                .as_bytes()
                                .get(index)
                                .copied()
                                .unwrap_or(b'0'),
                        )
                })
                .find(|ordering| !ordering.is_eq())
                .unwrap_or(std::cmp::Ordering::Equal)
        })
    }
}

impl PartialOrd for Instant {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

fn days_in_month(year: i64, month: u32) -> u32 {
    match month {
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = year - i64::from(month <= 2);
    let era = if year >= 0 { year } else { year - 399 } / 400;
    let year_of_era = year - era * 400;
    let month = i64::from(month);
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

#[cfg(test)]
#[test]
fn carried_simple_fields_preserve_schema_and_selected_source_custody() {
    let document = |body: &str| {
        format!("<w:document xmlns:w=\"{WORD_NS}\"><w:body>{body}</w:body></w:document>")
    };
    let fixture = |kind: &str, plain: &str| {
        let text = if kind == "del" { "delText" } else { "t" };
        document(&format!(
            r#"<w:p><w:{kind} w:id="10" w:author="New"><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText><w:fldChar w:fldCharType="separate"/></w:r></w:{kind}><w:fldSimple w:instr="PAGEREF target" w:dirty="off" xmlns:x="urn:producer" x:note="kept"><w:fldData>QUJD</w:fldData><w:{kind} w:id="11" w:author="New"><w:r><w:{text}>1</w:{text}></w:r></w:{kind}></w:fldSimple>{plain}<w:{kind} w:id="12" w:author="New"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:{kind}></w:p>"#
        ))
    };
    for (kind, action) in [("del", Resolution::Accept), ("ins", Resolution::Reject)] {
        let source = fixture(kind, "");
        for id in [10, 12] {
            let error =
                resolve_story_xml(source.as_bytes(), action, RevisionScope::Id(id)).unwrap_err();
            assert!(
                error.to_string().contains("break complete field carriage"),
                "{error}"
            );
        }
        let boundary = HashSet::from([10, 12]);
        assert!(
            resolve_story_xml(source.as_bytes(), action, RevisionScope::Ids(&boundary)).is_err()
        );
        let (cache_only, count) =
            resolve_story_xml(source.as_bytes(), action, RevisionScope::Id(11)).unwrap();
        assert_eq!(count, 1);
        let cache_only = std::str::from_utf8(&cache_only).unwrap();
        assert!(cache_only.contains("fldSimple"));
        assert!(cache_only.contains("PAGEREF target"));
        assert!(cache_only.contains("x:note=\"kept\""));
        assert!(cache_only.contains("QUJD"));
        assert!(!cache_only.contains(">1<"));
        let (removed, count) =
            resolve_story_xml(source.as_bytes(), action, RevisionScope::Author("New")).unwrap();
        assert_eq!(count, 3);
        let removed = std::str::from_utf8(&removed).unwrap();
        assert!(!removed.contains("fldSimple"));
        assert!(!removed.contains("fldChar"));
        let plain = fixture(kind, "<w:r><w:t>unselected</w:t></w:r>");
        assert!(resolve_story_xml(plain.as_bytes(), action, RevisionScope::All).is_err());
    }
    let nested = document(
        r#"<w:p><w:del w:id="1" w:author="Outer"><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText><w:fldChar w:fldCharType="separate"/></w:r></w:del><w:del w:id="2" w:author="Inner"><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>QUOTE</w:instrText><w:fldChar w:fldCharType="separate"/></w:r></w:del><w:fldSimple w:instr="PAGE"><w:del w:id="3" w:author="Inner"><w:r><w:delText>1</w:delText></w:r></w:del></w:fldSimple><w:del w:id="4" w:author="Inner"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del><w:del w:id="5" w:author="Outer"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del></w:p>"#,
    );
    let (inner_removed, count) = resolve_story_xml(
        nested.as_bytes(),
        Resolution::Accept,
        RevisionScope::Author("Inner"),
    )
    .unwrap();
    assert_eq!(count, 3);
    let inner_removed = std::str::from_utf8(&inner_removed).unwrap();
    assert!(!inner_removed.contains("fldSimple"));
    assert!(inner_removed.contains("TOC"));
    assert!(
        resolve_story_xml(nested.as_bytes(), Resolution::Accept, RevisionScope::Id(2)).is_err()
    );
    let prior = fixture(
        "del",
        "<w:ins w:id=\"99\" w:author=\"Prior\"><w:r><w:t>prior</w:t></w:r></w:ins>",
    );
    assert!(
        resolve_story_xml(
            prior.as_bytes(),
            Resolution::Accept,
            RevisionScope::Author("New")
        )
        .is_err()
    );

    let complete = fixture("del", "");
    for attribute in [
        "",
        "w:fldCharType=\"unknown\"",
        "xmlns:x=\"urn:foreign\" x:fldCharType=\"separate\"",
    ] {
        let imported = complete.replace("w:fldCharType=\"separate\"", attribute);
        let tree = XmlTree::parse(imported.as_bytes()).unwrap();
        let (fields, _) = tree.carried_field_sources().unwrap();
        assert!(!fields[0].valid);
        let unrelated = imported.replace("</w:body>", "<w:p><w:ins w:id=\"80\" w:author=\"Other\"><w:r><w:t>outside</w:t></w:r></w:ins></w:p></w:body>");
        let (kept, count) = resolve_story_xml(
            unrelated.as_bytes(),
            Resolution::Accept,
            RevisionScope::Id(80),
        )
        .unwrap();
        assert_eq!(count, 1);
        let expected_body = imported
            .split("<w:body>")
            .nth(1)
            .unwrap()
            .split("</w:body>")
            .next()
            .unwrap();
        assert!(std::str::from_utf8(&kept).unwrap().contains(expected_body));
    }
    let foreign_marker = complete.replace(
        "<w:instrText>TOC</w:instrText>",
        "<w:instrText>TOC</w:instrText><x:fldChar xmlns:x=\"urn:foreign\" x:fldCharType=\"end\"/>",
    );
    assert!(
        XmlTree::parse(foreign_marker.as_bytes())
            .unwrap()
            .carried_field_sources()
            .unwrap()
            .0[0]
            .valid
    );
    let body = complete
        .split("<w:body>")
        .nth(1)
        .unwrap()
        .split("</w:body>")
        .next()
        .unwrap();
    for related in [
        format!("<w:hdr xmlns:w=\"{WORD_NS}\">{body}</w:hdr>"),
        format!(
            "<w:footnotes xmlns:w=\"{WORD_NS}\"><w:footnote w:id=\"2\">{body}</w:footnote></w:footnotes>"
        ),
    ] {
        let (removed, count) =
            resolve_story_xml(related.as_bytes(), Resolution::Accept, RevisionScope::All).unwrap();
        assert_eq!(count, 3);
        assert!(!std::str::from_utf8(&removed).unwrap().contains("fldSimple"));
    }
    let source = format!(
        r#"<w:p><w:r><w:t>before</w:t></w:r><q:fldSimple xmlns:q="{WORD_NS}" xmlns:w="urn:foreign" q:instr="PAGE" q:dirty="off" w:note="root"><q:fldData>QUJD</q:fldData><q:r><q:t w:note="leaf">1</q:t></q:r><q:fldSimple q:instr="DATE"><q:r><q:t>nested</q:t></q:r></q:fldSimple></q:fldSimple></w:p>"#
    );
    let mut id = 0;
    let tracked = tracked_paragraph_children(&source, true, false, &mut |content| {
        id += 1;
        Ok(format!(
            "<w:del xmlns:w=\"{WORD_NS}\" w:id=\"{id}\" w:author=\"New\">{content}</w:del>"
        ))
    })
    .unwrap();
    let tracked = document(&format!("<w:p>{tracked}</w:p>"));
    let tree = XmlTree::parse(tracked.as_bytes()).unwrap();
    for (index, element) in tree.elements.iter().enumerate() {
        if element.word && element.local == "fldSimple" {
            let parent = &tree.elements[element.parent.unwrap()];
            assert!(!matches!(
                parent.local.as_str(),
                "ins" | "del" | "moveFrom" | "moveTo"
            ));
        }
        if element.word
            && element.local == "delText"
            && tree.source[element.start..element.end]
                .windows(4)
                .any(|bytes| bytes == b"leaf")
        {
            assert!(
                tree.source_scope(index)
                    .iter()
                    .any(|(prefix, uri)| prefix == "w" && uri == "urn:foreign")
            );
        }
    }
    assert!(tracked.contains("q:dirty=\"off\""));
    assert!(tracked.contains("q:instr=\"PAGE\""));
    assert!(tracked.contains("<q:fldData>QUJD</q:fldData>"));
    for atomic in [
        r#"<q:r w:note="q:t > &lt;q:t&gt;"><q:t w:note="q:t">value<!-- marker > --><![CDATA[x > y]]></q:t><q:instrText>PAGE &amp; q:t</q:instrText><q:t/></q:r>"#,
        r#"<q:r w:note="empty"/>"#,
    ] {
        let source = format!(
            r#"<w:p><q:fldSimple xmlns:q="{WORD_NS}" xmlns:w="urn:producer" q:instr="PAGE">{atomic}</q:fldSimple></w:p>"#
        );
        let tracked = tracked_paragraph_children(&source, true, false, &mut |content| {
            Ok(format!(
                "<w:del xmlns:w=\"{WORD_NS}\" w:id=\"1\" w:author=\"New\">{content}</w:del>"
            ))
        })
        .unwrap();
        let tracked = document(&format!("<w:p>{tracked}</w:p>"));
        let tree = XmlTree::parse(tracked.as_bytes()).unwrap();
        let (index, run) = tree
            .elements
            .iter()
            .enumerate()
            .find(|(_, element)| element.word && element.local == "r")
            .unwrap();
        let opening = std::str::from_utf8(&tree.source[run.start..run.open_end]).unwrap();
        assert!(opening.contains("xmlns:w=\"urn:producer\""), "{tracked}");
        assert!(
            tree.source_scope(index)
                .iter()
                .any(|(prefix, uri)| { prefix == "w" && uri == "urn:producer" })
        );
        if !run.empty {
            assert_eq!(&tree.source[run.close_start..run.end], b"</q:r>");
            assert!(opening.contains(r#"w:note="q:t > &lt;q:t&gt;""#));
            assert!(tracked.contains(
                r#"<q:delText w:note="q:t">value<!-- marker > --><![CDATA[x > y]]></q:delText>"#
            ));
            assert!(tracked.contains("<q:delInstrText>PAGE &amp; q:t</q:delInstrText>"));
            assert!(tracked.contains("<q:delText/>"));
        } else {
            assert!(opening.contains("w:note=\"empty\""));
            assert!(opening.ends_with("/>"));
        }
    }
    let same_scope = tracked_paragraph_children(
        "<w:p><w:r><w:t>plain</w:t></w:r></w:p>",
        true,
        false,
        &mut |content| Ok(format!("<w:del>{content}</w:del>")),
    )
    .unwrap();
    assert_eq!(
        same_scope,
        "<w:del><w:r><w:delText>plain</w:delText></w:r></w:del>"
    );
    // A textbox re-entry recognized by the all-story scanner is still opaque
    // to the enclosing paragraph publisher's local inline ownership.
    let pict = r#"<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>opaque-text</w:t><w:instrText>opaque-code</w:instrText></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>"#;
    let source = format!("<w:p><w:r>{pict}<w:t>outer</w:t></w:r></w:p>");
    let tracked = tracked_paragraph_children(&source, true, false, &mut |content| {
        Ok(format!("<w:del>{content}</w:del>"))
    })
    .unwrap();
    assert!(tracked.contains(pict), "opaque host changed: {tracked}");
    assert!(
        tracked.contains("<w:delText>outer</w:delText>"),
        "{tracked}"
    );
    let boxed = document(&format!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText><w:fldChar w:fldCharType="separate"/>{}</w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>body</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        pict.replace(
            "<w:r><w:t>opaque-text</w:t><w:instrText>opaque-code</w:instrText></w:r>",
            "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>box</w:t></w:r></w:fldSimple>"
        )
    ));
    let boxed_tree = XmlTree::parse(boxed.as_bytes()).unwrap();
    let (_, owners) = boxed_tree.carried_field_sources().unwrap();
    let simple = owners
        .iter()
        .filter(|(index, _)| boxed_tree.elements[*index].local == "fldSimple")
        .collect::<Vec<_>>();
    assert_eq!(simple.len(), 2);
    assert!(simple[0].1.is_none(), "textbox borrowed body carrier");
    assert!(simple[1].1.is_some(), "direct body field lost carrier");
    let imported_boxed = CT_Document::from_xml(boxed.as_bytes()).unwrap();
    let boxed_paragraphs = imported_boxed.body.paragraphs().collect::<Vec<_>>();
    validate_carried_paragraph_fields(&boxed_paragraphs).unwrap();
    let atomic = document(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText><w:fldChar w:fldCharType="separate"/></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>body</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="end"/></w:r><w:sdt><w:sdtPr/><w:sdtContent><w:fldSimple w:instr="PAGE"><w:r><w:t>atomic</w:t></w:r></w:fldSimple></w:sdtContent></w:sdt></w:p>"#,
    );
    let imported_atomic = CT_Document::from_xml(atomic.as_bytes()).unwrap();
    let atomic_paragraphs = imported_atomic.body.paragraphs().collect::<Vec<_>>();
    validate_carried_paragraph_fields(&atomic_paragraphs).unwrap();
    validate_carried_paragraph_fields(&[]).unwrap();
    let standalone = document("<w:p><w:fldSimple w:instr=\"PAGE\"/></w:p>");
    let imported_standalone = CT_Document::from_xml(standalone.as_bytes()).unwrap();
    let standalone_paragraphs = imported_standalone.body.paragraphs().collect::<Vec<_>>();
    assert!(validate_carried_paragraph_fields(&standalone_paragraphs).is_err());
    let body = boxed_tree
        .direct_word_child(boxed_tree.root, "body")
        .unwrap();
    let paragraph = boxed_tree.direct_word_child(body, "p").unwrap();
    assert_eq!(
        boxed_tree.paragraph_field_events(paragraph),
        [FieldEvent::Begin, FieldEvent::Owner, FieldEvent::End]
    );
    for (hidden, expected) in [
        (
            r#"<x:wrapper><w:fldChar w:fldCharType="end"/></x:wrapper><x:fldChar w:fldCharType="end"/><w:fldChar x:fldCharType="end"/>"#,
            vec![FieldEvent::Begin, FieldEvent::Owner, FieldEvent::End],
        ),
        (
            r#"<w:fldChar w:fldCharType="end"/>"#,
            vec![
                FieldEvent::Begin,
                FieldEvent::End,
                FieldEvent::Owner,
                FieldEvent::End,
            ],
        ),
    ] {
        let source = document(&format!(
            r#"<w:p xmlns:q="{WORD_NS}" xmlns:x="urn:foreign"><q:r><q:fldChar q:fldCharType="b&#101;gin"/><q:fldChar q:fldCharType="separate"/>{hidden}</q:r><x:fldSimple/><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><q:r><q:fldChar q:fldCharType="end"/></q:r></w:p>"#
        ));
        let tree = XmlTree::parse(source.as_bytes()).unwrap();
        let body = tree.direct_word_child(tree.root, "body").unwrap();
        let paragraph = tree.direct_word_child(body, "p").unwrap();
        assert_eq!(tree.paragraph_field_events(paragraph), expected);
    }

    let simple = "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>";
    for depth in [16, 32] {
        let result = format!(
            "{}{}{}",
            "<w:fldSimple w:instr=\"QUOTE\">".repeat(depth),
            simple,
            "</w:fldSimple>".repeat(depth)
        );
        let source = document(&format!("<w:p>{result}</w:p>"));
        CT_Document::from_xml(source.as_bytes()).unwrap();
        let tracked = tracked_paragraph_children(
            &format!("<w:p>{result}</w:p>"),
            false,
            false,
            &mut |content| Ok(format!("<w:ins>{content}</w:ins>")),
        )
        .unwrap();
        assert_eq!(tracked.matches("<w:fldSimple ").count(), depth + 1);
        assert!(!tracked.contains("<w:ins><w:fldSimple"));
    }
    let too_deep = document(&format!(
        "<w:p>{}<w:r><w:t>x</w:t></w:r>{}</w:p>",
        "<w:sdt><w:sdtContent>".repeat(257),
        "</w:sdtContent></w:sdt>".repeat(257)
    ));
    assert!(CT_Document::from_xml(too_deep.as_bytes()).is_err());
    let nested_link = "<w:p><w:hyperlink w:anchor=\"a\"><w:fldSimple w:instr=\"PAGE\"><w:hyperlink w:anchor=\"b\"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:fldSimple></w:hyperlink></w:p>";
    assert!(
        tracked_paragraph_children(nested_link, false, false, &mut |content| Ok(
            content.to_owned()
        ))
        .is_err()
    );
}

#[cfg(test)]
fn issue434_snapshot_texts(body: &str) -> Vec<(RevisionKind, String)> {
    let source = format!(
        r#"<w:document xmlns:w="{WORD_NS}" xmlns:a="{WORD_NS}" xmlns:x="urn:foreign"><w:body>{body}</w:body></w:document>"#
    );
    let tree = XmlTree::parse(source.as_bytes()).unwrap();
    let facts = tree.snapshot_facts();
    tree.elements
        .iter()
        .enumerate()
        .filter_map(|(index, element)| {
            element
                .revision
                .as_ref()
                .map(|metadata| (metadata.kind, tree.snapshot_text(index, &facts).unwrap()))
        })
        .collect()
}

#[test]
fn issue434_snapshot_scalars_current_properties_and_opaque_paths() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:pPr><w:pPrChange w:id="1" w:author="A"><w:pPr><w:r><w:t>old poison</w:t></w:r></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="2" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t> A &amp; B&#x43; <![CDATA[<raw>]]>&unknown;</w:t><w:tab/><w:br/><w:cr/><w:drawing><w:txbxContent><w:p><w:r><w:t>drawing poison</w:t></w:r></w:p></w:txbxContent></w:drawing><x:t>foreign poison</x:t><w:sym w:char="0041"/></w:r><w:del w:id="3" w:author="A"><w:r><w:delText>removed</w:delText></w:r></w:del><w:ins w:id="4" w:author="A"><a:r><a:t>new</a:t></a:r></w:ins></w:p>"#,
    );
    let scalar = " A & BC <raw>&unknown;\t\n\n";
    assert_eq!(
        revisions,
        vec![
            (
                RevisionKind::ParagraphPropertyChange,
                format!("{scalar}new")
            ),
            (RevisionKind::RunPropertyChange, scalar.into()),
            (RevisionKind::Deletion, "removed".into()),
            (RevisionKind::Insertion, "new".into()),
        ]
    );
}

#[test]
fn issue434_snapshot_move_mark_section_and_empty_kinds() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="A"/></w:rPr><w:pPrChange w:id="2" w:author="A"><w:pPr/></w:pPrChange><w:sectPr><w:sectPrChange w:id="3" w:author="A"><w:sectPr/></w:sectPrChange></w:sectPr></w:pPr><w:r><w:t>retained</w:t></w:r><w:moveFrom w:id="4" w:author="A"><w:r><w:t>from</w:t></w:r></w:moveFrom><w:moveTo w:id="5" w:author="A"><w:r><w:t>to</w:t></w:r></w:moveTo><w:ins w:id="6" w:author="A"/></w:p>"#,
    );
    assert_eq!(
        revisions,
        vec![
            (RevisionKind::Deletion, String::new()),
            (RevisionKind::ParagraphPropertyChange, "retainedto".into()),
            (RevisionKind::SectionPropertyChange, String::new()),
            (RevisionKind::MoveFrom, "from".into()),
            (RevisionKind::MoveTo, "to".into()),
            (RevisionKind::Insertion, String::new()),
        ]
    );
}

#[test]
fn issue434_snapshot_table_separators_nested_controls_and_removed_rows() {
    let revisions = issue434_snapshot_texts(
        r#"<w:tbl><w:tblPr><w:tblPrChange w:id="1" w:author="A"><w:tblPr/></w:tblPrChange></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>B</w:t></w:r></w:p></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:del w:id="2" w:author="A"/></w:trPr><w:tc><w:p><w:r><w:t>removed row</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:moveFrom w:id="3" w:author="A"/></w:trPr><w:tc><w:p><w:r><w:t>moved row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    assert_eq!(
        revisions[0],
        (RevisionKind::TablePropertyChange, "A\tB\tC\tD\t\n".into())
    );
    assert_eq!(revisions[1].1, "");
    assert_eq!(revisions[2].1, "");
}

#[test]
fn issue434_snapshot_stored_field_caches_and_wrapper_frontiers() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:pPr><w:pPrChange w:id="1" w:author="A"><w:pPr/></w:pPrChange></w:pPr><w:smartTag><w:r><w:t>smart</w:t></w:r><w:customXml><w:r><w:t>custom</w:t></w:r></w:customXml></w:smartTag><w:fldSimple w:instr="instruction poison"><w:r><w:t>simple</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>outer poison</w:instrText><w:t>instruction text poison</w:t></w:r><w:r><w:fldChar w:fldCharType="separate"/><w:t>outer</w:t><w:fldChar w:fldCharType="begin"/><w:instrText>inner poison</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>inner</w:t><w:fldChar w:fldCharType="end"/><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="begin"/><w:t>no cache poison</w:t><w:fldChar w:fldCharType="end"/></w:r><w:hyperlink><w:smartTag><w:r><w:t>opaque nested wrapper</w:t></w:r></w:smartTag><w:r><w:t>link</w:t></w:r></w:hyperlink><w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r><w:customXml><w:r><w:t>opaque control wrapper</w:t></w:r></w:customXml></w:sdtContent></w:sdt><x:customXml><w:r><w:t>foreign</w:t></w:r></x:customXml><x:opaque><w:r><w:t>opaque</w:t></w:r></x:opaque></w:p>"#,
    );
    assert_eq!(revisions[0].1, "smartcustomsimpleouterinnerlinkcontrol");
}

#[test]
fn issue434_snapshot_raw_wrapper_requires_base_canonical_writer_binding() {
    let source = format!(
        r#"<a:document xmlns:a="{WORD_NS}"><a:body><a:p><a:pPr><a:pPrChange a:id="1" a:author="A"><a:pPr/></a:pPrChange></a:pPr><a:r><a:t>direct</a:t></a:r><a:smartTag><a:r><a:t>excluded</a:t></a:r></a:smartTag><a:customXml xmlns:w="{WORD_NS}"><a:r><a:t>included</a:t></a:r></a:customXml></a:p></a:body></a:document>"#
    );
    let tree = XmlTree::parse(source.as_bytes()).unwrap();
    let facts = tree.snapshot_facts();
    let index = tree
        .elements
        .iter()
        .position(|element| element.revision.is_some())
        .unwrap();
    assert_eq!(tree.snapshot_text(index, &facts).unwrap(), "directincluded");
}

#[test]
fn issue434_snapshot_field_context_precedes_selected_owner() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>CODE</w:instrText></w:r></w:p><w:p><w:r><w:rPr><w:rPrChange w:id="1" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>instruction poison</w:t></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rPrChange w:id="2" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>cache</w:t></w:r><w:ins w:id="3" w:author="A"><w:r><w:t>inserted cache</w:t></w:r></w:ins><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    assert_eq!(
        revisions
            .iter()
            .map(|(_, text)| text.as_str())
            .collect::<Vec<_>>(),
        ["", "cache", "inserted cache"]
    );
}

#[test]
fn issue434_snapshot_nested_field_result_cannot_escape_outer_instruction() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:pPr><w:pPrChange w:id="1" w:author="A"><w:pPr/></w:pPrChange></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:sdt><w:sdtContent><w:r><w:fldChar w:fldCharType="end"/><w:t>bad domain poison</w:t></w:r></w:sdtContent></w:sdt><w:r><w:fldChar w:fldCharType="begin"/><w:fldChar w:fldCharType="separate"/><w:t>inner instruction poison</w:t><w:fldChar w:fldCharType="end"/><w:fldChar w:fldCharType="separate"/><w:t>outer cache</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    assert_eq!(revisions[0].1, "outer cache");
}

#[test]
fn issue434_snapshot_property_owner_in_removed_ancestor_is_empty() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:del w:id="1" w:author="A"><w:r><w:rPr><w:rPrChange w:id="2" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:delText>removed</w:delText></w:r></w:del></w:p><w:tbl><w:tr><w:trPr><w:del w:id="3" w:author="A"/></w:trPr><w:tc><w:p><w:pPr><w:pPrChange w:id="4" w:author="A"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>removed row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    assert_eq!(
        revisions
            .iter()
            .map(|(_, text)| text.as_str())
            .collect::<Vec<_>>(),
        ["removed", "", "", ""]
    );
}

#[test]
fn issue434_snapshot_story_roots_share_owned_projection_and_order() {
    use crate::StoryKind;
    let body = r#"<w:p><w:del w:id="1" w:author="A"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="2" w:author="A"><w:r><w:t>new</w:t></w:r></w:ins></w:p>"#;
    for (kind, opening, closing) in [
        (
            StoryKind::Body,
            "<w:document><w:body>",
            "</w:body></w:document>",
        ),
        (StoryKind::Header, "<w:hdr>", "</w:hdr>"),
        (StoryKind::Footer, "<w:ftr>", "</w:ftr>"),
        (
            StoryKind::Footnote,
            "<w:footnotes><w:footnote w:id=\"1\">",
            "</w:footnote></w:footnotes>",
        ),
        (
            StoryKind::Endnote,
            "<w:endnotes><w:endnote w:id=\"1\">",
            "</w:endnote></w:endnotes>",
        ),
        (
            StoryKind::Comment,
            "<w:comments><w:comment w:id=\"1\">",
            "</w:comment></w:comments>",
        ),
        (StoryKind::TextBox, "<w:txbxContent>", "</w:txbxContent>"),
    ] {
        let raw = format!("{opening}{body}{closing}").replacen(
            '>',
            &format!(" xmlns:w=\"{WORD_NS}\">"),
            1,
        );
        let story = StoryId::new(kind, "/word/test.xml", 0);
        let snapshots = story_part_revisions(
            "/word/test.xml",
            raw.as_bytes(),
            &[(0..raw.len(), story.clone())],
        )
        .unwrap();
        assert_eq!(
            snapshots
                .iter()
                .map(|revision| revision.text())
                .collect::<Vec<_>>(),
            ["old", "new"],
            "{kind:?}: {raw}"
        );
        assert!(snapshots.iter().all(|revision| revision.story() == &story));
    }
}

#[test]
fn issue434_snapshot_private_package_revision_has_empty_text() {
    use base64::Engine;
    let mut document = Document::new();
    let record = br#"{"id":7,"author":"A","timestamp":"2026-11-03T00:00:00Z"}"#;
    let encoded = base64::engine::general_purpose::STANDARD.encode(record);
    document.package.set_part(
        crate::comparison::COMMENT_COMPARISON_PART,
        format!(
            "{}{}{}",
            crate::comparison::COMMENT_COMPARISON_OPEN,
            encoded,
            crate::comparison::COMMENT_COMPARISON_CLOSE
        )
        .into_bytes(),
    );
    let revision = comment_comparison_revision(&document, StoryId::body())
        .unwrap()
        .unwrap();
    assert_eq!(revision.id(), 7);
    assert_eq!(revision.text(), "");
}

#[test]
fn issue434_snapshot_controls_keep_destination_grammar() {
    let revisions = issue434_snapshot_texts(
        r#"<w:tbl><w:tblPr><w:tblPrChange w:id="1" w:author="A"><w:tblPr/></w:tblPrChange></w:tblPr><w:sdt><w:sdtContent><w:p><w:r><w:t>wrong table-control child</w:t></w:r></w:p><w:tr><w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt></w:tbl><w:p><w:pPr><w:pPrChange w:id="2" w:author="A"><w:pPr/></w:pPrChange></w:pPr><w:sdt><w:sdtContent><w:p><w:r><w:t>wrong inline child</w:t></w:r></w:p><w:r><w:t>inline</w:t></w:r></w:sdtContent></w:sdt></w:p>"#,
    );
    assert_eq!(revisions[0].1, "cell\t\n");
    assert_eq!(revisions[1].1, "inline");
}

#[test]
fn issue434_snapshot_nested_revision_local_payload_is_not_ancestor_hidden() {
    let revisions = issue434_snapshot_texts(
        r#"<w:p><w:del w:id="1" w:author="A"><w:r><w:delText>outer</w:delText></w:r><w:del w:id="2" w:author="A"><w:r><w:delText>inner</w:delText></w:r></w:del></w:del></w:p>"#,
    );
    assert_eq!(
        revisions
            .iter()
            .map(|(_, text)| text.as_str())
            .collect::<Vec<_>>(),
        ["outerinner", "inner"]
    );
}

#[test]
fn issue434_misplaced_properties_and_paragraph_mark_have_no_run_owner() {
    let texts = issue434_snapshot_texts(
        r#"<w:p><w:pPr><w:rPr><w:rPrChange w:id="4" w:author="A"><w:rPr/></w:rPrChange></w:rPr></w:pPr><w:sdt><w:sdtContent><w:rPrChange w:id="1" w:author="A"><w:rPr/></w:rPrChange><w:pPrChange w:id="2" w:author="A"><w:pPr/></w:pPrChange><w:tblPrChange w:id="3" w:author="A"><w:tblPr/></w:tblPrChange><w:r><w:t>unrelated control text</w:t></w:r></w:sdtContent></w:sdt></w:p>"#,
    );
    assert_eq!(texts.len(), 4);
    assert!(texts.iter().all(|(_, text)| text.is_empty()));
}

#[test]
fn issue434_disjoint_revision_projection_prepares_context_once() {
    for n in [1000, 2000] {
        let body = (0..n)
            .map(|id| {
                format!(
                    r#"<w:p><w:ins w:id="{id}" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#
                )
            })
            .collect::<String>();
        let source =
            format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body>{body}</w:body></w:document>"#);
        let tree = XmlTree::parse(source.as_bytes()).unwrap();
        let facts = tree.snapshot_facts();
        let texts = tree
            .elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                element
                    .revision
                    .as_ref()
                    .map(|_| tree.snapshot_text(index, &facts).unwrap())
            })
            .collect::<Vec<_>>();
        assert_eq!(texts.len(), n);
        assert!(texts.iter().all(|text| text == "x"));
        assert_eq!(facts.preparation_visits, 2 * tree.elements.len());
        assert_eq!(
            facts.owner_visits.get(),
            5 * n,
            "owner walks never replay earlier paragraphs"
        );
        assert!(facts.fields.is_empty());
        let story = StoryId::body();
        let snapshots = story_part_revisions(
            "/word/document.xml",
            source.as_bytes(),
            &[(0..source.len(), story)],
        )
        .unwrap();
        assert_eq!(
            snapshots
                .iter()
                .map(StoryRevision::text)
                .collect::<Vec<_>>(),
            texts
        );
    }
}

#[test]
fn issue434_field_checkpoint_deletion_forks_preserve_current_and_local_modes() {
    let texts = issue434_snapshot_texts(
        r#"<w:p><w:del w:id="1" w:author="A"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:ins w:id="2" w:author="A"><w:r><w:t>inner instruction poison</w:t></w:r></w:ins></w:del><w:ins w:id="3" w:author="A"><w:r><w:t>outside visible</w:t></w:r></w:ins><w:ins w:id="4" w:author="A"><w:del w:id="5" w:author="A"><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:del><w:r><w:t>outer local instruction poison</w:t></w:r></w:ins><w:ins w:id="6" w:author="A"><w:r><w:t>later visible</w:t></w:r></w:ins></w:p>"#,
    );
    assert_eq!(
        texts
            .iter()
            .map(|(_, text)| text.as_str())
            .collect::<Vec<_>>(),
        ["", "", "outside visible", "", "", "later visible"]
    );
}

#[test]
fn issue430_note_publication_removal_transport_preserves_ranges_and_classifies_drop() {
    let source = b"AAfirstBBsecondCC";
    let mut trace = NotePublication::new("/word/footnotes.xml", 7, &[2..7, 9..15]);
    trace.source_hash = Some(Sha256::digest(source).into());
    trace
        .remove_ranges("/word/footnotes.xml", source, std::slice::from_ref(&(0..2)))
        .unwrap();
    assert_eq!(trace.roots, vec![Some(0..5), Some(7..13)]);
    trace
        .remove_ranges(
            "/word/footnotes.xml",
            b"firstBBsecondCC",
            std::slice::from_ref(&(0..5)),
        )
        .unwrap();
    assert_eq!(trace.roots, vec![None, Some(2..8)]);
    trace.validate_source(b"BBsecondCC").unwrap();
    let roots = trace.roots.clone();
    let hash = trace.source_hash;
    assert!(
        trace
            .remove_ranges(
                "/word/footnotes.xml",
                b"BBsecondCC",
                std::slice::from_ref(&(3..4))
            )
            .is_err()
    );
    assert_eq!(trace.roots, roots);
    assert_eq!(trace.source_hash, hash);
    assert!(
        trace
            .remove_ranges(
                "/word/footnotes.xml",
                b"XXsecondCC",
                std::slice::from_ref(&(0..2))
            )
            .is_err()
    );
    assert_eq!(trace.roots, roots);
    assert_eq!(trace.source_hash, hash);
    assert!(
        trace
            .remove_ranges("/word/footnotes.xml", b"BBsecondCC", &[0..2, 1..2])
            .is_err()
    );
    assert_eq!(trace.roots, roots);
    assert_eq!(trace.source_hash, hash);
}
