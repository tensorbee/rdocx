//! Public comment handles and atomic document comment mutations.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;

use oxml_opc::OpcPackage;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::comments::{CT_Comment, CT_Comments};
use rdocx_oxml::comments_extended::{CT_CommentEx, CT_CommentsEx};
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::BodyContent;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, RangeAnchor};
#[cfg(test)]
use rdocx_oxml::text::{CT_R, CommentRangeMarker, HyperlinkSpan, RunContent};

use crate::document::visit_body_paragraphs_mut;
use crate::{ContentLocation, Document, Error, Result};

/// Refuses comment fields holding a character XML 1.0 cannot carry, as
/// python-docx refuses such text.
fn reject_non_xml_comment(author: &str, initials: Option<&str>, text: &str) -> Result<()> {
    oxml_core::xml::reject_non_xml_characters("comment author", author)?;
    oxml_core::xml::reject_non_xml_characters("comment initials", initials.unwrap_or_default())?;
    oxml_core::xml::reject_non_xml_characters("comment text", text)?;
    Ok(())
}

pub(crate) const COMMENTS_EXTENDED_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
pub(crate) const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
pub(crate) const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const DEFAULT_COMMENTS_PART: &str = "/word/comments.xml";
const DEFAULT_COMMENTS_EXTENDED_PART: &str = "/word/commentsExtended.xml";

/// A stable insertion point between runs in a body paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct RunPosition {
    /// Direct body child index, where a table or a block content control
    /// counts as one child. `Document::find_content_index` returns it.
    pub body_index: usize,
    /// Run boundary in the selected paragraph, counted over the runs that
    /// `Paragraph::runs` lists, including the runs inside inline content
    /// controls and tracked insertions.
    pub run_index: usize,
}

/// A half-open document run range, inclusive at `start` and exclusive at `end`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RunRange {
    pub start: RunPosition,
    pub end: RunPosition,
}

/// A stable insertion point between runs in a checked story paragraph.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct StoryRunPosition {
    pub location: ContentLocation,
    pub run_index: usize,
}

/// A half-open run range within one checked story owner.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct StoryRunRange {
    pub start: StoryRunPosition,
    pub end: StoryRunPosition,
}

/// The four paired marker families supported by checked story ranges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StoryRangeKind {
    Bookmark {
        id: i32,
        name: String,
    },
    Comment {
        id: i32,
    },
    Permission {
        id: i32,
        editor: Option<String>,
        group: Option<String>,
    },
    Proofing {
        kind: String,
    },
}

/// One immutable pair of accepted-view story positions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryRangeRef {
    kind: StoryRangeKind,
    range: StoryRunRange,
    pub(crate) start_ordinal: usize,
    pub(crate) end_ordinal: usize,
}

impl StoryRangeRef {
    pub fn kind(&self) -> &StoryRangeKind {
        &self.kind
    }

    pub fn range(&self) -> &StoryRunRange {
        &self.range
    }

    pub fn bookmark_id(&self) -> Option<i32> {
        match self.kind {
            StoryRangeKind::Bookmark { id, .. } => Some(id),
            _ => None,
        }
    }
}

#[derive(Debug, Clone)]
struct StoryMarker {
    kind: StoryRangeKind,
    start: bool,
    run_index: usize,
    span: Range<usize>,
    location: ContentLocation,
    ordinal: usize,
}

fn word_marker_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == name
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
        {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| Error::Other(error.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn scan_story_markers(xml: &[u8], location: &ContentLocation) -> Result<Vec<StoryMarker>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<(Vec<u8>, bool, Option<usize>)> = Vec::new();
    let mut markers = Vec::<StoryMarker>::new();
    let mut run_index = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("story marker scan failed: {error}")))?;
        let is_word = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes());
        drop(namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let hidden = stack.iter().any(|(name, word, _)| {
                    *word
                        && matches!(
                            name.as_slice(),
                            b"del"
                                | b"moveFrom"
                                | b"txbxContent"
                                | b"smartTag"
                                | b"customXml"
                                | b"fldSimple"
                        )
                });
                if is_word && local == b"r" && !hidden {
                    run_index += 1;
                }
                let marker = if is_word && !hidden {
                    let id = word_marker_attribute(&reader, element, b"id")?;
                    match local.as_slice() {
                        b"bookmarkStart" | b"bookmarkEnd" => {
                            let id = id.and_then(|id| id.parse().ok()).ok_or_else(|| {
                                Error::Other("bookmark marker has no valid id".to_owned())
                            })?;
                            let name = word_marker_attribute(&reader, element, b"name")?
                                .unwrap_or_default();
                            Some((
                                StoryRangeKind::Bookmark { id, name },
                                local == b"bookmarkStart",
                            ))
                        }
                        b"commentRangeStart" | b"commentRangeEnd" => {
                            let id = id.and_then(|id| id.parse().ok()).ok_or_else(|| {
                                Error::Other("comment marker has no valid id".to_owned())
                            })?;
                            Some((
                                StoryRangeKind::Comment { id },
                                local == b"commentRangeStart",
                            ))
                        }
                        b"permStart" | b"permEnd" => {
                            let id = id.and_then(|id| id.parse().ok()).ok_or_else(|| {
                                Error::Other("permission marker has no valid id".to_owned())
                            })?;
                            let editor = word_marker_attribute(&reader, element, b"ed")?;
                            let group = word_marker_attribute(&reader, element, b"edGrp")?;
                            Some((
                                StoryRangeKind::Permission { id, editor, group },
                                local == b"permStart",
                            ))
                        }
                        b"proofErr" => {
                            let kind = word_marker_attribute(&reader, element, b"type")?;
                            match kind.as_deref() {
                                Some("spellStart") => Some((
                                    StoryRangeKind::Proofing {
                                        kind: "spell".to_owned(),
                                    },
                                    true,
                                )),
                                Some("spellEnd") => Some((
                                    StoryRangeKind::Proofing {
                                        kind: "spell".to_owned(),
                                    },
                                    false,
                                )),
                                Some("gramStart") => Some((
                                    StoryRangeKind::Proofing {
                                        kind: "gram".to_owned(),
                                    },
                                    true,
                                )),
                                Some("gramEnd") => Some((
                                    StoryRangeKind::Proofing {
                                        kind: "gram".to_owned(),
                                    },
                                    false,
                                )),
                                _ => None,
                            }
                        }
                        _ => None,
                    }
                } else {
                    None
                };
                let marker_index = marker.map(|(kind, start)| {
                    markers.push(StoryMarker {
                        kind,
                        start,
                        run_index,
                        span: before..after,
                        location: location.clone(),
                        ordinal: 0,
                    });
                    markers.len() - 1
                });
                if matches!(event, Event::Start(_)) {
                    stack.push((local, is_word, marker_index));
                }
            }
            Event::End(_) => {
                if let Some((_, _, Some(index))) = stack.pop() {
                    markers[index].span.end = after;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(markers)
}

fn same_marker_identity(start: &StoryRangeKind, end: &StoryRangeKind) -> bool {
    match (start, end) {
        (StoryRangeKind::Bookmark { id: left, .. }, StoryRangeKind::Bookmark { id: right, .. })
        | (StoryRangeKind::Comment { id: left }, StoryRangeKind::Comment { id: right })
        | (
            StoryRangeKind::Permission { id: left, .. },
            StoryRangeKind::Permission { id: right, .. },
        ) => left == right,
        (StoryRangeKind::Proofing { kind: left }, StoryRangeKind::Proofing { kind: right }) => {
            left == right
        }
        _ => false,
    }
}

/// Immutable summary of one correlated bookmark or one reported marker issue.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BookmarkRef {
    id: Option<i32>,
    name: Option<String>,
    range: Option<RunRange>,
    direct_range: Option<RunRange>,
    text: String,
    issue: Option<String>,
}

impl BookmarkRef {
    pub fn id(&self) -> Option<i32> {
        self.id
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return the accepted-view half-open range reported by `Document::bookmarks`.
    ///
    /// Its body index is the paragraph ordinal counted recursively through
    /// tables and block content controls, which field numbering reads.
    /// [`Self::direct_range`] reports the direct body child index instead.
    pub fn range(&self) -> Option<RunRange> {
        self.range
    }

    /// Return the same range with the direct body child index that
    /// `RunPosition`, `Document::add_bookmark` and
    /// `Document::find_content_index` use.
    ///
    /// It is `None` when either marker sits in a table cell or a block content
    /// control, which have no direct body index. Run indexes are the same
    /// accepted-view boundaries as [`Self::range`], which are the
    /// `RunPosition` run indexes that `Document::add_bookmark` and
    /// `Document::add_comment` take.
    pub fn direct_range(&self) -> Option<RunRange> {
        self.direct_range
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn issue(&self) -> Option<&str> {
        self.issue.as_deref()
    }
}

/// Read-only view of a comment and its thread metadata.
#[derive(Clone, Copy)]
pub struct CommentRef<'a> {
    document: &'a Document,
    inner: &'a CT_Comment,
    extension: Option<&'a CT_CommentEx>,
    parent_id: Option<i32>,
}

impl std::fmt::Debug for CommentRef<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CommentRef")
            .field("inner", self.inner)
            .field("extension", &self.extension)
            .field("parent_id", &self.parent_id)
            .finish()
    }
}

impl CommentRef<'_> {
    /// Checked accepted-view range, without an invented parent range for replies.
    pub fn anchor(&self) -> Result<Option<StoryRunRange>> {
        self.document.comment_anchor(self.id())
    }

    /// Checked accepted span text. A point is empty and a known orphan is absent.
    pub fn anchor_text(&self) -> Result<Option<String>> {
        self.document.comment_anchor_text(self.id())
    }

    pub fn id(&self) -> i32 {
        self.inner.id
    }

    pub fn author(&self) -> Option<&str> {
        self.inner.author.as_deref()
    }

    pub fn initials(&self) -> Option<&str> {
        self.inner.initials.as_deref()
    }

    pub fn date(&self) -> Option<&str> {
        self.inner.date.as_deref()
    }

    pub fn text(&self) -> String {
        self.inner
            .paragraphs
            .iter()
            .map(CT_P::text)
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub fn parent_id(&self) -> Option<i32> {
        self.parent_id
    }

    pub fn resolved(&self) -> bool {
        self.extension.and_then(|entry| entry.done).unwrap_or(false)
    }
}

impl Document {
    /// Add a bookmark across one checked story owner.
    pub fn add_story_bookmark(&mut self, name: &str, range: StoryRunRange) -> Result<i32> {
        validate_bookmark_name(name)?;
        if self.story_ranges()?.iter().any(|entry| {
            matches!(&entry.kind, StoryRangeKind::Bookmark { name: existing, .. } if existing == name)
        }) {
            return Err(Error::Other(format!("bookmark name {name} already exists")));
        }
        let mut candidate = self.clone_for_staging();
        let mut identifiers = candidate.identifiers.clone();
        let id = identifiers.reserve_bookmark_id()?;
        candidate.anchor_story_range(
            &range,
            RangeAnchor::Bookmark { id, name },
            "bookmark",
            false,
        )?;
        candidate.identifiers = identifiers;
        candidate.story_ranges()?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(id)
    }

    /// Add a permission range with exactly one editor or editor group.
    pub fn add_story_permission_range(
        &mut self,
        editor: Option<&str>,
        group: Option<&str>,
        range: StoryRunRange,
    ) -> Result<i32> {
        if editor.is_some() == group.is_some()
            || editor.is_some_and(str::is_empty)
            || group.is_some_and(str::is_empty)
        {
            return Err(Error::Other(
                "permission range needs one editor or group".to_owned(),
            ));
        }
        oxml_core::xml::reject_non_xml_characters("permission editor", editor.unwrap_or_default())?;
        oxml_core::xml::reject_non_xml_characters("permission group", group.unwrap_or_default())?;
        let occupied = self
            .story_ranges()?
            .into_iter()
            .filter_map(|entry| match entry.kind {
                StoryRangeKind::Permission { id, .. } => Some(id),
                _ => None,
            })
            .collect::<HashSet<_>>();
        let id = (0..=i32::MAX)
            .find(|candidate| !occupied.contains(candidate))
            .ok_or_else(|| Error::Other("permission range identifiers are exhausted".to_owned()))?;
        let mut candidate = self.clone_for_staging();
        candidate.anchor_story_range(
            &range,
            RangeAnchor::Permission { id, editor, group },
            "permission",
            false,
        )?;
        candidate.story_ranges()?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(id)
    }

    /// Add a spelling (`spell`) or grammar (`gram`) proofing range.
    pub fn add_story_proofing_range(&mut self, kind: &str, range: StoryRunRange) -> Result<()> {
        if !matches!(kind, "spell" | "gram") {
            return Err(Error::Other(
                "proofing kind must be spell or gram".to_owned(),
            ));
        }
        self.story_ranges()?;
        let mut candidate = self.clone_for_staging();
        candidate.anchor_story_range(&range, RangeAnchor::Proofing { kind }, "proofing", false)?;
        candidate.story_ranges()?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(())
    }

    /// Remove one checked bookmark pair by its identifier.
    pub fn remove_story_bookmark(&mut self, id: i32) -> Result<bool> {
        let entry = self.story_ranges()?.into_iter().find(|entry| {
            matches!(entry.kind, StoryRangeKind::Bookmark { id: existing, .. } if existing == id)
        });
        let Some(entry) = entry else {
            return Ok(false);
        };
        let mut candidate = self.clone_for_staging();
        candidate.remove_story_range_markers(&entry, false)?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(true)
    }

    /// Move a checked pair to a new range in one story owner.
    /// Move one existing root comment without changing its thread identity.
    /// Unknown IDs, replies, incomplete anchors and unsupported destinations
    /// are refused without publishing any change.
    pub fn move_comment(&mut self, id: i32, range: StoryRunRange) -> Result<()> {
        let selected = self.checked_movable_comment(id)?;
        self.move_story_range(&selected, range)
    }

    /// Move an existing root comment onto a literal main-story occurrence.
    /// Search and run splitting follow [`Self::add_comment_on_text`].
    pub fn move_comment_to_text(&mut self, id: i32, anchor: &str, occurrence: usize) -> Result<()> {
        let selected = self.checked_movable_comment(id)?;
        let mut candidate = self.clone_for_staging();
        let reference = candidate.remove_comment_source_for_move(&selected)?;
        candidate.anchor_existing_comment_on_text(id, anchor, occurrence)?;
        candidate.restore_comment_reference_run(id, reference)?;
        candidate.flush_to_package()?;
        candidate.comment_ownership_at(&candidate.doc_part_name)?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(())
    }

    fn checked_movable_comment(&self, id: i32) -> Result<StoryRangeRef> {
        let mut proof = self.clone_for_staging();
        proof.flush_to_package()?;
        let ownership = proof.comment_ownership_at(&proof.doc_part_name)?;
        match ownership.parents.get(&id) {
            None => return Err(Error::Other(format!("unknown comment {id}"))),
            Some(Some(_)) => {
                return Err(Error::Other(format!(
                    "comment {id} is a reply, not a movable root"
                )));
            }
            Some(None) => {}
        }
        let counts = CommentOwnership::marker_counts(&ownership.markers);
        if counts.get(&id) != Some(&[1, 1, 1]) {
            return Err(Error::Other(format!(
                "comment {id} requires one complete paired range and reference"
            )));
        }
        // This exact accepted projection also rejects hidden, reversed,
        // cross-owner or unrepresentable endpoints rather than inventing one.
        self.comment_anchor(id)?
            .ok_or_else(|| Error::Other(format!("comment {id} has no movable paired range")))?;
        let mut selected = self.story_ranges()?.into_iter().filter(
            |entry| matches!(entry.kind(), StoryRangeKind::Comment { id: found } if *found == id),
        );
        let entry = selected
            .next()
            .ok_or_else(|| Error::Other(format!("comment {id} has no selected source range")))?;
        if selected.next().is_some() {
            return Err(Error::Other(format!(
                "comment {id} has ambiguous source ranges"
            )));
        }
        Ok(entry)
    }

    pub fn move_story_range(
        &mut self,
        selected: &StoryRangeRef,
        range: StoryRunRange,
    ) -> Result<()> {
        if !self.story_ranges()?.contains(selected) {
            return Err(Error::Other("selected story range is stale".to_owned()));
        }
        if let StoryRangeKind::Comment { id } = selected.kind() {
            let checked = self.checked_movable_comment(*id)?;
            if &checked != selected {
                return Err(Error::Other("selected comment range is stale".into()));
            }
            let mut candidate = self.clone_for_staging();
            candidate.move_comment_story_range(selected, &range, *id)?;
            candidate.flush_to_package()?;
            candidate.comment_ownership_at(&candidate.doc_part_name)?;
            candidate.story_ranges()?;
            let reopened = candidate.prepare_and_reopen_staged()?;
            self.commit_staged_mutation(reopened);
            return Ok(());
        }
        let mut candidate = self.clone_for_staging();
        candidate.remove_story_range_markers(selected, true)?;
        let anchor = match selected.kind() {
            StoryRangeKind::Bookmark { id, name } => RangeAnchor::Bookmark { id: *id, name },
            StoryRangeKind::Comment { id } => RangeAnchor::Comment(*id),
            StoryRangeKind::Permission { id, editor, group } => RangeAnchor::Permission {
                id: *id,
                editor: editor.as_deref(),
                group: group.as_deref(),
            },
            StoryRangeKind::Proofing { kind } => RangeAnchor::Proofing { kind },
        };
        let mut placement_check = self.clone_for_staging();
        placement_check.anchor_story_range(&range, anchor, "story", false)?;
        let refreshed = candidate.stories()?;
        let rebase = |position: &StoryRunPosition| -> Result<StoryRunPosition> {
            let source = position.location.story();
            let story = refreshed
                .iter()
                .find(|story| {
                    story.kind() == source.kind()
                        && story.part_name() == source.part_name()
                        && story.owner_index() == source.owner_index()
                })
                .ok_or_else(|| {
                    Error::Other("target story disappeared during range move".to_owned())
                })?;
            Ok(StoryRunPosition {
                location: ContentLocation::new(
                    story.clone(),
                    position.location.item_kind(),
                    position.location.index_path().to_vec(),
                ),
                run_index: position.run_index,
            })
        };
        let range = StoryRunRange {
            start: rebase(&range.start)?,
            end: rebase(&range.end)?,
        };
        candidate.anchor_story_range(&range, anchor, "story", false)?;
        candidate.story_ranges()?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(())
    }

    /// Remove one checked pair. Its attached comment definition, if any,
    /// remains a separate comment operation.
    pub fn remove_story_range(&mut self, selected: &StoryRangeRef) -> Result<bool> {
        if !self.story_ranges()?.contains(selected) {
            return Ok(false);
        }
        let mut candidate = self.clone_for_staging();
        candidate.remove_story_range_markers(selected, false)?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(true)
    }

    /// Return checked paired-marker ranges in physical story order.
    ///
    /// Unmatched, reversed and crossing markers are rejected. Unsupported
    /// marker families remain opaque and are not included.
    pub fn story_ranges(&self) -> Result<Vec<StoryRangeRef>> {
        let mut ranges = Vec::new();
        let mut open = Vec::<StoryMarker>::new();
        let mut story = None;
        let mut bookmark_names = HashSet::new();
        let mut bookmark_ids = HashSet::new();
        let mut comment_ids = HashSet::new();
        let mut permission_ids = HashSet::new();
        let mut marker_ordinal = 0usize;
        for (location, xml) in self.story_range_paragraphs()? {
            if story.as_ref() != Some(location.story()) {
                if !open.is_empty() {
                    return Err(Error::Other(
                        "paired marker crosses a story owner".to_owned(),
                    ));
                }
                story = Some(location.story().clone());
                marker_ordinal = 0;
            }
            for mut marker in scan_story_markers(&xml, &location)? {
                marker.ordinal = marker_ordinal;
                marker_ordinal += 1;
                if marker.start {
                    open.push(marker);
                    continue;
                }
                let start = open
                    .pop()
                    .ok_or_else(|| Error::Other("paired marker has an unmatched end".to_owned()))?;
                if !same_marker_identity(&start.kind, &marker.kind) {
                    return Err(Error::Other(
                        "paired markers cross or use different identities".to_owned(),
                    ));
                }
                if let StoryRangeKind::Bookmark { name, .. } = &start.kind
                    && (name.is_empty() || !bookmark_names.insert(name.clone()))
                {
                    return Err(Error::Other(
                        "bookmark name is missing or duplicated".to_owned(),
                    ));
                }
                let unique = match &start.kind {
                    StoryRangeKind::Bookmark { id, .. } => bookmark_ids.insert(*id),
                    StoryRangeKind::Comment { id } => comment_ids.insert(*id),
                    StoryRangeKind::Permission { id, editor, group } => {
                        if editor.is_some() == group.is_some() {
                            return Err(Error::Other(
                                "permission start needs one editor or group".to_owned(),
                            ));
                        }
                        permission_ids.insert(*id)
                    }
                    StoryRangeKind::Proofing { .. } => true,
                };
                if !unique {
                    return Err(Error::Other(
                        "paired marker identity is duplicated".to_owned(),
                    ));
                }
                ranges.push(StoryRangeRef {
                    kind: start.kind,
                    range: StoryRunRange {
                        start: StoryRunPosition {
                            location: start.location,
                            run_index: start.run_index,
                        },
                        end: StoryRunPosition {
                            location: marker.location,
                            run_index: marker.run_index,
                        },
                    },
                    start_ordinal: start.ordinal,
                    end_ordinal: marker.ordinal,
                });
            }
        }
        if !open.is_empty() {
            return Err(Error::Other(
                "paired marker has an unmatched start".to_owned(),
            ));
        }
        Ok(ranges)
    }

    pub(crate) fn ensure_fragment_comment_models_staged(&mut self) -> Result<String> {
        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        self.comments_part_name
            .clone()
            .ok_or_else(|| Error::Other("comments part name is missing".to_owned()))
    }

    pub(crate) fn fragment_comment_dependency_xml(
        source: Option<&CT_Comments>,
        source_extended: Option<&CT_CommentsEx>,
        ids: &[String],
    ) -> Result<Option<Vec<u8>>> {
        if ids.is_empty() {
            return Ok(None);
        }
        let source = source.ok_or_else(|| {
            Error::Other("document fragment references a missing comments part".to_owned())
        })?;
        let included_ids = selected_fragment_comment_ids(source, source_extended, ids)?;
        let mut selected = source.clone();
        selected
            .comments
            .retain(|comment| included_ids.contains(&comment.id));
        selected.extra_xml.clear();
        Ok(Some(selected.to_xml()?))
    }

    pub(crate) fn import_fragment_comments_staged(
        &mut self,
        source: Option<&CT_Comments>,
        source_extended: Option<&CT_CommentsEx>,
        ids: &[String],
    ) -> Result<BTreeMap<String, String>> {
        if ids.is_empty() {
            return Ok(BTreeMap::new());
        }
        let source = source.ok_or_else(|| {
            Error::Other("document fragment references a missing comments part".to_owned())
        })?;
        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        let included_ids = selected_fragment_comment_ids(source, source_extended, ids)?;

        let mut remap = BTreeMap::<String, String>::new();
        let mut paragraph_remap = HashMap::<String, String>::new();
        let mut reserved_para_ids =
            occupied_para_ids(self.comments.as_ref(), self.comments_extended.as_ref());
        for comment in source
            .comments
            .iter()
            .filter(|comment| included_ids.contains(&comment.id))
        {
            let new_id = self.identifiers.reserve_comment_id()?;
            remap.insert(comment.id.to_string(), new_id.to_string());
            for old_para_id in comment.paragraph_ids.iter().flatten() {
                let new_para_id = allocate_para_id_from_occupied(&mut reserved_para_ids)?;
                paragraph_remap.insert(old_para_id.clone(), new_para_id);
            }
        }

        let mut imported_comments = Vec::new();
        for source_comment in source
            .comments
            .iter()
            .filter(|comment| included_ids.contains(&comment.id))
        {
            let new_id = remap
                .get(&source_comment.id.to_string())
                .and_then(|id| id.parse::<i32>().ok())
                .ok_or_else(|| {
                    Error::Other("document fragment comment remap is incomplete".to_owned())
                })?;
            let mut imported = source_comment.clone();
            imported.id = new_id;
            for para_id in imported.paragraph_ids.iter_mut().flatten() {
                *para_id = paragraph_remap.get(para_id).cloned().ok_or_else(|| {
                    Error::Other("document fragment paragraph remap is incomplete".to_owned())
                })?;
            }
            imported_comments.push(imported);
        }
        self.comments
            .as_mut()
            .expect("fragment comment models were initialized")
            .comments
            .extend(imported_comments);

        if let Some(source_extended) = source_extended {
            let mut imported_extended = Vec::new();
            for extension in &source_extended.comments {
                let Some(para_id) = paragraph_remap.get(&extension.para_id).cloned() else {
                    continue;
                };
                let para_id_parent = extension
                    .para_id_parent
                    .as_ref()
                    .map(|parent| {
                        paragraph_remap.get(parent).cloned().ok_or_else(|| {
                            Error::Other(
                                "document fragment comment parent is outside the selected closure"
                                    .to_owned(),
                            )
                        })
                    })
                    .transpose()?;
                let mut imported = extension.clone();
                imported.para_id = para_id;
                imported.para_id_parent = para_id_parent;
                imported_extended.push(imported);
            }
            self.comments_extended
                .as_mut()
                .expect("fragment comment extension model was initialized")
                .comments
                .extend(imported_extended);
        }
        self.comments_dirty = true;
        Ok(remap)
    }

    /// Return bookmarks and malformed marker reports in main-story paragraph order.
    ///
    /// Reported body indexes count typed paragraphs recursively through tables and
    /// block content controls. Reported run indexes use accepted-view run boundaries.
    /// `BookmarkRef::direct_range` also reports the direct body child index when
    /// both markers sit in direct body paragraphs.
    pub fn bookmarks(&self) -> Vec<BookmarkRef> {
        #[derive(Clone)]
        struct Marker {
            position: RunPosition,
            direct_position: Option<RunPosition>,
            start: bool,
            id: Option<i32>,
            name: Option<String>,
        }

        let mut markers = Vec::new();
        let mut paragraphs = Vec::new();
        for (direct_index, item) in self.document.body.content.iter().enumerate() {
            let mut item_paragraphs = Vec::new();
            collect_main_story_paragraphs(std::slice::from_ref(item), &mut item_paragraphs);
            let direct_index = matches!(item, BodyContent::Paragraph(_)).then_some(direct_index);
            paragraphs.extend(
                item_paragraphs
                    .into_iter()
                    .map(|paragraph| (paragraph, direct_index)),
            );
        }
        for (body_index, (paragraph, direct_index)) in paragraphs.into_iter().enumerate() {
            for marker in &paragraph.bookmark_markers {
                let run_index = marker.projected_run_index();
                markers.push(Marker {
                    position: RunPosition {
                        body_index,
                        run_index,
                    },
                    direct_position: direct_index.map(|body_index| RunPosition {
                        body_index,
                        run_index,
                    }),
                    start: marker.is_start(),
                    id: marker.id(),
                    name: marker.name().map(str::to_owned),
                });
            }
        }

        let mut by_id: HashMap<i32, Vec<usize>> = HashMap::new();
        let mut results = Vec::new();
        for (index, marker) in markers.iter().enumerate() {
            if let Some(id) = marker.id {
                by_id.entry(id).or_default().push(index);
            } else {
                results.push((
                    index,
                    BookmarkRef {
                        id: None,
                        name: marker.name.clone(),
                        range: None,
                        direct_range: None,
                        text: String::new(),
                        issue: Some("bookmark marker has a malformed or missing id".to_owned()),
                    },
                ));
            }
        }

        for (id, indices) in by_id {
            let starts = indices
                .iter()
                .copied()
                .filter(|index| markers[*index].start)
                .collect::<Vec<_>>();
            let ends = indices
                .iter()
                .copied()
                .filter(|index| !markers[*index].start)
                .collect::<Vec<_>>();
            let first = indices.iter().copied().min().unwrap_or(0);
            let name = starts
                .first()
                .and_then(|index| markers[*index].name.clone());
            let (range, issue) = if starts.len() != 1 || ends.len() != 1 {
                (
                    None,
                    Some(format!(
                        "bookmark id {id} has {} start markers and {} end markers",
                        starts.len(),
                        ends.len()
                    )),
                )
            } else if name.is_none() {
                (None, Some(format!("bookmark id {id} has a missing name")))
            } else {
                let candidate = RunRange {
                    start: markers[starts[0]].position,
                    end: markers[ends[0]].position,
                };
                if candidate.start > candidate.end
                    || (candidate.start == candidate.end && starts[0] > ends[0])
                {
                    (
                        None,
                        Some(format!("bookmark id {id} ends before it starts")),
                    )
                } else {
                    (Some(candidate), None)
                }
            };
            let direct_range = range.and_then(|_| {
                Some(RunRange {
                    start: markers[starts[0]].direct_position?,
                    end: markers[ends[0]].direct_position?,
                })
            });
            let text = range
                .map(|_| {
                    bookmark_range_text(
                        &self.document.body.content,
                        markers[starts[0]].position.body_index,
                        markers[starts[0]].position.run_index,
                        markers[ends[0]].position.body_index,
                        markers[ends[0]].position.run_index,
                    )
                })
                .unwrap_or_default();
            results.push((
                first,
                BookmarkRef {
                    id: Some(id),
                    name,
                    range,
                    direct_range,
                    text,
                    issue,
                },
            ));
        }

        let mut name_counts = HashMap::new();
        for (_, bookmark) in &results {
            if bookmark.range.is_some()
                && let Some(name) = bookmark.name.as_deref()
            {
                *name_counts.entry(name.to_owned()).or_insert(0usize) += 1;
            }
        }
        for (_, bookmark) in &mut results {
            if bookmark
                .name
                .as_ref()
                .is_some_and(|name| name_counts.get(name).copied().unwrap_or(0) > 1)
            {
                bookmark.issue = Some(format!(
                    "bookmark name {} is duplicated",
                    bookmark.name.as_deref().unwrap_or("")
                ));
                bookmark.range = None;
                bookmark.direct_range = None;
                bookmark.text.clear();
            }
        }
        results.sort_by_key(|(index, _)| *index);
        results.into_iter().map(|(_, bookmark)| bookmark).collect()
    }

    /// Insert a bookmark over a half-open range of body paragraph runs.
    ///
    /// Run indexes count the runs that `Paragraph::runs` lists. The markers
    /// go inside an inline content control when the range starts or ends
    /// between two of its runs, and around the control when the range covers
    /// it. A range that cannot be anchored exactly, such as one that crosses
    /// the edge of a control or ends between two runs of a tracked insertion,
    /// is an error and leaves the document unchanged.
    pub fn add_bookmark(&mut self, name: &str, range: RunRange) -> Result<i32> {
        self.insert_bookmark(name, range)
    }

    /// The name of a bookmark around the whole of direct body paragraph
    /// `body_index`, adding one when it has none, so an internal hyperlink
    /// can target the paragraph as Word's and Google Docs' "link to a
    /// heading" does.
    ///
    /// An existing bookmark that covers exactly the paragraph's runs, such as
    /// Word's `_Toc` bookmark of a heading, is reused, and otherwise one that
    /// starts at the paragraph start. A new one is named
    /// after the paragraph text, such as `Heading_Results`, with a number
    /// added when that name is taken.
    pub fn heading_bookmark(&mut self, body_index: usize) -> Result<String> {
        let Some(BodyContent::Paragraph(paragraph)) = self.document.body.content.get(body_index)
        else {
            return Err(Error::Other(format!(
                "body item {body_index} is not a direct body paragraph"
            )));
        };
        let paragraph = crate::ParagraphRef { inner: paragraph };
        let range = RunRange {
            start: RunPosition {
                body_index,
                run_index: 0,
            },
            end: RunPosition {
                body_index,
                run_index: paragraph.run_count(),
            },
        };
        let text = paragraph.text();
        let bookmarks = self.bookmarks();
        // A bookmark around the paragraph, or else one that starts with it,
        // such as the point bookmark Google Docs puts before a heading.
        let starts_here = |bookmark: &&BookmarkRef| {
            bookmark
                .direct_range()
                .is_some_and(|candidate| candidate.start == range.start)
        };
        if let Some(name) = bookmarks
            .iter()
            .find(|bookmark| bookmark.direct_range() == Some(range))
            .or_else(|| bookmarks.iter().find(starts_here))
            .and_then(BookmarkRef::name)
        {
            return Ok(name.to_owned());
        }
        let mut stem = String::from("Heading_");
        for word in text.split(|character: char| !character.is_ascii_alphanumeric()) {
            if !word.is_empty() && stem.len() + word.len() < 34 {
                stem.push_str(word);
                stem.push('_');
            }
        }
        let stem = stem.trim_end_matches('_').to_owned();
        let taken = |name: &str| {
            bookmarks
                .iter()
                .any(|bookmark| bookmark.name() == Some(name))
        };
        let name = (1..)
            .map(|number| {
                if number == 1 {
                    stem.clone()
                } else {
                    format!("{stem}_{number}")
                }
            })
            .find(|name| !taken(name))
            .expect("an unbounded sequence holds a free name");
        self.insert_bookmark(&name, range)?;
        Ok(name)
    }

    fn insert_bookmark(&mut self, name: &str, range: RunRange) -> Result<i32> {
        validate_bookmark_name(name)?;
        validate_bookmark_range(&self.document.body.content, range)?;
        if self
            .bookmarks()
            .iter()
            .any(|bookmark| bookmark.name() == Some(name))
        {
            return Err(Error::Other(format!("bookmark name {name} already exists")));
        }
        let mut identifiers = self.identifiers.clone();
        let id = identifiers.reserve_bookmark_id()?;
        anchor_body_range(
            &mut self.document.body.content,
            range,
            RangeAnchor::Bookmark { id, name },
            "bookmark",
        )?;
        self.identifiers = identifiers;
        self.invalidate_layout();
        Ok(id)
    }

    /// Return comments in their package part order.
    pub fn comments(&self) -> Vec<CommentRef<'_>> {
        let Some(comments) = self.comments.as_ref() else {
            return Vec::new();
        };
        let by_para_id = comments
            .comments
            .iter()
            .flat_map(|comment| para_ids(comment).map(|para_id| (para_id, comment.id)))
            .collect::<HashMap<_, _>>();
        comments
            .comments
            .iter()
            .map(|comment| {
                let extension = last_para_id(comment).and_then(|para_id| {
                    self.comments_extended
                        .as_ref()?
                        .comments
                        .iter()
                        .find(|entry| entry.para_id == para_id)
                });
                let parent_id = extension
                    .and_then(|entry| entry.para_id_parent.as_deref())
                    .and_then(|para_id| by_para_id.get(para_id).copied());
                CommentRef {
                    document: self,
                    inner: comment,
                    extension,
                    parent_id,
                }
            })
            .collect()
    }

    /// Add a comment over a half-open range of body paragraph runs.
    ///
    /// Run indexes count the runs that `Paragraph::runs` lists, and the
    /// range markers are placed as [`Self::add_bookmark`] places its markers.
    /// The reference run follows the end marker. A range that cannot be
    /// anchored exactly is an error and leaves the document unchanged. Each
    /// line of `text` becomes one paragraph of the comment.
    pub fn add_comment(
        &mut self,
        range: RunRange,
        author: &str,
        initials: Option<&str>,
        text: &str,
    ) -> Result<i32> {
        self.add_comment_with_date(range, author, initials, text, None)
    }

    /// Add a dated comment over a half-open range of body paragraph runs.
    ///
    /// `date`, when present, must be an RFC 3339 timestamp. No date is the
    /// deterministic default used by [`Document::add_comment`]. Each line of
    /// `text` becomes one paragraph of the comment.
    pub fn add_comment_with_date(
        &mut self,
        range: RunRange,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        reject_non_xml_comment(author, initials, text)?;
        let mut candidate = self.clone_for_staging();
        let id = candidate.add_comment_staged(range, author, initials, text, date)?;
        candidate.flush_dirty_related_story_models()?;
        self.commit_staged_mutation(candidate);
        Ok(id)
    }

    /// Add a dated comment over a checked body, table-cell, header, footer, or note run range.
    ///
    /// A body location can also name a paragraph inside a block content
    /// control with the two-segment path that
    /// [`Document::paragraph_story_location`] returns. Run indexes count the
    /// runs that `Paragraph::runs` lists, as in [`Document::add_comment`].
    /// Each line of `text` becomes one paragraph of the comment.
    pub fn add_story_comment_with_date(
        &mut self,
        range: StoryRunRange,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        let mut candidate = self.clone_for_staging();
        let id = candidate.add_story_comment_staged(range, author, initials, text, date)?;
        let reopened = candidate.prepare_and_reopen_staged()?;
        reopened.story_ranges()?;
        self.commit_staged_mutation(reopened);
        Ok(id)
    }

    /// Add a comment over a checked body, table-cell, header, footer, or note run range, as
    /// [`Self::add_story_comment_with_date`] does without a date.
    pub fn add_story_comment(
        &mut self,
        range: StoryRunRange,
        author: &str,
        initials: Option<&str>,
        text: &str,
    ) -> Result<i32> {
        self.add_story_comment_with_date(range, author, initials, text, None)
    }

    fn add_story_comment_staged(
        &mut self,
        range: StoryRunRange,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        validate_comment_date(date)?;
        if range.start.location.story() != range.end.location.story() {
            return Err(Error::Other(
                "comment story range start must not follow its end".to_owned(),
            ));
        }
        if range.start.location != range.end.location
            || range.start.location.index_path().len() == 2
            || range.end.location.index_path().len() == 2
            || matches!(
                range.start.location.story().kind(),
                crate::StoryKind::Header
                    | crate::StoryKind::Footer
                    | crate::StoryKind::Footnote
                    | crate::StoryKind::Endnote
                    | crate::StoryKind::Comment
                    | crate::StoryKind::TextBox
            )
        {
            let mut identifiers = self.identifiers.clone();
            let id = identifiers.reserve_comment_id()?;
            self.anchor_story_range(&range, RangeAnchor::Comment(id), "comment", false)?;
            self.ensure_comment_models()?;
            self.ensure_comment_relationships()?;
            self.push_comment_definition(id, author, initials, text, date)?;
            self.identifiers = identifiers;
            self.comments_dirty = true;
            self.invalidate_layout();
            return Ok(id);
        }
        let mut start = self.story_paragraph_mut(&range.start.location)?.clone();
        let mut end = self.story_paragraph_mut(&range.end.location)?.clone();
        for (label, position, paragraph) in
            [("start", &range.start, &start), ("end", &range.end, &end)]
        {
            let run_count = paragraph.accepted_run_paths().len();
            if position.run_index > run_count {
                return Err(Error::Other(format!(
                    "comment range {label} run index {} exceeds paragraph run count {run_count}",
                    position.run_index,
                )));
            }
        }
        if range.start.location == range.end.location && range.start.run_index > range.end.run_index
        {
            return Err(Error::Other(
                "comment story range start must not follow its end".to_owned(),
            ));
        }

        let mut identifiers = self.identifiers.clone();
        let id = identifiers.reserve_comment_id()?;
        if range.start.location == range.end.location {
            anchor_paragraph_range(
                &mut start,
                Some(range.start.run_index),
                Some(range.end.run_index),
                RangeAnchor::Comment(id),
                "comment",
            )?;
            *self.story_paragraph_mut(&range.start.location)? = start;
        } else {
            anchor_paragraph_range(
                &mut start,
                Some(range.start.run_index),
                None,
                RangeAnchor::Comment(id),
                "comment",
            )?;
            anchor_paragraph_range(
                &mut end,
                None,
                Some(range.end.run_index),
                RangeAnchor::Comment(id),
                "comment",
            )?;
            *self.story_paragraph_mut(&range.start.location)? = start;
            *self.story_paragraph_mut(&range.end.location)? = end;
        }

        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        self.push_comment_definition(id, author, initials, text, date)?;
        self.identifiers = identifiers;
        self.comments_dirty = true;
        self.invalidate_layout();
        Ok(id)
    }

    fn add_comment_staged(
        &mut self,
        range: RunRange,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        validate_comment_date(date)?;
        self.validate_run_range(range)?;
        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        let mut identifiers = self.identifiers.clone();
        let id = identifiers.reserve_comment_id()?;
        self.push_comment_definition(id, author, initials, text, date)?;
        anchor_body_range(
            &mut self.document.body.content,
            range,
            RangeAnchor::Comment(id),
            "comment",
        )?;
        self.identifiers = identifiers;
        self.comments_dirty = true;
        self.invalidate_layout();
        Ok(id)
    }

    /// Add a comment on the `occurrence`-th match of `anchor`, counted from
    /// zero, in the main story.
    ///
    /// Matches are case-sensitive and non-overlapping, in document order
    /// through body paragraphs, tables and block content controls, and one
    /// match never spans two paragraphs. They are found in the literal run
    /// text that [`Document::split_run`] offsets count, so tabs and breaks
    /// have no width. The runs at both ends of the match are split and the
    /// comment is anchored on the runs between the splits as
    /// [`Self::add_comment`] anchors a run range. `date`, when present, must
    /// be an RFC 3339 timestamp. A missing occurrence or a match that cannot
    /// be anchored exactly is an error and leaves the document unchanged. A
    /// match is not exact when its range would also show text that the
    /// literal text leaves out, such as the result of a field between two of
    /// its runs. Each line of `text` becomes one paragraph of the comment.
    pub fn add_comment_on_text(
        &mut self,
        anchor: &str,
        occurrence: usize,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        reject_non_xml_comment(author, initials, text)?;
        let mut candidate = self.clone_for_staging();
        let id = candidate
            .add_comment_on_text_staged(anchor, occurrence, author, initials, text, date)?;
        candidate.flush_dirty_related_story_models()?;
        self.commit_staged_mutation(candidate);
        Ok(id)
    }

    fn add_comment_on_text_staged(
        &mut self,
        anchor: &str,
        occurrence: usize,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        validate_comment_date(date)?;
        if anchor.is_empty() {
            return Err(Error::Other(
                "comment anchor text must not be empty".to_owned(),
            ));
        }
        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        let mut identifiers = self.identifiers.clone();
        let id = identifiers.reserve_comment_id()?;
        self.anchor_existing_comment_on_text(id, anchor, occurrence)?;
        self.push_comment_definition(id, author, initials, text, date)?;
        self.identifiers = identifiers;
        self.comments_dirty = true;
        self.invalidate_layout();
        Ok(id)
    }

    // Shared existing literal finder and checked splitter. Moving a comment
    // supplies its identity directly and never allocates a temporary thread.
    fn anchor_existing_comment_on_text(
        &mut self,
        id: i32,
        anchor: &str,
        occurrence: usize,
    ) -> Result<()> {
        if anchor.is_empty() {
            return Err(Error::Other("comment anchor text must not be empty".into()));
        }
        let mut remaining = occurrence;
        let mut anchored = None;
        visit_body_paragraphs_mut(&mut self.document.body.content, &mut |paragraph| {
            if anchored.is_some() {
                return;
            }
            let literal = paragraph.accepted_literal_text();
            let matches = literal
                .match_indices(anchor)
                .map(|(byte, _)| byte)
                .collect::<Vec<_>>();
            let Some(byte) = matches.get(remaining) else {
                remaining -= matches.len();
                return;
            };
            let start = literal[..*byte].chars().count();
            let end = start + anchor.chars().count();
            anchored = Some(
                paragraph
                    .split_accepted_literal_span(start, end)
                    .map_err(|error| {
                        Error::Other(format!("comment anchor text cannot be split: {error}"))
                    })
                    .and_then(|(start, end)| {
                        anchor_paragraph_range(
                            paragraph,
                            Some(start),
                            Some(end),
                            RangeAnchor::Comment(id),
                            "comment",
                        )
                    }),
            );
        });
        anchored.ok_or_else(|| {
            // Every paragraph was searched, so `remaining` counts past them all.
            let found = occurrence - remaining;
            let times = if found == 1 { "time" } else { "times" };
            Error::Other(format!(
                "comment anchor text {anchor:?} has no occurrence {occurrence}: it occurs {found} {times} in the main story"
            ))
        })??;
        // Read actual namespace scopes after the mutable walk releases its borrow.
        // Preserved field results count, while tabs and breaks have zero width.
        let shown = self.comment_literal_range_text(id)?.unwrap_or_default();
        if shown != anchor {
            return Err(Error::Other(format!(
                "comment anchor text {anchor:?} occurrence {occurrence} cannot be anchored exactly: its range would show {shown:?}"
            )));
        }
        self.invalidate_layout();
        Ok(())
    }

    /// Append comment `id`, holding one paragraph per line of `text`, and its
    /// thread entry to the comment models, which must already exist.
    fn push_comment_definition(
        &mut self,
        id: i32,
        author: &str,
        initials: Option<&str>,
        text: &str,
        date: Option<&str>,
    ) -> Result<()> {
        let mut occupied =
            occupied_para_ids(self.comments.as_ref(), self.comments_extended.as_ref());
        let (paragraphs, paragraph_ids, para_id) = comment_text_paragraphs(text, &mut occupied)?;
        self.comments
            .as_mut()
            .expect("comment model was initialized")
            .comments
            .push(CT_Comment {
                id,
                author: Some(author.to_owned()),
                date: date.map(str::to_owned),
                initials: initials.map(str::to_owned),
                paragraphs,
                paragraph_ids,
                extra_attributes: Vec::new(),
                extra_xml: Vec::new(),
            });
        self.comments_extended
            .as_mut()
            .expect("comments-extended model was initialized")
            .comments
            .push(CT_CommentEx {
                para_id,
                para_id_parent: None,
                done: None,
                extra_attributes: Vec::new(),
            });
        Ok(())
    }

    /// Add a reply linked to the selected comment's last paragraph. Each line
    /// of `text` becomes one paragraph of the reply.
    pub fn reply_to(&mut self, parent_id: i32, author: &str, text: &str) -> Result<i32> {
        self.reply_to_with_date(parent_id, author, text, None)
    }

    /// Add a dated reply linked to the selected comment's last paragraph.
    ///
    /// `date`, when present, must be an RFC 3339 timestamp. Each line of
    /// `text` becomes one paragraph of the reply.
    pub fn reply_to_with_date(
        &mut self,
        parent_id: i32,
        author: &str,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        let mut candidate = self.clone_for_staging();
        let id = candidate.reply_to_staged(parent_id, author, text, date)?;
        candidate.flush_dirty_related_story_models()?;
        self.commit_staged_mutation(candidate);
        Ok(id)
    }

    fn reply_to_staged(
        &mut self,
        parent_id: i32,
        author: &str,
        text: &str,
        date: Option<&str>,
    ) -> Result<i32> {
        validate_comment_date(date)?;
        let parent_index = self
            .comments
            .as_ref()
            .and_then(|comments| {
                comments
                    .comments
                    .iter()
                    .position(|item| item.id == parent_id)
            })
            .ok_or_else(|| Error::Other(format!("comment id {parent_id} does not exist")))?;
        let existing_parent_para_id = last_para_id(
            &self
                .comments
                .as_ref()
                .expect("parent lookup proved the model exists")
                .comments[parent_index],
        )
        .map(str::to_owned);
        let mut occupied =
            occupied_para_ids(self.comments.as_ref(), self.comments_extended.as_ref());
        let parent_para_id = match existing_parent_para_id {
            Some(para_id) => para_id,
            None => allocate_para_id_from_occupied(&mut occupied)?,
        };
        let (paragraphs, paragraph_ids, para_id) = comment_text_paragraphs(text, &mut occupied)?;
        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        let id = self.identifiers.reserve_comment_id()?;

        let comments = self.comments.as_mut().expect("model was initialized");
        let parent = &mut comments.comments[parent_index];
        if parent.paragraphs.is_empty() {
            parent.paragraphs.push(CT_P::new());
        }
        if parent.paragraph_ids.len() < parent.paragraphs.len() {
            parent.paragraph_ids.resize(parent.paragraphs.len(), None);
        }
        parent.paragraph_ids[parent.paragraphs.len() - 1] = Some(parent_para_id.clone());
        comments.comments.push(CT_Comment {
            id,
            author: Some(author.to_owned()),
            date: date.map(str::to_owned),
            initials: None,
            paragraphs,
            paragraph_ids,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
        });

        let extended = self
            .comments_extended
            .as_mut()
            .expect("model was initialized");
        if extended
            .comments
            .iter()
            .all(|entry| entry.para_id != parent_para_id)
        {
            extended.comments.push(CT_CommentEx {
                para_id: parent_para_id.clone(),
                para_id_parent: None,
                done: None,
                extra_attributes: Vec::new(),
            });
        }
        extended.comments.push(CT_CommentEx {
            para_id,
            para_id_parent: Some(parent_para_id),
            done: None,
            extra_attributes: Vec::new(),
        });
        self.comments_dirty = true;
        self.invalidate_layout();
        Ok(id)
    }

    /// Set or clear the resolved state for a comment.
    pub fn resolve_comment(&mut self, id: i32, resolved: bool) -> Result<bool> {
        let mut candidate = self.clone_for_staging();
        let updated = candidate.resolve_comment_staged(id, resolved)?;
        if updated {
            candidate.flush_dirty_related_story_models()?;
            self.commit_staged_mutation(candidate);
        }
        Ok(updated)
    }

    fn resolve_comment_staged(&mut self, id: i32, resolved: bool) -> Result<bool> {
        let Some(comment_index) = self
            .comments
            .as_ref()
            .and_then(|comments| comments.comments.iter().position(|item| item.id == id))
        else {
            return Ok(false);
        };
        let existing_para_id = last_para_id(
            &self
                .comments
                .as_ref()
                .expect("comment lookup proved the model exists")
                .comments[comment_index],
        )
        .map(str::to_owned);
        let para_id = match existing_para_id {
            Some(para_id) => para_id,
            None => allocate_para_id(self.comments.as_ref(), self.comments_extended.as_ref())?,
        };
        self.ensure_comment_models()?;
        self.ensure_comment_relationships()?;
        let comment = &mut self
            .comments
            .as_mut()
            .expect("model was initialized")
            .comments[comment_index];
        if comment.paragraphs.is_empty() {
            comment.paragraphs.push(CT_P::new());
        }
        if comment.paragraph_ids.len() < comment.paragraphs.len() {
            comment.paragraph_ids.resize(comment.paragraphs.len(), None);
        }
        let last = comment.paragraphs.len() - 1;
        comment.paragraph_ids[last] = Some(para_id.clone());
        let extended = self
            .comments_extended
            .as_mut()
            .expect("model was initialized");
        let root_para_id = thread_root_para_id(extended, &para_id);
        if let Some(entry) = extended
            .comments
            .iter_mut()
            .find(|entry| entry.para_id == root_para_id)
        {
            entry.done = Some(resolved);
        } else {
            extended.comments.push(CT_CommentEx {
                para_id: root_para_id,
                para_id_parent: None,
                done: Some(resolved),
                extra_attributes: Vec::new(),
            });
        }
        self.comments_dirty = true;
        self.invalidate_layout();
        Ok(true)
    }

    /// Remove a qualified comment thread and its source anchors atomically.
    /// Unsupported or ambiguous companion ownership leaves the document unchanged.
    pub fn remove_comment(&mut self, id: i32) -> Result<bool> {
        if !self.comments().iter().any(|comment| comment.id() == id) {
            return Ok(false);
        }
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let ownership = candidate.comment_ownership_at(&candidate.doc_part_name)?;
        let removed = ownership.descendants(id);
        candidate.remove_comment_ids_staged_at(&candidate.doc_part_name.clone(), &removed)?;
        // An inverse add/remove can recover the complete retained producer source.
        // Compare every modeled value and raw child, ignoring only declarations
        // introduced by the intermediate canonical serialization.
        if let Some(retained) = self.package.get_part(&self.doc_part_name)
            && let Ok(original) = rdocx_oxml::document::CT_Document::from_xml(retained)
        {
            let declarations_preserved = original
                .extra_namespaces
                .iter()
                .all(|binding| candidate.document.extra_namespaces.contains(binding));
            let only_added_canonical_wp = candidate.document.extra_namespaces.iter().all(|binding|
                original.extra_namespaces.contains(binding)
                    || (binding.0 == "xmlns:wp"
                        && binding.1 == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        && !original.extra_namespaces.iter().any(|original| original.0 == binding.0)));
            let mut restored = candidate.document.clone();
            restored.extra_namespaces = original.extra_namespaces.clone();
            // Full canonical XML equality includes all modeled values and opaque
            // bytes while ignoring unused inherited-prefix projection caches.
            if declarations_preserved
                && only_added_canonical_wp
                && restored.to_xml()? == original.to_xml()?
            {
                let part = candidate.doc_part_name.clone();
                crate::document::set_story_source_xml(&mut candidate, &part, retained.to_vec())?;
            }
        }
        let reopened = candidate.prepare_and_reopen_staged()?;
        self.commit_staged_mutation(reopened);
        Ok(true)
    }

    fn validate_run_range(&self, range: RunRange) -> Result<()> {
        if range.start > range.end {
            return Err(Error::Other(
                "comment range start must not follow its end".to_owned(),
            ));
        }
        for (label, position) in [("start", range.start), ("end", range.end)] {
            let paragraph = body_paragraph(&self.document.body.content, position.body_index)
                .ok_or_else(|| {
                    Error::Other(format!(
                        "comment range {label} body index {} is not a paragraph",
                        position.body_index
                    ))
                })?;
            let run_count = paragraph.accepted_run_paths().len();
            if position.run_index > run_count {
                return Err(Error::Other(format!(
                    "comment range {label} run index {} exceeds paragraph run count {run_count}",
                    position.run_index,
                )));
            }
        }
        Ok(())
    }

    fn ensure_comment_models(&mut self) -> Result<()> {
        self.identifiers.observe_package_graph(&self.package)?;
        if self.comments.is_none() {
            self.comments = Some(CT_Comments::new());
        }
        if self.comments_part_name.is_none() {
            self.comments_part_name = Some(
                self.identifiers
                    .reserve_preferred_part_name(DEFAULT_COMMENTS_PART)?,
            );
            self.comments_owned = true;
        }
        if self.comments_extended.is_none() {
            self.comments_extended = Some(CT_CommentsEx::new());
        }
        if self.comments_extended_part_name.is_none() {
            self.comments_extended_part_name = Some(
                self.identifiers
                    .reserve_preferred_part_name(DEFAULT_COMMENTS_EXTENDED_PART)?,
            );
            self.comments_extended_owned = true;
        }
        Ok(())
    }

    fn ensure_comment_relationships(&mut self) -> Result<()> {
        let comments_part = self
            .comments_part_name
            .clone()
            .ok_or_else(|| Error::Other("comments part name is missing".to_owned()))?;
        let comments_extended_part = self
            .comments_extended_part_name
            .clone()
            .ok_or_else(|| Error::Other("comments-extended part name is missing".to_owned()))?;
        self.ensure_part_relationship_checked(
            &comments_part,
            oxml_opc::relationship::rel_types::COMMENTS,
            COMMENTS_CONTENT_TYPE,
        )
        .map_err(|error| {
            Error::Other(format!("comments relationship allocation failed: {error}"))
        })?;
        self.ensure_part_relationship_checked(
            &comments_extended_part,
            COMMENTS_EXTENDED_REL_TYPE,
            COMMENTS_EXTENDED_CONTENT_TYPE,
        )
        .map_err(|error| {
            Error::Other(format!(
                "comments-extended relationship allocation failed: {error}"
            ))
        })?;
        Ok(())
    }

    fn remove_owned_empty_comment_parts(&mut self) {
        if self.comments.as_ref().is_some_and(|comments| {
            comments.comments.is_empty()
                && comments.extra_xml.is_empty()
                && comments
                    .root_attributes
                    .iter()
                    .all(|(key, _)| key == "xmlns" || key.starts_with("xmlns:"))
        }) && self.comments_owned
        {
            if let Some(part) = self.comments_part_name.take() {
                remove_owned_part(self, &part, oxml_opc::relationship::rel_types::COMMENTS);
            }
            self.comments = None;
            self.comments_owned = false;
        }
        if self.comments_extended.as_ref().is_some_and(|extended| {
            extended.comments.is_empty()
                && extended.extra_xml.is_empty()
                && extended
                    .root_attributes
                    .iter()
                    .all(|(key, _)| key == "xmlns" || key.starts_with("xmlns:"))
        }) && self.comments_extended_owned
        {
            if let Some(part) = self.comments_extended_part_name.take() {
                remove_owned_part(self, &part, COMMENTS_EXTENDED_REL_TYPE);
            }
            self.comments_extended = None;
            self.comments_extended_owned = false;
        }
    }
}

fn validate_comment_date(date: Option<&str>) -> Result<()> {
    if let Some(value) = date
        && crate::revision::parse_rfc3339(value).is_none()
    {
        return Err(Error::Other(format!(
            "invalid RFC 3339 comment timestamp: {value}"
        )));
    }
    Ok(())
}

fn selected_fragment_comment_ids(
    source: &CT_Comments,
    source_extended: Option<&CT_CommentsEx>,
    ids: &[String],
) -> Result<HashSet<i32>> {
    let mut included_ids = ids
        .iter()
        .map(|id| {
            id.parse::<i32>()
                .map_err(|_| Error::Other(format!("document fragment comment id {id} is invalid")))
        })
        .collect::<Result<HashSet<_>>>()?;
    for id in &included_ids {
        if !source.comments.iter().any(|comment| comment.id == *id) {
            return Err(Error::Other(format!(
                "document fragment comment id {id} has no definition"
            )));
        }
    }
    if let Some(extended) = source_extended {
        loop {
            let included_para_ids = source
                .comments
                .iter()
                .filter(|comment| included_ids.contains(&comment.id))
                .flat_map(para_ids)
                .collect::<HashSet<_>>();
            let before = included_ids.len();
            for extension in &extended.comments {
                if extension
                    .para_id_parent
                    .as_deref()
                    .is_some_and(|parent| included_para_ids.contains(parent))
                    && let Some(comment) = source
                        .comments
                        .iter()
                        .find(|comment| para_ids(comment).any(|para| para == extension.para_id))
                {
                    included_ids.insert(comment.id);
                }
            }
            if included_ids.len() == before {
                break;
            }
        }
    }
    Ok(included_ids)
}

/// The `w14:paraId` of the comment's last paragraph, which keys its
/// `w15:commentEx` entry and names it as a reply's `w15:paraIdParent`.
fn last_para_id(comment: &CT_Comment) -> Option<&str> {
    let last = comment.paragraphs.len().checked_sub(1)?;
    comment.paragraph_ids.get(last)?.as_deref()
}

/// Every `w14:paraId` of the comment. Identifiers are unique, so a parent
/// link naming any of them, as older producers wrote, still finds it.
fn para_ids(comment: &CT_Comment) -> impl Iterator<Item = &str> {
    comment.paragraph_ids.iter().filter_map(Option::as_deref)
}

/// Build one comment paragraph per line of `text`, each with a fresh
/// `w14:paraId`, and return the last one's identifier with them.
fn comment_text_paragraphs(
    text: &str,
    occupied: &mut HashSet<u32>,
) -> Result<(Vec<CT_P>, Vec<Option<String>>, String)> {
    let mut paragraphs = Vec::new();
    let mut paragraph_ids = Vec::new();
    for line in text.split('\n') {
        let mut paragraph = CT_P::new();
        paragraph.add_run(line.strip_suffix('\r').unwrap_or(line));
        paragraphs.push(paragraph);
        paragraph_ids.push(Some(allocate_para_id_from_occupied(occupied)?));
    }
    let para_id = paragraph_ids
        .last()
        .cloned()
        .flatten()
        .expect("splitting text yields at least one line");
    Ok((paragraphs, paragraph_ids, para_id))
}

fn validate_bookmark_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::Other("bookmark name must not be empty".to_owned()));
    }
    if name.starts_with('_') {
        return Err(Error::Other(format!(
            "bookmark name {name} is reserved for producer use"
        )));
    }
    if name.len() > 40
        || !name
            .chars()
            .next()
            .is_some_and(|character| character == '_' || character.is_ascii_alphabetic())
        || name
            .chars()
            .any(|character| !(character == '_' || character.is_ascii_alphanumeric()))
    {
        return Err(Error::Other(format!(
            "bookmark name {name} is not a valid Word bookmark name"
        )));
    }
    Ok(())
}

fn validate_bookmark_range(content: &[BodyContent], range: RunRange) -> Result<()> {
    if range.start > range.end {
        return Err(Error::Other(
            "bookmark range start must not follow its end".to_owned(),
        ));
    }
    for (label, position) in [("start", range.start), ("end", range.end)] {
        let paragraph = body_paragraph(content, position.body_index).ok_or_else(|| {
            Error::Other(format!(
                "bookmark range {label} body index {} is not a paragraph",
                position.body_index
            ))
        })?;
        let run_count = paragraph.accepted_run_paths().len();
        if position.run_index > run_count {
            return Err(Error::Other(format!(
                "bookmark range {label} run index {} exceeds paragraph run count {run_count}",
                position.run_index,
            )));
        }
    }
    Ok(())
}

fn bookmark_range_text(
    content: &[BodyContent],
    start_body_index: usize,
    start_run_index: usize,
    end_body_index: usize,
    end_run_index: usize,
) -> String {
    let mut paragraphs = Vec::<String>::new();
    for body_index in start_body_index..=end_body_index {
        let Some(paragraph) = story_paragraph(content, body_index) else {
            continue;
        };
        let runs = paragraph.accepted_bookmark_runs();
        let start = if body_index == start_body_index {
            start_run_index
        } else {
            0
        };
        let end = if body_index == end_body_index {
            end_run_index
        } else {
            runs.len()
        };
        let start = start.min(runs.len());
        let end = end.min(runs.len());
        paragraphs.push(runs[start..end].iter().map(|run| run.text()).collect());
    }
    paragraphs.join("\n")
}

fn story_paragraph(content: &[BodyContent], index: usize) -> Option<&CT_P> {
    let mut remaining = index;
    for item in content {
        if let Some(paragraph) = paragraph_in_body_content(item, &mut remaining) {
            return Some(paragraph);
        }
    }
    None
}

fn body_paragraph_mut(content: &mut [BodyContent], index: usize) -> Option<&mut CT_P> {
    match content.get_mut(index)? {
        BodyContent::Paragraph(paragraph) => Some(paragraph),
        BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => None,
    }
}

fn body_paragraph(content: &[BodyContent], index: usize) -> Option<&CT_P> {
    match content.get(index)? {
        BodyContent::Paragraph(paragraph) => Some(paragraph),
        BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => None,
    }
}

/// Write range markers over a validated body range. The paragraphs change
/// only when every marker can be placed exactly.
fn anchor_body_range(
    content: &mut [BodyContent],
    range: RunRange,
    anchor: RangeAnchor<'_>,
    label: &str,
) -> Result<()> {
    let (start, end) = (range.start, range.end);
    let paragraph = |content: &[BodyContent], index| {
        body_paragraph(content, index)
            .cloned()
            .expect("range was validated")
    };
    let mut first = paragraph(content, start.body_index);
    if start.body_index == end.body_index {
        anchor_paragraph_range(
            &mut first,
            Some(start.run_index),
            Some(end.run_index),
            anchor,
            label,
        )?;
    } else {
        let mut last = paragraph(content, end.body_index);
        anchor_paragraph_range(&mut first, Some(start.run_index), None, anchor, label)?;
        anchor_paragraph_range(&mut last, None, Some(end.run_index), anchor, label)?;
        *body_paragraph_mut(content, end.body_index).expect("range was validated") = last;
    }
    *body_paragraph_mut(content, start.body_index).expect("range was validated") = first;
    Ok(())
}

/// Write range markers at accepted-view run boundaries, the run index space
/// that `Paragraph::runs` lists. A missing side continues in another
/// paragraph.
fn anchor_paragraph_range(
    paragraph: &mut CT_P,
    start: Option<usize>,
    end: Option<usize>,
    anchor: RangeAnchor<'_>,
    label: &str,
) -> Result<()> {
    paragraph
        .anchor_accepted_range(start, end, anchor)
        .map_err(|error| Error::Other(format!("{label} range cannot be anchored: {error}")))
}

fn collect_main_story_paragraphs<'a>(content: &'a [BodyContent], output: &mut Vec<&'a CT_P>) {
    for item in content {
        match item {
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
            .filter(|(at, _, _)| *at == boundary)
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
            .filter(|(at, _, _)| *at == boundary)
        {
            collect_control_paragraphs(control, BlockControlOwner::Row, output);
        }
        if let Some(cell) = row.cells.get(boundary) {
            collect_cell_paragraphs(cell, output);
        }
    }
}

fn collect_cell_paragraphs<'a>(cell: &'a CT_Tc, output: &mut Vec<&'a CT_P>) {
    for item in &cell.content {
        match item {
            CellContent::Paragraph(paragraph) => output.push(paragraph),
            CellContent::Table(table) => collect_table_paragraphs(table, output),
            CellContent::ContentControl(control) => {
                collect_control_paragraphs(control, BlockControlOwner::Cell, output)
            }
        }
    }
}

#[derive(Clone, Copy)]
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
    for item in &control.content {
        match (owner, item) {
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

fn paragraph_in_body_content<'a>(
    content: &'a BodyContent,
    remaining: &mut usize,
) -> Option<&'a CT_P> {
    match content {
        BodyContent::Paragraph(paragraph) => take_paragraph(paragraph, remaining),
        BodyContent::Table(table) => paragraph_in_table(table, remaining),
        BodyContent::ContentControl(control) => {
            paragraph_in_control(control, BlockControlOwner::Body, remaining)
        }
        BodyContent::RawXml(_) => None,
    }
}

fn paragraph_in_table<'a>(table: &'a CT_Tbl, remaining: &mut usize) -> Option<&'a CT_P> {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            if let Some(paragraph) =
                paragraph_in_control(control, BlockControlOwner::Table, remaining)
            {
                return Some(paragraph);
            }
        }
        if let Some(row) = table.rows.get(boundary)
            && let Some(paragraph) = paragraph_in_row(row, remaining)
        {
            return Some(paragraph);
        }
    }
    None
}

fn paragraph_in_row<'a>(row: &'a CT_Row, remaining: &mut usize) -> Option<&'a CT_P> {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            if let Some(paragraph) =
                paragraph_in_control(control, BlockControlOwner::Row, remaining)
            {
                return Some(paragraph);
            }
        }
        if let Some(cell) = row.cells.get(boundary)
            && let Some(paragraph) = paragraph_in_cell(cell, remaining)
        {
            return Some(paragraph);
        }
    }
    None
}

fn paragraph_in_cell<'a>(cell: &'a CT_Tc, remaining: &mut usize) -> Option<&'a CT_P> {
    for content in &cell.content {
        let paragraph = match content {
            CellContent::Paragraph(paragraph) => take_paragraph(paragraph, remaining),
            CellContent::Table(table) => paragraph_in_table(table, remaining),
            CellContent::ContentControl(control) => {
                paragraph_in_control(control, BlockControlOwner::Cell, remaining)
            }
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn paragraph_in_control<'a>(
    control: &'a CT_Sdt,
    owner: BlockControlOwner,
    remaining: &mut usize,
) -> Option<&'a CT_P> {
    for content in &control.content {
        let paragraph = match (owner, content) {
            (
                BlockControlOwner::Body | BlockControlOwner::Cell,
                SdtContent::Paragraph(paragraph),
            ) => take_paragraph(paragraph, remaining),
            (BlockControlOwner::Body | BlockControlOwner::Cell, SdtContent::Table(table)) => {
                paragraph_in_table(table, remaining)
            }
            (BlockControlOwner::Table, SdtContent::Row(row)) => paragraph_in_row(row, remaining),
            (BlockControlOwner::Row, SdtContent::Cell(cell)) => paragraph_in_cell(cell, remaining),
            (_, SdtContent::ContentControl(control)) => {
                paragraph_in_control(control, owner, remaining)
            }
            _ => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn take_paragraph<'a>(paragraph: &'a CT_P, remaining: &mut usize) -> Option<&'a CT_P> {
    if *remaining == 0 {
        Some(paragraph)
    } else {
        *remaining -= 1;
        None
    }
}

fn thread_root_para_id(extended: &CT_CommentsEx, para_id: &str) -> String {
    let mut current = para_id.to_owned();
    let mut seen = HashSet::new();
    while seen.insert(current.clone()) {
        let Some(parent) = extended
            .comments
            .iter()
            .find(|entry| entry.para_id == current)
            .and_then(|entry| entry.para_id_parent.as_ref())
        else {
            break;
        };
        current.clone_from(parent);
    }
    current
}

fn allocate_para_id(
    comments: Option<&CT_Comments>,
    extended: Option<&CT_CommentsEx>,
) -> Result<String> {
    allocate_para_id_from_occupied(&mut occupied_para_ids(comments, extended))
}

fn occupied_para_ids(
    comments: Option<&CT_Comments>,
    extended: Option<&CT_CommentsEx>,
) -> HashSet<u32> {
    let mut occupied = comments
        .into_iter()
        .flat_map(|comments| comments.comments.iter())
        .flat_map(|comment| comment.paragraph_ids.iter())
        .filter_map(Option::as_deref)
        .filter_map(parse_para_id)
        .collect::<HashSet<_>>();
    occupied.extend(
        extended
            .into_iter()
            .flat_map(|extended| extended.comments.iter())
            .filter_map(|entry| parse_para_id(&entry.para_id)),
    );
    occupied
}

pub(crate) fn allocate_para_id_from_occupied(occupied: &mut HashSet<u32>) -> Result<String> {
    if let Some(max) = occupied.iter().copied().max()
        && max < u32::MAX
    {
        let allocated = max + 1;
        occupied.insert(allocated);
        return Ok(format!("{allocated:08X}"));
    }
    let allocated = (1..=u32::MAX)
        .find(|candidate| !occupied.contains(candidate))
        .ok_or_else(|| Error::Other("no available comment paragraph id remains".to_owned()))?;
    occupied.insert(allocated);
    Ok(format!("{allocated:08X}"))
}

fn parse_para_id(value: &str) -> Option<u32> {
    (value.len() == 8)
        .then(|| u32::from_str_radix(value, 16).ok())
        .flatten()
}

fn remove_owned_part(document: &mut Document, part: &str, relationship_type: &str) {
    document.package.remove_part(part);
    document.package.remove_part_rels(part);
    document.package.content_types.remove_override(part);
    document.identifiers.retire_authored_part(part);
    let owner = document.doc_part_name.clone();
    let mut removed_relationship_ids = Vec::new();
    if let Some(relationships) = document.package.get_part_rels_mut(&owner) {
        relationships.items.retain(|relationship| {
            let targets_part = relationship.rel_type == relationship_type
                && crate::document::relationship_is_internal(relationship)
                && OpcPackage::resolve_rel_target(&owner, &relationship.target) == part;
            let remove = targets_part
                && !document
                    .identifiers
                    .relationship_is_preserved(&owner, &relationship.id);
            if remove {
                removed_relationship_ids.push(relationship.id.clone());
            }
            !remove
        });
        if relationships.items.is_empty() {
            document.package.remove_part_rels(&owner);
        }
    }
    document
        .identifiers
        .retire_authored_story_relationships(&owner, removed_relationship_ids);
}

#[cfg(test)]
fn remove_anchors_from_table(table: &mut CT_Tbl, ids: &HashSet<i32>) {
    for (_, _, control) in &mut table.content_controls {
        remove_anchors_from_control(control, ids);
    }
    for row in &mut table.rows {
        remove_anchors_from_row(row, ids);
    }
}

#[cfg(test)]
fn remove_anchors_from_row(row: &mut CT_Row, ids: &HashSet<i32>) {
    for (_, _, control) in &mut row.content_controls {
        remove_anchors_from_control(control, ids);
    }
    for cell in &mut row.cells {
        remove_anchors_from_cell(cell, ids);
    }
}

#[cfg(test)]
fn remove_anchors_from_cell(cell: &mut CT_Tc, ids: &HashSet<i32>) {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => remove_anchors_from_paragraph(paragraph, ids),
            CellContent::Table(table) => remove_anchors_from_table(table, ids),
            CellContent::ContentControl(control) => remove_anchors_from_control(control, ids),
        }
    }
}

#[cfg(test)]
fn remove_anchors_from_control(control: &mut CT_Sdt, ids: &HashSet<i32>) {
    // Markers written inside `w:sdtContent` are preserved children there.
    control.remove_comment_anchors(&ids.iter().copied().collect::<Vec<_>>());
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => remove_anchors_from_paragraph(paragraph, ids),
            SdtContent::Table(table) => remove_anchors_from_table(table, ids),
            SdtContent::Row(row) => remove_anchors_from_row(row, ids),
            SdtContent::Cell(cell) => remove_anchors_from_cell(cell, ids),
            SdtContent::ContentControl(control) => remove_anchors_from_control(control, ids),
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
}

#[cfg(test)]
fn remove_anchors_from_paragraph(paragraph: &mut CT_P, ids: &HashSet<i32>) {
    for (_, _, _, control) in &mut paragraph.content_controls {
        remove_anchors_from_control(control, ids);
    }
    let ids = ids.iter().copied().collect::<Vec<_>>();
    paragraph.remove_comment_anchors(&ids);
}

// Source marker presence deliberately ignores accepted-view range pairing.
// A raw reference inside an opaque wrapper still owns its definition.
#[derive(Debug)]
pub(crate) struct CommentSourceMarker {
    pub(crate) id: i32,
    pub(crate) family: usize,
    pub(crate) span: Range<usize>,
    empty_run: Option<Range<usize>>,
}

struct CommentPartEntry {
    attributes: BTreeMap<String, String>,
    span: Range<usize>,
}

struct CommentOwnership {
    markers: BTreeMap<String, Vec<CommentSourceMarker>>,
    parents: BTreeMap<i32, Option<i32>>,
    entries: BTreeMap<String, Vec<(i32, Range<usize>)>>,
}

impl CommentOwnership {
    fn marker_counts(
        markers: &BTreeMap<String, Vec<CommentSourceMarker>>,
    ) -> BTreeMap<i32, [usize; 3]> {
        let mut counts = BTreeMap::new();
        for marker in markers.values().flatten() {
            counts.entry(marker.id).or_insert([0; 3])[marker.family] += 1;
        }
        counts
    }

    fn descendants(&self, id: i32) -> HashSet<i32> {
        let mut ids = HashSet::from([id]);
        loop {
            let before = ids.len();
            for (&child, parent) in &self.parents {
                if parent.is_some_and(|parent| ids.contains(&parent)) {
                    ids.insert(child);
                }
            }
            if ids.len() == before {
                return ids;
            }
        }
    }
}

pub(crate) fn comment_source_markers(xml: &[u8]) -> Result<Vec<CommentSourceMarker>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut markers = Vec::new();
    let mut preceding_run: Option<(usize, usize, Vec<u8>)> = None;
    loop {
        let start = reader.buffer_position() as usize;
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                let preceding = preceding_run.take();
                let end = reader.buffer_position() as usize;
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                {
                    if local.as_ref() == b"r"
                        && xml.get(end.saturating_sub(2)..end) != Some(b"/>")
                        && element.attributes().all(|attribute| {
                            attribute.is_ok_and(|attribute| {
                                attribute.key.as_ref() == b"xmlns"
                                    || attribute.key.as_ref().starts_with(b"xmlns:")
                            })
                        })
                    {
                        preceding_run = Some((start, end, element.name().as_ref().to_vec()));
                    }
                    let family = match local.as_ref() {
                        b"commentRangeStart" => Some(0),
                        b"commentRangeEnd" => Some(1),
                        b"commentReference" => Some(2),
                        _ => None,
                    };
                    if let Some(family) = family {
                        let mut ids = 0;
                        for attribute in element.attributes() {
                            let attribute =
                                attribute.map_err(|error| Error::Other(error.to_string()))?;
                            let (namespace, local) =
                                reader.resolver().resolve_attribute(attribute.key);
                            if local.as_ref() == b"id"
                                && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                            {
                                ids += 1;
                            }
                        }
                        if ids != 1 {
                            return Err(Error::Other(
                                "comment marker has missing or ambiguous qualified id".into(),
                            ));
                        }
                        let id = word_marker_attribute(&reader, &element, b"id")?
                            .and_then(|id| id.parse::<i32>().ok())
                            .ok_or_else(|| {
                                Error::Other("comment marker has an invalid qualified id".into())
                            })?;
                        if xml.get(end.saturating_sub(2)..end) != Some(b"/>") {
                            reader
                                .read_to_end_into(element.name(), &mut Vec::new())
                                .map_err(|error| Error::Other(error.to_string()))?;
                        }
                        let marker_end = reader.buffer_position() as usize;
                        let empty_run = preceding.and_then(|(run_start, run_end, name)| {
                            let mut close = b"</".to_vec();
                            close.extend(name);
                            close.push(b'>');
                            (family == 2
                                && run_end == start
                                && xml.get(marker_end..marker_end + close.len())
                                    == Some(close.as_slice()))
                            .then_some(run_start..marker_end + close.len())
                        });
                        markers.push(CommentSourceMarker {
                            id,
                            family,
                            span: start..marker_end,
                            empty_run,
                        });
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(markers)
}

/// Scan only qualified direct entries, retaining exact source spans for removal.
fn comment_part_entries(
    xml: &[u8],
    namespace: &str,
    root: &[u8],
    item: &[u8],
) -> Result<Vec<CommentPartEntry>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut entries = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let empty = xml.get(end.saturating_sub(2)..end) == Some(b"/>");
                let (resolved, local) = reader.resolver().resolve_element(element.name());
                let qualified = matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace.as_bytes());
                if depth == 0 {
                    if saw_root || !qualified || local.as_ref() != root {
                        return Err(Error::Other(
                            "comment part has no unique qualified root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 1 && qualified && local.as_ref() == item {
                    let mut attributes = BTreeMap::new();
                    for attribute in element.attributes() {
                        let attribute =
                            attribute.map_err(|error| Error::Other(error.to_string()))?;
                        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
                        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace.as_bytes())
                        {
                            let key = String::from_utf8_lossy(local.as_ref()).into_owned();
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| Error::Other(error.to_string()))?
                                .into_owned();
                            if attributes.insert(key, value).is_some() {
                                return Err(Error::Other(
                                    "comment part has duplicate qualified attributes".into(),
                                ));
                            }
                        }
                    }
                    if !empty {
                        reader
                            .read_to_end_into(element.name(), &mut Vec::new())
                            .map_err(|error| Error::Other(error.to_string()))?;
                    }
                    entries.push(CommentPartEntry {
                        attributes,
                        span: start..reader.buffer_position() as usize,
                    });
                    buffer.clear();
                    continue;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Other("comment XML is unbalanced".into()))?;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !saw_root || depth != 0 {
        return Err(Error::Other("comment XML has no complete root".into()));
    }
    Ok(entries)
}

impl Document {
    /// Return a selected comment's exact accepted-view run range.
    /// Unknown or malformed ownership and unrepresentable endpoints are errors.
    /// Known orphans and reference-only points have no paired range.
    pub fn comment_anchor(&self, id: i32) -> Result<Option<StoryRunRange>> {
        Ok(self
            .comment_anchor_snapshots_selected(Some(id))?
            .remove(&id)
            .ok_or_else(|| Error::Other(format!("unknown comment {id}")))?
            .0)
    }

    /// Return accepted span text, empty for a reference-only point and absent
    /// for a known orphan or a reply without its own source markers.
    pub fn comment_anchor_text(&self, id: i32) -> Result<Option<String>> {
        Ok(self
            .comment_anchor_snapshots_selected(Some(id))?
            .remove(&id)
            .ok_or_else(|| Error::Other(format!("unknown comment {id}")))?
            .1)
    }

    /// Build checked owned anchor snapshots with one source inventory for listings.
    #[doc(hidden)]
    #[allow(clippy::type_complexity)] // Concrete shared range and text listing contract.
    pub fn comment_anchor_snapshots(
        &self,
    ) -> Result<BTreeMap<i32, (Option<StoryRunRange>, Option<String>)>> {
        self.comment_anchor_snapshots_selected(None)
    }

    #[allow(clippy::type_complexity)] // Same concrete payload as the public batch accessor.
    fn comment_anchor_snapshots_selected(
        &self,
        selected: Option<i32>,
    ) -> Result<BTreeMap<i32, (Option<StoryRunRange>, Option<String>)>> {
        let mut source = self.clone_for_staging();
        source.flush_to_package()?;
        let ownership = source.comment_owned_graph_at(&source.doc_part_name)?;
        if let Some(id) = selected
            && !ownership.parents.contains_key(&id)
        {
            return Err(Error::Other(format!("unknown comment {id}")));
        }
        let counts = CommentOwnership::marker_counts(&ownership.markers);
        let paragraphs = source
            .comment_story_range_paragraphs()?
            .into_iter()
            .map(|(location, xml)| {
                let ids = comment_source_markers(&xml)?
                    .into_iter()
                    .map(|marker| marker.id)
                    .collect::<HashSet<_>>();
                Ok((location, CT_P::from_xml_fragment(&xml)?, ids, xml))
            })
            .collect::<Result<Vec<_>>>()?;
        let mut marker_paragraphs = HashMap::<i32, Vec<usize>>::new();
        let mut owner_paragraphs = HashMap::<crate::StoryId, Vec<usize>>::new();
        for (index, (location, _, ids, _)) in paragraphs.iter().enumerate() {
            owner_paragraphs
                .entry(location.story().clone())
                .or_default()
                .push(index);
            for &id in ids {
                marker_paragraphs.entry(id).or_default().push(index);
            }
        }
        let mut snapshots = BTreeMap::new();
        for &id in ownership
            .parents
            .keys()
            .filter(|id| selected.is_none_or(|selected| selected == **id))
        {
            let count = counts.get(&id).copied().unwrap_or_default();
            if count == [0; 3] {
                snapshots.insert(id, (None, None));
                continue;
            }
            if count == [0, 0, 1] {
                snapshots.insert(id, (None, Some(String::new())));
                continue;
            }
            if count[0] != 1 || count[1] != 1 || count[2] > 1 {
                return Err(Error::Other(format!(
                    "comment {id} has duplicate or unmatched source markers"
                )));
            }
            let mut start = None::<StoryRunPosition>;
            let mut end = None::<StoryRunPosition>;
            let mut open = false;
            let mut text = String::new();
            let mut first_paragraph = None;
            for &index in marker_paragraphs.get(&id).into_iter().flatten() {
                let (_, _, _, xml) = &paragraphs[index];
                let (_, _, first, last) = CT_P::accepted_comment_source_projection(xml, id, false)?;
                if first.is_some() || last.is_some() {
                    first_paragraph = Some(index);
                    break;
                }
            }
            let first_paragraph = first_paragraph.ok_or_else(|| Error::Other(format!(
                "comment {id} source cannot be projected faithfully on the existing accepted run axis"
            )))?;
            let owner = &owner_paragraphs[paragraphs[first_paragraph].0.story()];
            let offset = owner
                .binary_search(&first_paragraph)
                .expect("registered owner paragraph");
            for &index in &owner[offset..] {
                let (location, paragraph, ids, xml) = &paragraphs[index];
                let was_open = open;
                let (contribution, next_open, first, last) = if ids.contains(&id) {
                    CT_P::accepted_comment_source_projection(xml, id, open)?
                } else {
                    (paragraph.accepted_text(), open, None, None)
                };
                if let Some(index) = first {
                    if start.is_some() {
                        return Err(Error::Other(format!(
                            "comment {id} has duplicate projected starts"
                        )));
                    }
                    start = Some(StoryRunPosition {
                        location: location.clone(),
                        run_index: index,
                    });
                }
                if was_open || first.is_some() {
                    if was_open {
                        text.push('\n');
                    }
                    text.push_str(&contribution);
                }
                if let Some(index) = last {
                    if end.is_some() {
                        return Err(Error::Other(format!(
                            "comment {id} has duplicate projected ends"
                        )));
                    }
                    end = Some(StoryRunPosition {
                        location: location.clone(),
                        run_index: index,
                    });
                }
                open = next_open;
                if end.is_some() {
                    break;
                }
            }
            let (Some(start), Some(end)) = (start, end) else {
                return Err(Error::Other(format!(
                    "comment {id} source cannot be projected faithfully on the existing accepted run axis"
                )));
            };
            if open || start.location.story() != end.location.story() {
                return Err(Error::Other(format!(
                    "comment {id} range has no ordered end in the same story owner"
                )));
            }
            snapshots.insert(id, (Some(StoryRunRange { start, end }), Some(text)));
        }
        Ok(snapshots)
    }

    /// Classify glossary review ownership from one actual internal relationship.
    pub(crate) fn glossary_comment_owner(&self) -> Result<Option<String>> {
        let Some(owner) = self.glossary_part_name.as_deref() else {
            return Ok(None);
        };
        let Some(_part) =
            self.comment_relationship_part_at(owner, oxml_opc::relationship::rel_types::COMMENTS)?
        else {
            if self.package.get_part_rels(owner).iter().flat_map(|rels| &rels.items).any(|relationship| {
                matches!(relationship.rel_type.as_str(), COMMENTS_EXTENDED_REL_TYPE
                    | "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
                    | "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible")
            }) {
                return Err(Error::Other("glossary comment companions have no owned definitions relationship".into()));
            }
            return Ok(None);
        };
        self.comment_definition_entries_at(owner)?;
        self.comment_owned_graph_at(owner)?;
        // Two review graphs cannot claim the same physical part.
        for local in self
            .package
            .get_part_rels(owner)
            .iter()
            .flat_map(|rels| &rels.items)
        {
            if !matches!(
                local.rel_type.as_str(),
                oxml_opc::relationship::rel_types::COMMENTS
                    | COMMENTS_EXTENDED_REL_TYPE
                    | "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
                    | "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
            ) {
                continue;
            }
            if !crate::document::relationship_is_internal(local) {
                return Err(Error::Other(
                    "glossary comment companion is external".into(),
                ));
            }
            let target = OpcPackage::resolve_rel_target(owner, &local.target);
            if self
                .package
                .get_part_rels(&self.doc_part_name)
                .iter()
                .flat_map(|rels| &rels.items)
                .any(|main| {
                    crate::document::relationship_is_internal(main)
                        && crate::document::part_name_identity(&OpcPackage::resolve_rel_target(
                            &self.doc_part_name,
                            &main.target,
                        )) == crate::document::part_name_identity(&target)
                })
            {
                return Err(Error::Other(
                    "glossary and main comment owners share a dependency target".into(),
                ));
            }
        }
        // Shared unannotated notes are harmless. A marked physical source,
        // however, cannot belong to both independent review graphs.
        let main_sources = self.word_story_part_names();
        for local in self
            .package
            .get_part_rels(owner)
            .iter()
            .flat_map(|rels| &rels.items)
        {
            if !crate::document::relationship_is_internal(local)
                || !matches!(
                    local.rel_type.as_str(),
                    oxml_opc::relationship::rel_types::FOOTNOTES
                        | oxml_opc::relationship::rel_types::ENDNOTES
                )
            {
                continue;
            }
            let target = OpcPackage::resolve_rel_target(owner, &local.target);
            let identity = crate::document::part_name_identity(&target);
            if main_sources.iter().any(|part| {
                crate::document::part_name_identity(part) == identity
                    && crate::document::part_name_identity(part)
                        != crate::document::part_name_identity(owner)
            }) && let Some(xml) = self.package.get_part(&target)
                && !comment_source_markers(xml)?.is_empty()
            {
                return Err(Error::Other(
                    "glossary and main review owners share a marked physical source".into(),
                ));
            }
        }
        Ok(Some(owner.to_owned()))
    }

    /// Omit only qualified marker spans, retaining mixed runs and opaque siblings.
    pub(crate) fn omit_comment_markers(xml: &[u8]) -> Result<Vec<u8>> {
        let markers = comment_source_markers(xml)?;
        let mut result = xml.to_vec();
        for marker in markers.into_iter().rev() {
            let mut reader = quick_xml::Reader::from_reader(&xml[marker.span.clone()]);
            if matches!(
                reader
                    .read_event()
                    .map_err(|error| Error::Other(error.to_string()))?,
                Event::Start(_)
            ) && !matches!(
                reader
                    .read_event()
                    .map_err(|error| Error::Other(error.to_string()))?,
                Event::End(_)
            ) {
                return Err(Error::Other(format!(
                    "cannot omit comment {} with opaque marker content",
                    marker.id
                )));
            }
            result.drain(marker.span);
        }
        Ok(result)
    }

    fn comment_relationship_part_at(
        &self,
        owner: &str,
        relationship_type: &str,
    ) -> Result<Option<String>> {
        let mut target = None;
        for relationship in self
            .package
            .get_part_rels(owner)
            .iter()
            .flat_map(|rels| &rels.items)
            .filter(|relationship| relationship.rel_type == relationship_type)
        {
            if !crate::document::relationship_is_internal(relationship) || target.is_some() {
                return Err(Error::Other(
                    "comment part ownership is external or ambiguous".into(),
                ));
            }
            let part = OpcPackage::resolve_rel_target(owner, &relationship.target);
            if self.package.get_part(&part).is_none() {
                return Err(Error::Other(format!(
                    "comment relationship targets missing part {part}"
                )));
            }
            target = Some(part);
        }
        Ok(target)
    }

    fn comment_source_inventory_at(
        &self,
        owner: &str,
    ) -> Result<BTreeMap<String, Vec<CommentSourceMarker>>> {
        if self.package.get_part(&self.doc_part_name).is_none() {
            return Err(Error::Other("main comment source part is missing".into()));
        }
        let mut markers = BTreeMap::new();
        let local = if owner == self.doc_part_name {
            self.glossary_comment_owner()?
        } else {
            None
        };
        let parts = if owner == self.doc_part_name {
            self.word_story_part_names()
                .into_iter()
                .filter(|part| local.as_deref() != Some(part.as_str()))
                .collect::<Vec<_>>()
        } else {
            let mut parts = vec![owner.to_owned()];
            for relationship in self
                .package
                .get_part_rels(owner)
                .iter()
                .flat_map(|rels| &rels.items)
            {
                if crate::document::relationship_is_internal(relationship)
                    && matches!(
                        relationship.rel_type.as_str(),
                        oxml_opc::relationship::rel_types::COMMENTS
                            | oxml_opc::relationship::rel_types::FOOTNOTES
                            | oxml_opc::relationship::rel_types::ENDNOTES
                    )
                {
                    parts.push(OpcPackage::resolve_rel_target(owner, &relationship.target));
                }
            }
            parts
        };
        for part in parts {
            // Dangling story relationships contain no source to inventory.
            // Package validation diagnoses them independently of comment ownership.
            if let Some(xml) = self.package.get_part(&part) {
                markers.insert(part, comment_source_markers(xml)?);
            }
        }
        Ok(markers)
    }

    fn comment_definition_entries_at(
        &self,
        owner: &str,
    ) -> Result<Option<(String, Vec<CommentPartEntry>)>> {
        let Some(part) =
            self.comment_relationship_part_at(owner, oxml_opc::relationship::rel_types::COMMENTS)?
        else {
            if owner == self.doc_part_name && self.comments_part_name.is_some() {
                return Err(Error::Other(
                    "owned comment definition relationship is missing".into(),
                ));
            }
            return Ok(None);
        };
        let entries = comment_part_entries(
            self.package.get_part(&part).expect("checked part"),
            rdocx_oxml::namespace::W_NS,
            b"comments",
            b"comment",
        )?;
        let mut ids = HashSet::new();
        for entry in &entries {
            let id = entry
                .attributes
                .get("id")
                .and_then(|id| id.parse::<i32>().ok())
                .ok_or_else(|| Error::Other("comment definition has invalid id".into()))?;
            if !ids.insert(id) {
                return Err(Error::Other(format!(
                    "comment {id} has duplicate definitions"
                )));
            }
        }
        Ok(Some((part, entries)))
    }

    fn comment_owned_graph_at(&self, owner: &str) -> Result<CommentOwnership> {
        let mut ownership = CommentOwnership {
            markers: self.comment_source_inventory_at(owner)?,
            parents: BTreeMap::new(),
            entries: BTreeMap::new(),
        };
        let source_definitions = self.comment_definition_entries_at(owner)?;
        let comments = if let Some((part, _)) = &source_definitions {
            CT_Comments::from_xml(self.package.get_part(part).expect("checked part"))?
        } else {
            CT_Comments::new()
        };
        let mut by_para = BTreeMap::new();
        let mut by_parent_para = BTreeMap::new();
        let mut definitions = Vec::new();
        for entry in source_definitions
            .as_ref()
            .into_iter()
            .flat_map(|(_, entries)| entries)
        {
            let id = entry
                .attributes
                .get("id")
                .and_then(|id| id.parse::<i32>().ok())
                .ok_or_else(|| Error::Other("comment definition has invalid id".into()))?;
            if ownership.parents.insert(id, None).is_some() {
                return Err(Error::Other(format!(
                    "comment {id} has duplicate definitions"
                )));
            }
            definitions.push((id, entry.span.clone()));
        }
        if comments.comments.len() != definitions.len() {
            return Err(Error::Other(
                "comment definitions cannot be qualified without losing opaque ownership".into(),
            ));
        }
        for comment in &comments.comments {
            for para in para_ids(comment) {
                if by_parent_para
                    .insert(para.to_ascii_uppercase(), comment.id)
                    .is_some()
                {
                    return Err(Error::Other(
                        "comment paragraph identity is ambiguous".into(),
                    ));
                }
            }
            if !ownership.parents.contains_key(&comment.id) {
                return Err(Error::Other(
                    "comment model differs from qualified source definitions".into(),
                ));
            }
            if let Some(para) = last_para_id(comment)
                && by_para
                    .insert(para.to_ascii_uppercase(), comment.id)
                    .is_some()
            {
                return Err(Error::Other(
                    "comment last-paragraph identity is ambiguous".into(),
                ));
            }
        }
        if let Some((part, _)) = source_definitions {
            ownership.entries.insert(part, definitions);
        }
        const IDS_REL: &str =
            "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
        const EXTENSIBLE_REL: &str =
            "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
        const IDS_NS: &str = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
        const CEX_NS: &str = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
        let mut by_durable = BTreeMap::new();
        for (relationship, namespace, root, item, key) in [
            (
                COMMENTS_EXTENDED_REL_TYPE,
                rdocx_oxml::comments_extended::W15_NS,
                b"commentsEx".as_slice(),
                b"commentEx".as_slice(),
                "paraId",
            ),
            (
                IDS_REL,
                IDS_NS,
                b"commentsIds".as_slice(),
                b"commentId".as_slice(),
                "paraId",
            ),
            (
                EXTENSIBLE_REL,
                CEX_NS,
                b"commentsExtensible".as_slice(),
                b"commentExtensible".as_slice(),
                "durableId",
            ),
        ] {
            let Some(part) = self.comment_relationship_part_at(owner, relationship)? else {
                continue;
            };
            let xml = self.package.get_part(&part).expect("checked part");
            let mut mapped = Vec::new();
            let mut seen = HashSet::new();
            for entry in comment_part_entries(xml, namespace, root, item)? {
                let value = entry
                    .attributes
                    .get(key)
                    .ok_or_else(|| Error::Other(format!("comment companion has no {key}")))?
                    .to_ascii_uppercase();
                if !seen.insert(value.clone()) {
                    return Err(Error::Other(format!(
                        "comment companion {key} {value} is ambiguous"
                    )));
                }
                let id = if key == "paraId" {
                    by_para.get(&value)
                } else {
                    by_durable.get(&value)
                }
                .copied()
                .ok_or_else(|| {
                    Error::Other(format!(
                        "comment companion {key} {value} has unprovable linkage"
                    ))
                })?;
                if relationship == COMMENTS_EXTENDED_REL_TYPE
                    && let Some(parent) = entry.attributes.get("paraIdParent")
                {
                    let parent = by_parent_para
                        .get(&parent.to_ascii_uppercase())
                        .copied()
                        .ok_or_else(|| {
                            Error::Other(format!("comment {id} has unknown parent {parent}"))
                        })?;
                    ownership.parents.insert(id, Some(parent));
                }
                if relationship == IDS_REL {
                    let durable = entry
                        .attributes
                        .get("durableId")
                        .ok_or_else(|| {
                            Error::Other(format!("comment {id} has no durable identity"))
                        })?
                        .to_ascii_uppercase();
                    if durable.len() != 8
                        || u32::from_str_radix(&durable, 16).is_err()
                        || by_durable.insert(durable, id).is_some()
                    {
                        return Err(Error::Other(format!(
                            "comment {id} has invalid or duplicate durable identity"
                        )));
                    }
                }
                mapped.push((id, entry.span));
            }
            ownership.entries.insert(part, mapped);
        }
        for &id in ownership.parents.keys() {
            let mut seen = HashSet::new();
            let mut ancestor = Some(id);
            while let Some(current) = ancestor {
                if !seen.insert(current) {
                    return Err(Error::Other(format!(
                        "comment {id} has cyclic thread ownership"
                    )));
                }
                ancestor = ownership.parents[&current];
            }
        }
        Ok(ownership)
    }

    fn comment_ownership_at(&self, owner: &str) -> Result<CommentOwnership> {
        let ownership = self.comment_owned_graph_at(owner)?;
        let counts = CommentOwnership::marker_counts(&ownership.markers);
        if !counts.is_empty()
            && self
                .comment_relationship_part_at(owner, oxml_opc::relationship::rel_types::COMMENTS)?
                .is_none()
        {
            return Err(Error::Other(
                "comment markers have no owned definitions part".into(),
            ));
        }
        for id in counts.keys() {
            if !ownership.parents.contains_key(id) {
                return Err(Error::Other(format!(
                    "comment {id} has markers but no definition"
                )));
            }
        }
        Ok(ownership)
    }

    /// Check raw comment presence and thread ownership across relationship-resolved stories.
    #[doc(hidden)]
    pub fn validate_comment_ownership(&self) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let mut owners = vec![candidate.doc_part_name.clone()];
        owners.extend(candidate.glossary_comment_owner()?);
        for owner in owners {
            let ownership = candidate.comment_ownership_at(&owner)?;
            let counts = CommentOwnership::marker_counts(&ownership.markers);
            for (&id, parent) in &ownership.parents {
                let count = counts.get(&id).copied().unwrap_or_default();
                if count.iter().any(|count| *count > 1) || count[0] != count[1] {
                    return Err(Error::Other(format!(
                        "comment {id} has duplicate or unmatched source markers"
                    )));
                }
                if parent.is_none() && count == [0; 3] {
                    return Err(Error::Other(format!(
                        "comment {id} is an orphan root with no source range or reference"
                    )));
                }
            }
        }
        Ok(())
    }

    /// Compare one unpublished destructive edit with its frozen source inventory.
    pub(crate) fn reconcile_comment_removal(&mut self, before: &Document) -> Result<()> {
        let mut owners = vec![before.doc_part_name.clone()];
        let local = before.glossary_comment_owner()?;
        if local != self.glossary_comment_owner()? {
            return Err(Error::Other(
                "glossary comment relationship owner changed during edit".into(),
            ));
        }
        owners.extend(local);
        for owner in owners {
            self.reconcile_comment_removal_at(before, &owner)?;
        }
        Ok(())
    }

    fn reconcile_comment_removal_at(&mut self, before: &Document, owner: &str) -> Result<()> {
        let mut source = before.clone_for_staging();
        source.flush_to_package()?;
        self.flush_to_package()?;
        let before_counts =
            CommentOwnership::marker_counts(&source.comment_source_inventory_at(owner)?);
        let after_counts =
            CommentOwnership::marker_counts(&self.comment_source_inventory_at(owner)?);
        // Only source ownership decreases need deletion proof. Unchanged
        // or cloned anchors must not validate or repair unrelated producer
        // definitions. Explicit deletion and CLI validation remain strict.
        if before_counts.iter().all(|(id, count)| {
            let after = after_counts.get(id).copied().unwrap_or_default();
            count
                .iter()
                .zip(after)
                .all(|(before, after)| after >= *before)
        }) {
            return Ok(());
        }
        // A demonstrably undefined source marker cannot orphan an absent
        // definition. Retain its established editing/comparison lifecycle.
        // Raw qualified entries prove absence even if typed projection omits one.
        let definitions = source.comment_definition_entries_at(owner)?;
        let defined_decrease = definitions.as_ref().is_some_and(|(_, entries)| {
            entries.iter().any(|entry| {
                let id = entry.attributes["id"]
                    .parse::<i32>()
                    .expect("validated definition id");
                let before = before_counts.get(&id).copied().unwrap_or_default();
                let after = after_counts.get(&id).copied().unwrap_or_default();
                before
                    .iter()
                    .zip(after)
                    .any(|(before, after)| after < *before)
            })
        });
        if !defined_decrease {
            return Ok(());
        }
        let original = source.comment_ownership_at(owner)?;
        let remaining = self.comment_ownership_at(owner)?;
        let before_counts = CommentOwnership::marker_counts(&original.markers);
        let after_counts = CommentOwnership::marker_counts(&remaining.markers);
        let mut removed = HashSet::new();
        for (&id, count) in &before_counts {
            let after = after_counts.get(&id).copied().unwrap_or_default();
            if count
                .iter()
                .zip(after)
                .all(|(before, after)| after >= *before)
            {
                continue;
            }
            if after != [0; 3] || count.iter().any(|count| *count > 1) || count[0] != count[1] {
                return Err(Error::Other(format!(
                    "cannot remove part of comment {id}; its source graph survives or is ambiguous"
                )));
            }
            removed.extend(original.descendants(id));
        }
        for id in &removed {
            if after_counts.get(id).is_some_and(|counts| *counts != [0; 3]) {
                return Err(Error::Other(format!(
                    "cannot remove comment {id}; a descendant source anchor survives"
                )));
            }
        }
        if !removed.is_empty() {
            self.remove_comment_ids_staged_at(owner, &removed)?;
        }
        Ok(())
    }

    pub(crate) fn refuse_commented_fragment(xml: &[u8]) -> Result<()> {
        if let Some(marker) = comment_source_markers(xml)?.first() {
            return Err(Error::Other(format!(
                "cannot detach content bearing comment {}; fragments do not own comment threads",
                marker.id
            )));
        }
        Ok(())
    }

    fn remove_comment_ids_staged_at(&mut self, owner: &str, ids: &HashSet<i32>) -> Result<()> {
        let ownership = self.comment_ownership_at(owner)?;
        for (part, entries) in &ownership.entries {
            for (_, span) in entries.iter().filter(|(id, _)| ids.contains(id)) {
                if let Some(marker) =
                    ownership
                        .markers
                        .get(part)
                        .into_iter()
                        .flatten()
                        .find(|marker| {
                            !ids.contains(&marker.id)
                                && span.start <= marker.span.start
                                && marker.span.end <= span.end
                        })
                {
                    return Err(Error::Other(format!(
                        "cannot remove a comment definition carrying unrelated comment {}",
                        marker.id
                    )));
                }
            }
        }
        let mut spans = BTreeMap::<String, Vec<Range<usize>>>::new();
        for (part, markers) in ownership.markers {
            for marker in markers
                .into_iter()
                .filter(|marker| ids.contains(&marker.id))
            {
                spans
                    .entry(part.clone())
                    .or_default()
                    .push(marker.empty_run.unwrap_or(marker.span));
            }
        }
        for (part, entries) in ownership.entries {
            for (id, span) in entries.into_iter().filter(|(id, _)| ids.contains(id)) {
                let _ = id;
                spans.entry(part.clone()).or_default().push(span);
            }
        }
        for (part, mut ranges) in spans {
            ranges.sort_by_key(|range| range.start);
            ranges.dedup();
            // A removed definition can contain a source marker of another removed comment.
            let mut outer = Vec::<Range<usize>>::new();
            for range in ranges {
                if let Some(previous) = outer.last() {
                    if range.end <= previous.end {
                        continue;
                    }
                    if range.start < previous.end {
                        return Err(Error::Other(
                            "comment removal spans overlap ambiguously".into(),
                        ));
                    }
                }
                outer.push(range);
            }
            let mut xml = self
                .package
                .get_part(&part)
                .ok_or_else(|| Error::Other("comment removal part disappeared".into()))?
                .to_vec();
            for range in outer.into_iter().rev() {
                xml.drain(range);
            }
            if self.comments_extended_part_name.as_deref() == Some(part.as_str()) {
                self.comments_extended = Some(CT_CommentsEx::from_xml(&xml)?);
            }
            crate::document::set_story_source_xml(self, &part, xml)?;
        }
        if owner == self.doc_part_name {
            self.identifiers
                .retire_authored_comment_ids(ids.iter().copied());
            self.remove_owned_empty_comment_parts();
            self.comments_dirty = false;
        }
        self.invalidate_layout();
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::{Path, PathBuf};
    use std::process::Command;

    use super::*;

    const WORD_VERSION: &str = "16.104";
    const WORD_BUILD: &str = "16.104.25121423";
    const WORD_COMMENT_CANDIDATE_SHA256: &str =
        "b7e1f39a5af80d9928ed671fa45557d485c2c70c8761439a09df66c027274995";

    fn word_comment_candidate() -> Document {
        let mut document =
            Document::new_with_profile(crate::document::WordCreationProfile::Minimal(
                crate::document::WordPackageClass::Document,
            ));
        let mut paragraph = document.add_paragraph("");
        paragraph.add_run("Review ");
        paragraph.add_run("this sentence.");
        let root = document
            .add_comment(
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: RunPosition {
                        body_index: 0,
                        run_index: 2,
                    },
                },
                "Ada Lovelace",
                Some("AL"),
                "Please verify this sentence.",
            )
            .expect("add candidate comment");
        document
            .reply_to(root, "Ben", "Verified and ready.")
            .expect("add candidate reply");
        assert!(
            document
                .resolve_comment(root, true)
                .expect("resolve candidate thread")
        );
        document
    }

    #[test]
    fn comment_reference_is_inserted_at_the_half_open_end() {
        let mut document = Document::new();
        let mut paragraph = document.add_paragraph("");
        paragraph.add_run("left");
        paragraph.add_run("right");
        let id = document
            .add_comment(
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
                "Ada",
                None,
                "Review",
            )
            .unwrap();

        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            panic!("body item should remain a paragraph");
        };
        assert!(matches!(
            paragraph.runs[1].content.as_slice(),
            [RunContent::CommentReference { id: reference, .. }] if *reference == id
        ));
        assert!(paragraph.comment_ranges.iter().any(|marker| matches!(
            marker,
            CommentRangeMarker::End {
                id: marker_id,
                run_index: 1,
                ..
            } if *marker_id == id
        )));
    }

    #[test]
    fn adding_a_comment_reserves_both_relationships_before_publication() {
        let mut source = Document::new_with_profile(crate::document::WordCreationProfile::Minimal(
            crate::document::WordPackageClass::Document,
        ));
        source.add_paragraph("review");
        let source_bytes = source.to_bytes().unwrap();
        let mut package =
            oxml_opc::OpcPackage::from_reader(std::io::Cursor::new(source_bytes)).unwrap();
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        let mut document = Document::from_bytes(bytes.get_ref()).unwrap();
        let package_bytes = |document: &Document| {
            let mut bytes = std::io::Cursor::new(Vec::new());
            document.package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        };
        let before = package_bytes(&document);

        let error = document
            .add_comment(
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
                "Ada",
                None,
                "Review",
            )
            .unwrap_err();

        assert!(
            error
                .to_string()
                .contains("comments relationship allocation failed")
        );
        assert_eq!(package_bytes(&document), before);
        assert!(document.comments.is_none());
        assert!(document.comments_extended.is_none());
        assert!(document.comments_part_name.is_none());
        assert!(document.comments_extended_part_name.is_none());
    }

    #[test]
    fn reply_and_resolve_roll_back_comments_extended_relationship_exhaustion() {
        fn exhausted_document() -> Document {
            let mut source =
                Document::new_with_profile(crate::document::WordCreationProfile::Minimal(
                    crate::document::WordPackageClass::Document,
                ));
            source.add_paragraph("review");
            source
                .add_comment(
                    RunRange {
                        start: RunPosition {
                            body_index: 0,
                            run_index: 0,
                        },
                        end: RunPosition {
                            body_index: 0,
                            run_index: 1,
                        },
                    },
                    "Ada",
                    None,
                    "Review",
                )
                .unwrap();
            let mut package =
                OpcPackage::from_reader(std::io::Cursor::new(source.to_bytes().unwrap())).unwrap();
            package.parts.remove(DEFAULT_COMMENTS_EXTENDED_PART);
            package
                .content_types
                .overrides
                .remove(DEFAULT_COMMENTS_EXTENDED_PART);
            let relationships = package.get_or_create_part_rels("/word/document.xml");
            relationships
                .items
                .retain(|relationship| relationship.rel_type != COMMENTS_EXTENDED_REL_TYPE);
            relationships.add_with_id(
                &format!("rId{}", u32::MAX),
                "urn:exhaustion",
                "unchanged.bin",
            );
            let mut bytes = std::io::Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        let package_bytes = |document: &Document| {
            let mut bytes = std::io::Cursor::new(Vec::new());
            document.package.write_to(&mut bytes).unwrap();
            bytes.into_inner()
        };

        let mut reply = exhausted_document();
        let root = reply.comments()[0].id();
        let before = package_bytes(&reply);
        let error = reply.reply_to(root, "Ben", "Done").unwrap_err();
        assert!(
            error
                .to_string()
                .contains("comments-extended relationship allocation failed")
        );
        assert_eq!(package_bytes(&reply), before);
        assert!(reply.comments_extended.is_none());
        assert!(reply.comments_extended_part_name.is_none());
        assert_eq!(reply.comments().len(), 1);

        let mut resolved = exhausted_document();
        let root = resolved.comments()[0].id();
        let before = package_bytes(&resolved);
        let error = resolved.resolve_comment(root, true).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("comments-extended relationship allocation failed")
        );
        assert_eq!(package_bytes(&resolved), before);
        assert!(resolved.comments_extended.is_none());
        assert!(resolved.comments_extended_part_name.is_none());
        assert!(!resolved.comments()[0].resolved());
    }

    #[test]
    fn successful_legacy_comment_upgrades_keep_reserved_relationships_and_parts() {
        fn legacy_document() -> Document {
            let mut source =
                Document::new_with_profile(crate::document::WordCreationProfile::Minimal(
                    crate::document::WordPackageClass::Document,
                ));
            source.add_paragraph("review");
            source
                .add_comment(
                    RunRange {
                        start: RunPosition {
                            body_index: 0,
                            run_index: 0,
                        },
                        end: RunPosition {
                            body_index: 0,
                            run_index: 1,
                        },
                    },
                    "Ada",
                    None,
                    "Review",
                )
                .unwrap();
            let mut package =
                OpcPackage::from_reader(std::io::Cursor::new(source.to_bytes().unwrap())).unwrap();
            package.parts.remove(DEFAULT_COMMENTS_EXTENDED_PART);
            package
                .content_types
                .overrides
                .remove(DEFAULT_COMMENTS_EXTENDED_PART);
            package
                .get_or_create_part_rels("/word/document.xml")
                .items
                .retain(|relationship| relationship.rel_type != COMMENTS_EXTENDED_REL_TYPE);
            let mut bytes = std::io::Cursor::new(Vec::new());
            package.write_to(&mut bytes).unwrap();
            Document::from_bytes(bytes.get_ref()).unwrap()
        }

        for reply in [false, true] {
            let mut document = legacy_document();
            let root = document.comments()[0].id();
            if reply {
                document.reply_to(root, "Ben", "Done").unwrap();
            } else {
                assert!(document.resolve_comment(root, true).unwrap());
            }
            let extended_part = document.comments_extended_part_name.clone().unwrap();
            let extended_relationship = document
                .package
                .get_part_rels("/word/document.xml")
                .and_then(|relationships| relationships.get_by_type(COMMENTS_EXTENDED_REL_TYPE))
                .unwrap()
                .id
                .clone();

            let hyperlink = document.add_hyperlink_relationship("https://example.com");
            let imported_part = document
                .identifiers
                .reserve_fragment_part_name(&extended_part)
                .unwrap();

            assert_ne!(hyperlink, extended_relationship);
            assert_ne!(imported_part, extended_part);
        }
    }

    #[test]
    fn comment_removal_retains_owned_empty_parts_with_opaque_root_payload() {
        let mut document = Document::new();
        document.add_paragraph("anchor");
        let id = document
            .add_comment(
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
                "Ada",
                None,
                "root",
            )
            .unwrap();
        let raw =
            br#"<x:opaque xmlns:x="urn:producer" x:flag='exact'><x:child /></x:opaque>"#.to_vec();
        document
            .comments
            .as_mut()
            .unwrap()
            .extra_xml
            .push((1, raw.clone()));
        document
            .comments_extended
            .as_mut()
            .unwrap()
            .extra_xml
            .push((1, raw.clone()));
        document.comments_dirty = true;
        assert!(document.remove_comment(id).unwrap());
        assert!(document.comments().is_empty());
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        for part in [DEFAULT_COMMENTS_PART, DEFAULT_COMMENTS_EXTENDED_PART] {
            let xml = package.get_part(part).unwrap();
            assert!(xml.windows(raw.len()).any(|window| window == raw), "{part}");
        }
    }

    #[test]
    fn removing_the_last_owned_comment_retires_its_complete_identifier_bundle() {
        fn final_document(with_history: bool) -> Document {
            let mut document =
                Document::new_with_profile(crate::document::WordCreationProfile::Minimal(
                    crate::document::WordPackageClass::Document,
                ));
            document.add_paragraph("review");
            let range = RunRange {
                start: RunPosition {
                    body_index: 0,
                    run_index: 0,
                },
                end: RunPosition {
                    body_index: 0,
                    run_index: 1,
                },
            };
            if with_history {
                let old = document.add_comment(range, "Ada", None, "Old").unwrap();
                assert!(document.remove_comment(old).unwrap());
                assert!(document.comments_part_name.is_none());
                assert!(document.comments_extended_part_name.is_none());
                assert!(document.package.get_part(DEFAULT_COMMENTS_PART).is_none());
                assert!(
                    document
                        .package
                        .get_part(DEFAULT_COMMENTS_EXTENDED_PART)
                        .is_none()
                );
            }
            let id = document.add_comment(range, "Ada", None, "Final").unwrap();
            assert_eq!(id, 0);
            document
        }

        let mut history = final_document(true);
        let mut direct = final_document(false);
        let history_bytes = history.to_bytes().unwrap();
        let direct_bytes = direct.to_bytes().unwrap();
        assert_eq!(history_bytes, direct_bytes);
        let history_package = OpcPackage::from_reader(std::io::Cursor::new(history_bytes)).unwrap();
        let direct_package = OpcPackage::from_reader(std::io::Cursor::new(direct_bytes)).unwrap();
        assert_eq!(
            history_package.get_part(DEFAULT_COMMENTS_PART),
            direct_package.get_part(DEFAULT_COMMENTS_PART)
        );
    }

    #[test]
    fn comment_reference_splits_a_hyperlink_at_the_range_end() {
        let mut document = Document::new();
        let mut paragraph = document.add_paragraph("");
        paragraph.add_run("left");
        paragraph.add_run("right");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            panic!("body item should remain a paragraph");
        };
        paragraph.hyperlinks.push(HyperlinkSpan {
            rel_id: Some("rIdLink".to_owned()),
            anchor: None,
            tooltip: None,
            doc_location: None,
            run_start: 0,
            run_end: 2,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });

        document
            .add_comment(
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
                "Ada",
                None,
                "Review",
            )
            .unwrap();

        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            panic!("body item should remain a paragraph");
        };
        assert_eq!(paragraph.hyperlinks.len(), 2);
        assert_eq!(paragraph.hyperlinks[0].run_start, 0);
        assert_eq!(paragraph.hyperlinks[0].run_end, 1);
        assert_eq!(paragraph.hyperlinks[1].run_start, 2);
        assert_eq!(paragraph.hyperlinks[1].run_end, 3);
    }

    #[test]
    fn resolving_a_reply_resolves_the_thread_root() {
        let mut document = Document::new();
        document.add_paragraph("thread");
        let root = document
            .add_comment(
                RunRange {
                    start: RunPosition {
                        body_index: 0,
                        run_index: 0,
                    },
                    end: RunPosition {
                        body_index: 0,
                        run_index: 1,
                    },
                },
                "Ada",
                None,
                "Review",
            )
            .unwrap();
        let reply = document.reply_to(root, "Ben", "Done").unwrap();

        assert!(document.resolve_comment(reply, true).unwrap());
        let comments = document.comments();
        assert!(comments[0].resolved());
        assert!(!comments[1].resolved());
    }

    /// Five comments whose `w15:commentEx` rows are keyed, as Word and Google
    /// Docs key them, by the paraId of each comment's last paragraph: a reply
    /// of two paragraphs, a reply to a parent of two paragraphs and a
    /// resolved comment of two paragraphs.
    fn multi_paragraph_threads() -> Document {
        const COMMENTS: [(&str, &[&str]); 5] = [
            ("Ada", &["1A000001"]),
            ("Ben", &["1B000001", "1B000002"]),
            ("Ada", &["2A000001", "2A000002"]),
            ("Ben", &["2B000001"]),
            ("Ada", &["3A000001", "3A000002"]),
        ];
        let mut document = Document::new();
        let mut paragraph = document.add_paragraph("");
        paragraph.add_run("anchor");
        let range = RunRange {
            start: RunPosition {
                body_index: 0,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 0,
                run_index: 1,
            },
        };
        let ids = COMMENTS
            .iter()
            .map(|(author, _)| document.add_comment(range, author, None, "x").unwrap())
            .collect::<Vec<_>>();
        let comments = ids
            .iter()
            .zip(COMMENTS)
            .map(|(id, (author, para_ids))| {
                let paragraphs = para_ids
                    .iter()
                    .map(|para_id| {
                        format!(
                            r#"<w:p w14:paraId="{para_id}"><w:r><w:t>{para_id}</w:t></w:r></w:p>"#
                        )
                    })
                    .collect::<String>();
                format!(r#"<w:comment w:id="{id}" w:author="{author}">{paragraphs}</w:comment>"#)
            })
            .collect::<String>();
        let comments = format!(
            r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">{comments}</w:comments>"#
        );
        let extended = r#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="1A000001" w15:done="0"/><w15:commentEx w15:paraId="1B000002" w15:paraIdParent="1A000001" w15:done="0"/><w15:commentEx w15:paraId="2A000002" w15:done="0"/><w15:commentEx w15:paraId="2B000001" w15:paraIdParent="2A000002" w15:done="0"/><w15:commentEx w15:paraId="3A000002" w15:done="1"/></w15:commentsEx>"#;
        let mut package =
            OpcPackage::from_reader(std::io::Cursor::new(document.to_bytes().unwrap())).unwrap();
        package
            .parts
            .insert(DEFAULT_COMMENTS_PART.to_owned(), comments.into_bytes());
        package.parts.insert(
            DEFAULT_COMMENTS_EXTENDED_PART.to_owned(),
            extended.as_bytes().to_vec(),
        );
        let mut bytes = std::io::Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        Document::from_bytes(bytes.get_ref()).unwrap()
    }

    fn comment_extensions(document: &Document) -> Vec<(String, Option<String>, Option<bool>)> {
        document
            .comments_extended
            .as_ref()
            .unwrap()
            .comments
            .iter()
            .map(|entry| {
                (
                    entry.para_id.clone(),
                    entry.para_id_parent.clone(),
                    entry.done,
                )
            })
            .collect()
    }

    #[test]
    fn threads_and_resolved_state_read_through_the_last_paragraph() {
        let document = multi_paragraph_threads();
        let comments = document.comments();
        let ids = comments.iter().map(CommentRef::id).collect::<Vec<_>>();
        let observed = comments
            .iter()
            .map(|comment| (comment.parent_id(), comment.resolved()))
            .collect::<Vec<_>>();

        assert_eq!(
            observed,
            [
                (None, false),
                (Some(ids[0]), false),
                (None, false),
                (Some(ids[2]), false),
                (None, true),
            ]
        );
        assert_eq!(comments[1].text(), "1B000001\n1B000002");
    }

    #[test]
    fn reply_and_resolve_write_the_parent_last_paragraph() {
        let mut document = multi_paragraph_threads();
        let parent = document.comments()[2].id();

        let reply = document.reply_to(parent, "Cy", "new reply").unwrap();
        assert!(document.resolve_comment(parent, true).unwrap());

        let extensions = comment_extensions(&document);
        let reply_row = extensions.last().unwrap();
        assert_eq!(reply_row.1.as_deref(), Some("2A000002"));
        assert!(
            extensions
                .iter()
                .any(|(para_id, _, done)| para_id == "2A000002" && *done == Some(true))
        );
        assert!(
            extensions
                .iter()
                .all(|(para_id, _, _)| para_id != "2A000001")
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let comments = reopened.comments();
        let reread = comments
            .iter()
            .find(|comment| comment.id() == reply)
            .unwrap();
        assert_eq!(reread.parent_id(), Some(parent));
        assert!(comments[2].resolved());
    }

    #[test]
    fn removing_a_parent_of_several_paragraphs_removes_its_reply() {
        let mut document = multi_paragraph_threads();
        let ids = document
            .comments()
            .iter()
            .map(CommentRef::id)
            .collect::<Vec<_>>();

        assert!(document.remove_comment(ids[2]).unwrap());

        let remaining = document
            .comments()
            .iter()
            .map(CommentRef::id)
            .collect::<Vec<_>>();
        assert_eq!(remaining, [ids[0], ids[1], ids[4]]);
        assert!(
            comment_extensions(&document)
                .iter()
                .all(|(para_id, _, _)| para_id != "2A000002" && para_id != "2B000001")
        );
    }

    #[test]
    fn a_windows_line_ending_starts_a_paragraph_without_a_carriage_return() {
        let mut document = multi_paragraph_threads();
        let root = document.comments()[0].id();

        let reply = document.reply_to(root, "Cy", "first\r\nsecond").unwrap();

        let comments = document.comments();
        let reply = comments
            .iter()
            .find(|comment| comment.id() == reply)
            .unwrap();
        assert_eq!(reply.text(), "first\nsecond");
    }

    #[test]
    fn each_line_of_a_comment_text_becomes_one_paragraph() {
        let mut document = multi_paragraph_threads();
        let read = document.comments()[1].text();

        let root = document.comments()[0].id();
        let reply = document.reply_to(root, "Cy", &read).unwrap();
        let comment = document
            .comments
            .as_ref()
            .unwrap()
            .comments
            .iter()
            .find(|comment| comment.id == reply)
            .unwrap();
        assert_eq!(comment.paragraphs.len(), 2);
        let last = comment.paragraph_ids[1].clone().unwrap();
        assert_ne!(comment.paragraph_ids[0].as_deref(), Some(last.as_str()));
        assert_eq!(
            comment_extensions(&document).last().unwrap(),
            &(last, Some("1A000001".to_owned()), None)
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let comments = reopened.comments();
        let reread = comments
            .iter()
            .find(|comment| comment.id() == reply)
            .unwrap();
        assert_eq!(reread.text(), read);
        assert_eq!(reread.parent_id(), Some(root));
    }

    #[test]
    fn every_comment_entry_point_preserves_empty_lines_and_unique_paragraph_ids() {
        let mut document = Document::new();
        document.add_paragraph("anchor");
        let range = RunRange {
            start: RunPosition {
                body_index: 0,
                run_index: 0,
            },
            end: RunPosition {
                body_index: 0,
                run_index: 1,
            },
        };
        let text = "first\n\nlast\n";
        let root = document.add_comment(range, "Ada", None, text).unwrap();
        let found = document
            .add_comment_on_text("anchor", 0, "Ben", None, text, None)
            .unwrap();
        let location = document.paragraph_story_location(0).unwrap().unwrap();
        let story = document
            .add_story_comment(
                StoryRunRange {
                    start: StoryRunPosition {
                        location: location.clone(),
                        run_index: 0,
                    },
                    end: StoryRunPosition {
                        location,
                        run_index: 1,
                    },
                },
                "Cy",
                None,
                text,
            )
            .unwrap();
        let reply = document.reply_to(root, "Dee", text).unwrap();
        let empty = document.reply_to(root, "Eve", "").unwrap();
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let model = reopened.comments.as_ref().unwrap();
        let mut occupied = HashSet::new();
        for id in [root, found, story, reply, empty] {
            let comment = model.comments.iter().find(|item| item.id == id).unwrap();
            let expected = if id == empty { "" } else { text };
            assert_eq!(comment.paragraphs.len(), expected.split('\n').count());
            assert_eq!(
                reopened
                    .comments()
                    .iter()
                    .find(|item| item.id() == id)
                    .unwrap()
                    .text(),
                expected
            );
            for para_id in comment.paragraph_ids.iter().flatten() {
                assert!(occupied.insert(para_id.clone()));
            }
        }
        let xml = std::str::from_utf8(reopened.package.parts.get(DEFAULT_COMMENTS_PART).unwrap())
            .unwrap();
        assert!(!xml.contains("first\n"));
    }

    #[test]
    fn missing_final_paragraph_ids_are_allocated_without_replacing_earlier_ids() {
        for reply in [false, true] {
            let mut document = multi_paragraph_threads();
            let parent = document.comments()[2].id();
            document.comments.as_mut().unwrap().comments[2].paragraph_ids[1] = None;
            if reply {
                document.reply_to(parent, "Cy", "reply\nlast").unwrap();
            } else {
                document.resolve_comment(parent, true).unwrap();
            }
            let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
            let comment = &reopened.comments.as_ref().unwrap().comments[2];
            assert_eq!(comment.paragraph_ids[0].as_deref(), Some("2A000001"));
            let last = comment.paragraph_ids[1].as_deref().unwrap();
            assert_ne!(last, "2A000001");
            assert!(
                comment_extensions(&reopened)
                    .iter()
                    .any(|(id, _, done)| id == last && (reply || *done == Some(true)))
            );
            if reply {
                assert_eq!(
                    reopened.comments().last().unwrap().parent_id(),
                    Some(parent)
                );
            }
        }
    }

    #[test]
    fn fragment_import_keeps_last_paragraph_threads_and_unsupported_xml() {
        let mut source = multi_paragraph_threads();
        let root = source.comments()[2].id();
        let raw = br#"<x:keep xmlns:x="urn:producer" x:flag='exact'><x:child /></x:keep>"#.to_vec();
        source.comments.as_mut().unwrap().comments[2]
            .extra_xml
            .push((1, raw.clone()));
        let mut destination = Document::new();
        let remap = destination
            .import_fragment_comments_staged(
                source.comments.as_ref(),
                source.comments_extended.as_ref(),
                &[root.to_string()],
            )
            .unwrap();
        assert_eq!(remap.len(), 2);
        let bytes = destination.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let comments = reopened.comments();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[1].parent_id(), Some(comments[0].id()));
        let xml = reopened.package.parts.get(DEFAULT_COMMENTS_PART).unwrap();
        assert!(xml.windows(raw.len()).any(|window| window == raw));
    }

    #[test]
    fn removing_a_root_removes_nested_replies_linked_to_legacy_first_paragraphs() {
        let mut document = multi_paragraph_threads();
        let root = document.comments()[0].id();
        let child = document.comments()[1].id();
        let grandchild = document.reply_to(child, "Cy", "nested").unwrap();
        let own = document
            .comments
            .as_ref()
            .unwrap()
            .comments
            .last()
            .unwrap()
            .paragraph_ids[0]
            .clone()
            .unwrap();
        let row = document
            .comments_extended
            .as_mut()
            .unwrap()
            .comments
            .iter_mut()
            .find(|row| row.para_id == own)
            .unwrap();
        row.para_id_parent = Some("1B000001".to_owned());
        assert_eq!(document.comments().last().unwrap().parent_id(), Some(child));
        assert!(document.remove_comment(root).unwrap());
        assert!(
            document
                .comments()
                .iter()
                .all(|comment| ![root, child, grandchild].contains(&comment.id()))
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.comments().len(), 3);
    }

    #[test]
    fn removing_a_reference_run_keeps_an_unrelated_empty_run() {
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R {
            properties: None,
            content: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        paragraph.runs.push(CT_R {
            properties: Some(Default::default()),
            content: vec![RunContent::CommentReference {
                id: 7,
                raw_before: 0,
            }],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });

        remove_anchors_from_paragraph(&mut paragraph, &HashSet::from([7]));

        assert_eq!(paragraph.runs.len(), 1);
        assert!(paragraph.runs[0].content.is_empty());
    }

    #[test]
    fn word_comment_candidate_is_bound_to_recorded_sha() {
        let output = std::env::temp_dir().join(format!(
            "rdocx-f148-word-comment-{}.docx",
            std::process::id()
        ));
        word_comment_candidate()
            .save(&output)
            .expect("write SHA-bound candidate");
        assert_eq!(sha256(&output), WORD_COMMENT_CANDIDATE_SHA256);
        fs::remove_file(output).expect("remove temporary candidate");
    }

    #[test]
    #[ignore = "requires pinned Microsoft Word and human thread UI evidence"]
    fn word_opens_comment_reply_and_resolved_thread_without_repair() {
        let output = std::env::var_os("RDOCX_WORD_COMMENT_GATE_OUTPUT")
            .map(PathBuf::from)
            .expect("set RDOCX_WORD_COMMENT_GATE_OUTPUT to the SHA-bound .docx path");
        word_comment_candidate()
            .save(&output)
            .expect("write Word comment candidate");
        assert_eq!(sha256(&output), WORD_COMMENT_CANDIDATE_SHA256);
        let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
        assert_eq!(
            plist_value(plist, "CFBundleShortVersionString"),
            WORD_VERSION
        );
        assert_eq!(plist_value(plist, "CFBundleVersion"), WORD_BUILD);
    }

    fn sha256(path: &Path) -> String {
        let output = Command::new("shasum")
            .args(["-a", "256"])
            .arg(path)
            .output()
            .unwrap_or_else(|error| panic!("{}: run shasum: {error}", path.display()));
        assert!(
            output.status.success(),
            "shasum failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        String::from_utf8(output.stdout)
            .expect("shasum output is utf8")
            .split_whitespace()
            .next()
            .expect("shasum digest")
            .to_owned()
    }

    fn plist_value(path: &str, key: &str) -> String {
        let output = Command::new("/usr/libexec/PlistBuddy")
            .args(["-c", &format!("Print :{key}"), path])
            .output()
            .expect("read application plist");
        assert!(output.status.success());
        String::from_utf8(output.stdout)
            .expect("plist value is utf8")
            .trim()
            .to_owned()
    }
    #[test]
    fn glossary_omission_preserves_mixed_carriers_and_refuses_opaque_marker_payload() {
        let marker = r#"<q:commentReference q:id="4"/>"#;
        let xml = format!(
            r#"<q:p xmlns:q="{}" xmlns:x="urn:producer"><q:r x:keep="yes"><q:rPr><q:b/></q:rPr><?keep same?>{marker}<!--stay--><x:commentReference x:id="4"/></q:r></q:p>"#,
            rdocx_oxml::namespace::W_NS
        );
        assert_eq!(
            Document::omit_comment_markers(xml.as_bytes()).unwrap(),
            xml.replace(marker, "").as_bytes()
        );
        for invalid in [
            r#"<q:commentReference/>"#,
            r#"<q:commentReference q:id="bad"/>"#,
            r#"<q:commentReference q:id="4"><x:opaque/></q:commentReference>"#,
        ] {
            let invalid = xml.replace(marker, invalid);
            assert!(Document::omit_comment_markers(invalid.as_bytes()).is_err());
        }
    }
}
