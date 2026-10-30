//! Tracked revision metadata and read-only content projections.

use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use crate::content_control::{CT_Sdt, SdtContent};
use crate::document::{BodyContent, CT_Document, CT_SectPr};
use crate::namespace::W_NS;
use crate::numbering::{parse_scoped_ppr, parse_scoped_rpr, word_prefixes_at};
use crate::properties::{CT_PPr, CT_RPr, is_word_element};
use crate::raw_xml::capture_empty_element;
use crate::table::{CT_Row, CT_Tbl, CT_TblPr, CT_Tc, CellContent};
use crate::text::{AcceptedRunPath, AcceptedRunPathSegment, CT_P, CT_R, hyperlink_revision_index};

const MAX_REVISION_NESTING_DEPTH: usize = 32;

/// Physical legacy read facts captured by the one canonical paragraph parser.
/// These are operation-local, before complex fields regroup physical runs.
#[derive(Default)]
pub(crate) struct RevisionSourceFacts {
    pub(crate) runs: Vec<CT_R>,
    pub(crate) revisions: Vec<(usize, CT_Revision)>,
    copy_nodes_bound: usize,
    copy_bytes_bound: usize,
    copy_nodes: usize,
    copy_bytes: usize,
}

impl RevisionSourceFacts {
    fn with_copy_bound(bound: (usize, usize)) -> Self {
        Self {
            copy_nodes_bound: bound.0,
            copy_bytes_bound: bound.1,
            ..Self::default()
        }
    }

    fn charge(&mut self, nodes: usize, bytes: usize) -> crate::Result<()> {
        #[cfg(test)]
        let (node_increment, byte_increment) = (nodes, bytes);
        let nodes = self.copy_nodes.checked_add(nodes).ok_or_else(|| {
            crate::OxmlError::InvalidValue("secondary source-copy visit overflow".into())
        })?;
        let bytes = self.copy_bytes.checked_add(bytes).ok_or_else(|| {
            crate::OxmlError::InvalidValue("secondary source-copy byte overflow".into())
        })?;
        if nodes > self.copy_nodes_bound || bytes > self.copy_bytes_bound {
            return Err(crate::OxmlError::InvalidValue(
                "secondary source-copy work exceeds its admitted source ancestry".into(),
            ));
        }
        self.copy_nodes = nodes;
        self.copy_bytes = bytes;
        #[cfg(test)]
        {
            ORDERED_DIAGNOSTIC_SECONDARY_COPY_NODES.set(
                ORDERED_DIAGNOSTIC_SECONDARY_COPY_NODES
                    .get()
                    .checked_add(node_increment)
                    .expect("secondary copy counter overflow"),
            );
            ORDERED_DIAGNOSTIC_SECONDARY_CHARGED_SOURCE_BYTES.set(
                ORDERED_DIAGNOSTIC_SECONDARY_CHARGED_SOURCE_BYTES
                    .get()
                    .checked_add(byte_increment)
                    .expect("secondary copy counter overflow"),
            );
        }
        Ok(())
    }

    pub(crate) fn push_run(&mut self, run: &CT_R, source: &[u8]) -> crate::Result<()> {
        self.charge(1, source.len())?;
        self.runs.push(run.clone());
        Ok(())
    }

    pub(crate) fn push_revision(&mut self, revision: &CT_Revision) -> crate::Result<()> {
        let mut reader = Reader::from_reader(revision.raw_xml.as_slice());
        let root = match reader.read_event()? {
            Event::Start(root) | Event::Empty(root) => root,
            _ => return Ok(()),
        };
        let prefixes = word_prefixes_at(&root, &revision.parse_word_prefixes)?;
        if optional_word_attribute(&root, b"author", &prefixes)?.is_none() {
            return Ok(());
        }
        let mut nodes = 0usize;
        let mut visible_raw_bytes = 0usize;
        let mut pending = vec![revision];
        while let Some(current) = pending.pop() {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                crate::OxmlError::InvalidValue("secondary revision-copy visit overflow".into())
            })?;
            visible_raw_bytes = visible_raw_bytes
                .checked_add(current.raw_xml.len())
                .ok_or_else(|| {
                    crate::OxmlError::InvalidValue("secondary revision-copy byte overflow".into())
                })?;
            if let RevisionContent::Runs(runs) = current.content() {
                nodes = nodes.checked_add(runs.len()).ok_or_else(|| {
                    crate::OxmlError::InvalidValue("secondary run-copy visit overflow".into())
                })?;
                for run in runs {
                    for raw in &run.extra_xml {
                        visible_raw_bytes =
                            visible_raw_bytes.checked_add(raw.len()).ok_or_else(|| {
                                crate::OxmlError::InvalidValue(
                                    "secondary run-copy byte overflow".into(),
                                )
                            })?;
                    }
                    if let Some(properties) = &run.properties {
                        if let Some(change) = &properties.change {
                            visible_raw_bytes = visible_raw_bytes
                                .checked_add(change.raw_xml.len())
                                .ok_or_else(|| {
                                    crate::OxmlError::InvalidValue(
                                        "secondary typed run-property copy byte overflow".into(),
                                    )
                                })?;
                        }
                        for raw in &properties.revision_xml {
                            visible_raw_bytes =
                                visible_raw_bytes.checked_add(raw.len()).ok_or_else(|| {
                                    crate::OxmlError::InvalidValue(
                                        "secondary run-property copy byte overflow".into(),
                                    )
                                })?;
                        }
                    }
                }
            }
            let prior_raw: Vec<&[u8]> = match current.content() {
                RevisionContent::PriorRunProperties(properties) => {
                    properties.revision_xml.iter().map(Vec::as_slice).collect()
                }
                RevisionContent::PriorParagraphProperties(properties) => properties
                    .revision_xml
                    .iter()
                    .chain(&properties.numbering_revision_xml)
                    .map(Vec::as_slice)
                    .chain(
                        properties
                            .num_pr_extra_xml
                            .iter()
                            .map(|(_, _, raw)| raw.as_slice()),
                    )
                    .chain(
                        properties
                            .num_ilvl_raw
                            .iter()
                            .map(|(_, raw, _)| raw.as_slice()),
                    )
                    .chain(
                        properties
                            .num_id_raw
                            .iter()
                            .map(|(_, raw, _)| raw.as_slice()),
                    )
                    .collect(),
                RevisionContent::PriorTableProperties(properties) => properties
                    .revision_xml
                    .iter()
                    .map(Vec::as_slice)
                    .chain(properties.extra_xml.iter().map(|(_, raw)| raw.as_slice()))
                    .collect(),
                RevisionContent::PriorSectionProperties(properties) => {
                    properties.extra_xml.iter().map(Vec::as_slice).collect()
                }
                _ => Vec::new(),
            };
            for raw in prior_raw {
                visible_raw_bytes = visible_raw_bytes.checked_add(raw.len()).ok_or_else(|| {
                    crate::OxmlError::InvalidValue(
                        "secondary prior-property copy byte overflow".into(),
                    )
                })?;
            }
            pending.extend(current.nested_revisions.iter().map(|(_, nested)| nested));
        }
        // The visible retained buffers above are observations of known copies.
        // Source allowance also covers private payloads, never as a measurement.
        let copy_bound = crate::document::admit_word_xml(
            &revision.raw_xml,
            &revision.parse_word_prefixes,
            None,
        )?;
        self.charge(nodes, visible_raw_bytes.max(copy_bound.1))?;
        #[cfg(test)]
        {
            ORDERED_DIAGNOSTIC_SECONDARY_VISIBLE_RAW_BYTES.set(
                ORDERED_DIAGNOSTIC_SECONDARY_VISIBLE_RAW_BYTES
                    .get()
                    .checked_add(visible_raw_bytes)
                    .expect("secondary visible raw byte overflow"),
            );
        }
        let mut legacy = revision.clone();
        legacy.require_author = true;
        self.revisions.push((self.runs.len(), legacy));
        Ok(())
    }

    /// Only a hyperlink subtree that the canonical reader retained opaque.
    /// Its legacy direct runs have not been parsed by the canonical route.
    pub(crate) fn append_opaque_hyperlink(
        &mut self,
        raw: &[u8],
        prefixes: &[String],
    ) -> crate::Result<()> {
        let mut reader = Reader::from_reader(raw);
        let root = match reader.read_event()? {
            Event::Start(root) => root,
            _ => return Ok(()),
        };
        let prefixes = word_prefixes_at(&root, prefixes)?;
        parse_run_content_container(
            &mut reader,
            b"hyperlink",
            &prefixes,
            &mut self.runs,
            &mut self.revisions,
        )
    }
}

/// The modeled tracked-change element kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionKind {
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    RunPropertyChange,
    ParagraphPropertyChange,
    TablePropertyChange,
    SectionPropertyChange,
}

/// The typed content projected from a preserved revision subtree.
#[derive(Debug, Clone, PartialEq)]
pub enum RevisionContent {
    Runs(Vec<CT_R>),
    Marker,
    PriorRunProperties(Box<CT_RPr>),
    PriorParagraphProperties(Box<CT_PPr>),
    PriorTableProperties(Box<CT_TblPr>),
    PriorSectionProperties(Box<CT_SectPr>),
}

/// A tracked revision whose captured subtree remains its serialization source.
#[derive(Debug, Clone)]
#[allow(non_camel_case_types)]
pub struct CT_Revision {
    kind: RevisionKind,
    id: i32,
    author: String,
    timestamp: Option<String>,
    raw_xml: Vec<u8>,
    content: RevisionContent,
    content_paragraph: Option<Arc<CT_P>>,
    nested_revisions: Vec<(usize, CT_Revision)>,
    parse_word_prefixes: Vec<String>,
    require_author: bool,
    accepted_fields_pending: bool,
    accepted_payload_pending: bool,
}

impl PartialEq for CT_Revision {
    fn eq(&self, other: &Self) -> bool {
        if self.kind != other.kind
            || self.id != other.id
            || self.author != other.author
            || self.timestamp != other.timestamp
            || self.raw_xml != other.raw_xml
            || self.content != other.content
            || self.content_paragraph != other.content_paragraph
            || self.nested_revisions != other.nested_revisions
            || self.require_author != other.require_author
        {
            return false;
        }
        // Context enumeration order is parser bookkeeping, not revision identity.
        // Keep complete entries and multiplicity without changing parser inputs.
        let mut left = self.parse_word_prefixes.iter().collect::<Vec<_>>();
        let mut right = other.parse_word_prefixes.iter().collect::<Vec<_>>();
        left.sort_unstable();
        right.sort_unstable();
        left == right
    }
}

#[cfg(test)]
thread_local! {
    pub(crate) static ORDERED_DIAGNOSTIC_REVISION_PARSE_CALLS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static ORDERED_DIAGNOSTIC_REVISION_PARSE_BYTES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static ORDERED_DIAGNOSTIC_SECONDARY_COPY_NODES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static ORDERED_DIAGNOSTIC_SECONDARY_CHARGED_SOURCE_BYTES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static ORDERED_DIAGNOSTIC_SECONDARY_VISIBLE_RAW_BYTES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

impl CT_Revision {
    pub(crate) fn from_raw(raw_xml: Vec<u8>, word_prefixes: &[String]) -> Option<Self> {
        Self::try_from_raw(raw_xml, word_prefixes, true).ok()
    }

    pub(crate) fn from_raw_content(raw_xml: Vec<u8>, word_prefixes: &[String]) -> Option<Self> {
        Self::try_from_raw(raw_xml, word_prefixes, false).ok()
    }

    /// Whether a preserved wrapper is admitted by the complete typed parser.
    #[doc(hidden)]
    pub fn story_raw_is_typed(raw_xml: &[u8], word_prefixes: &[String]) -> bool {
        Self::try_from_raw(raw_xml.to_vec(), word_prefixes, true).is_ok()
    }

    pub(crate) fn into_raw_xml(self) -> Vec<u8> {
        self.raw_xml
    }

    fn try_from_raw(
        raw_xml: Vec<u8>,
        word_prefixes: &[String],
        require_author: bool,
    ) -> crate::Result<Self> {
        #[cfg(test)]
        {
            ORDERED_DIAGNOSTIC_REVISION_PARSE_CALLS.set(
                ORDERED_DIAGNOSTIC_REVISION_PARSE_CALLS
                    .get()
                    .checked_add(1)
                    .expect("constructor counter overflow"),
            );
            ORDERED_DIAGNOSTIC_REVISION_PARSE_BYTES.set(
                ORDERED_DIAGNOSTIC_REVISION_PARSE_BYTES
                    .get()
                    .checked_add(raw_xml.len())
                    .expect("constructor counter overflow"),
            );
        }
        validate_revision_nesting_depth(&raw_xml, word_prefixes)?;
        let mut reader = Reader::from_reader(raw_xml.as_slice());
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let (kind, prefixes) = loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let prefixes = word_prefixes_at(&start, word_prefixes)?;
                    let kind =
                        revision_kind(start.name().as_ref(), &prefixes).ok_or_else(|| {
                            crate::OxmlError::InvalidValue(
                                "unsupported revision element".to_owned(),
                            )
                        })?;
                    let id = required_word_attribute(&start, b"id", &prefixes)?.parse()?;
                    let author = optional_word_attribute(&start, b"author", &prefixes)?;
                    let author = match (author, require_author) {
                        (Some(author), _) => author,
                        (None, true) => required_word_attribute(&start, b"author", &prefixes)?,
                        (None, false) => String::new(),
                    };
                    let timestamp = optional_word_attribute(&start, b"date", &prefixes)?;
                    break ((kind, id, author, timestamp), prefixes);
                }
                Event::Empty(start) => {
                    let prefixes = word_prefixes_at(&start, word_prefixes)?;
                    let kind =
                        revision_kind(start.name().as_ref(), &prefixes).ok_or_else(|| {
                            crate::OxmlError::InvalidValue(
                                "unsupported revision element".to_owned(),
                            )
                        })?;
                    let id = required_word_attribute(&start, b"id", &prefixes)?.parse()?;
                    let author = optional_word_attribute(&start, b"author", &prefixes)?;
                    let author = match (author, require_author) {
                        (Some(author), _) => author,
                        (None, true) => required_word_attribute(&start, b"author", &prefixes)?,
                        (None, false) => String::new(),
                    };
                    let timestamp = optional_word_attribute(&start, b"date", &prefixes)?;
                    return Ok(Self {
                        kind,
                        id,
                        author,
                        timestamp,
                        raw_xml,
                        content: RevisionContent::Marker,
                        content_paragraph: None,
                        nested_revisions: Vec::new(),
                        parse_word_prefixes: word_prefixes.to_vec(),
                        require_author,
                        accepted_fields_pending: false,
                        accepted_payload_pending: false,
                    });
                }
                Event::Eof => {
                    return Err(crate::OxmlError::MissingElement(
                        "revision element".to_owned(),
                    ));
                }
                _ => {}
            }
            buffer.clear();
        };
        let (kind, id, author, timestamp) = kind;
        let (content, content_paragraph, nested_revisions) =
            if matches!(kind, RevisionKind::Insertion | RevisionKind::MoveTo)
                || matches!(kind, RevisionKind::Deletion | RevisionKind::MoveFrom)
                    && crate::text::source_has_ordered_field_owner(&raw_xml, &prefixes)?
            {
                let copy_bound = crate::document::admit_word_xml(&raw_xml, word_prefixes, None)?;
                let (paragraph, facts) =
                    parse_revision_content_with_facts(&raw_xml, &prefixes, copy_bound)?;
                let content = if facts.runs.is_empty() {
                    RevisionContent::Marker
                } else {
                    RevisionContent::Runs(facts.runs)
                };
                (content, Some(Arc::new(paragraph)), facts.revisions)
            } else {
                let (content, nested_revisions) = parse_content(&mut reader, kind, &prefixes)?;
                (content, None, nested_revisions)
            };
        Ok(Self {
            kind,
            id,
            author,
            timestamp,
            raw_xml,
            content,
            content_paragraph,
            nested_revisions,
            parse_word_prefixes: word_prefixes.to_vec(),
            require_author,
            accepted_fields_pending: false,
            accepted_payload_pending: false,
        })
    }

    pub fn kind(&self) -> RevisionKind {
        self.kind
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

    pub fn content(&self) -> &RevisionContent {
        &self.content
    }

    /// Return the paragraph projection for an insertion, when it has inline wrappers.
    #[doc(hidden)]
    pub fn content_paragraph(&self) -> Option<&CT_P> {
        self.content_paragraph.as_deref()
    }

    /// Mutate physical read-projection metadata without rewriting revision XML.
    /// Used on layout-only clones before selecting the accepted revision view.
    /// The original raw XML remains the serializer's authority.
    #[doc(hidden)]
    pub fn physical_content_paragraph_mut(&mut self) -> crate::Result<Option<&mut CT_P>> {
        if !matches!(
            self.kind,
            RevisionKind::Insertion
                | RevisionKind::Deletion
                | RevisionKind::MoveFrom
                | RevisionKind::MoveTo
        ) {
            return Ok(None);
        }
        if self.content_paragraph.is_none() {
            self.content_paragraph = Some(Arc::new(parse_accepted_revision_content(
                &self.raw_xml,
                &self.parse_word_prefixes,
            )?));
        }
        Ok(self.content_paragraph.as_mut().map(Arc::make_mut))
    }

    /// Return nested revision wrappers at their direct-run boundaries.
    #[doc(hidden)]
    pub fn nested_revisions(&self) -> &[(usize, CT_Revision)] {
        &self.nested_revisions
    }

    pub(crate) fn append_accepted_run_paths(
        &self,
        prefix: &mut Vec<AcceptedRunPathSegment>,
        output: &mut Vec<AcceptedRunPath>,
    ) {
        if !matches!(self.kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
            return;
        }
        if let Some(paragraph) = self.content_paragraph.as_deref() {
            super::text::append_accepted_paragraph_run_paths(paragraph, prefix, output);
        }
    }

    pub(crate) fn accepted_run_segments(&self, path: &[AcceptedRunPathSegment]) -> Option<&CT_R> {
        if !matches!(self.kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
            return None;
        }
        self.content_paragraph
            .as_deref()?
            .accepted_run_segments(path)
    }

    pub(crate) fn accepted_run_segments_mut(
        &mut self,
        path: &[AcceptedRunPathSegment],
    ) -> Option<&mut CT_R> {
        if !matches!(self.kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
            return None;
        }
        self.accepted_fields_pending = true;
        Arc::make_mut(self.content_paragraph.as_mut()?).accepted_run_segments_mut(path)
    }

    pub(crate) fn ordered_paragraph_mut(&mut self) -> Option<&mut CT_P> {
        if !matches!(self.kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
            return None;
        }
        self.accepted_fields_pending = true;
        self.accepted_payload_pending = true;
        Some(Arc::make_mut(self.content_paragraph.as_mut()?))
    }

    pub(crate) fn normalize_pending_fields(&mut self) -> crate::Result<()> {
        let Some(paragraph) = self.content_paragraph.as_deref() else {
            return Ok(());
        };
        if !self.accepted_fields_pending
            && !self.accepted_payload_pending
            && !paragraph.has_changed_field_owners()
        {
            return Ok(());
        }
        if !self.accepted_payload_pending && !paragraph.has_changed_field_owners() {
            self.accepted_fields_pending = false;
            return Ok(());
        }
        let mut candidate = self.clone();
        Arc::make_mut(
            candidate
                .content_paragraph
                .as_mut()
                .expect("stored paragraph"),
        )
        .normalize_field_owners()?;
        candidate.refresh_accepted_content()?;
        *self = candidate;
        Ok(())
    }

    pub(crate) fn replace_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
        replacement: CT_R,
    ) -> crate::Result<bool> {
        self.edit_accepted_paragraph(|paragraph| {
            paragraph.replace_accepted_run_segments(path, replacement)
        })
    }

    pub(crate) fn split_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
        offset: usize,
    ) -> crate::Result<bool> {
        self.edit_accepted_paragraph(|paragraph| {
            paragraph.split_accepted_run_segments(path, offset)
        })
    }

    /// Remove an accepted-view run of an insertion. Returns `None` for a
    /// stale path, and otherwise whether nothing is left in the wrapper, which
    /// is then left unchanged for its owner to remove. A move destination is
    /// refused, since its source would no longer match it.
    pub(crate) fn remove_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
    ) -> crate::Result<Option<bool>> {
        if self.kind == RevisionKind::MoveTo {
            return Err(crate::OxmlError::InvalidValue(
                "it is part of a tracked move destination, whose source would no longer match it"
                    .to_owned(),
            ));
        }
        let mut outcome = None;
        self.edit_accepted_paragraph(|paragraph| {
            if !paragraph.remove_accepted_run_segments(path)? {
                return Ok(false);
            }
            let emptied = paragraph.has_no_children();
            outcome = Some(emptied);
            Ok(!emptied)
        })?;
        Ok(outcome)
    }

    /// The paragraph of an insertion or a move destination, when
    /// [`Self::edit_accepted_paragraph`] can write an edit of it back.
    pub(crate) fn editable_accepted_paragraph(&self) -> Option<&CT_P> {
        if !matches!(self.kind, RevisionKind::Insertion | RevisionKind::MoveTo)
            || ensure_revision_word_prefix_is_mutable(&self.raw_xml).is_err()
        {
            return None;
        }
        self.content_paragraph.as_deref()
    }

    /// Apply `edit` to the paragraph of an insertion or a move destination
    /// and, when it reports a change, write its runs back inside the
    /// wrapper, whose start tag keeps its id, author and date. Returns false
    /// for any other revision.
    pub(crate) fn edit_accepted_paragraph(
        &mut self,
        edit: impl FnOnce(&mut CT_P) -> crate::Result<bool>,
    ) -> crate::Result<bool> {
        if !matches!(self.kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
            return Ok(false);
        }
        ensure_revision_word_prefix_is_mutable(&self.raw_xml)?;
        let mut candidate = self.clone();
        let Some(paragraph) = candidate.content_paragraph.as_mut().map(Arc::make_mut) else {
            return Ok(false);
        };
        if !edit(paragraph)? {
            return Ok(false);
        }
        candidate.refresh_accepted_content()?;
        *self = candidate;
        Ok(true)
    }

    pub(crate) fn shared_content_paragraph(&self) -> Option<Arc<CT_P>> {
        self.content_paragraph.clone()
    }

    // The caller retained this exact fresh parser occurrence before identity
    // rebinding. Legacy caches were cloned from it by RevisionSourceFacts.
    // Transfer only those proven construction aliases, not matching text/IDs.
    pub(crate) fn synchronize_source_fact_aliases(
        &mut self,
        before: &Arc<CT_P>,
    ) -> crate::Result<()> {
        let Some(after) = &mut self.content_paragraph else {
            return Ok(());
        };
        if Arc::ptr_eq(before, after) {
            return Ok(());
        }
        let after = Arc::make_mut(after);
        if before.revisions.len() != after.revisions.len() {
            return Err(crate::OxmlError::InvalidValue(
                "revision source-fact occurrences changed".into(),
            ));
        }
        for ((old_run, old_raw, old), (new_run, new_raw, new)) in
            before.revisions.iter().zip(&mut after.revisions)
        {
            if (old_run, old_raw) != (new_run, new_raw) || old.raw_xml != new.raw_xml {
                return Err(crate::OxmlError::InvalidValue(
                    "revision source-fact physical slot changed".into(),
                ));
            }
            if let Some(source) = &old.content_paragraph {
                new.synchronize_source_fact_aliases(source)?;
            }
        }
        for (_, legacy) in &mut self.nested_revisions {
            let Some(source) = &legacy.content_paragraph else {
                continue;
            };
            let mut matches =
                before
                    .revisions
                    .iter()
                    .enumerate()
                    .filter(|(_, (_, _, canonical))| {
                        canonical
                            .content_paragraph
                            .as_ref()
                            .is_some_and(|paragraph| Arc::ptr_eq(source, paragraph))
                    });
            let Some((index, (_, _, original))) = matches.next() else {
                continue;
            };
            if matches.next().is_some() || legacy.raw_xml != original.raw_xml {
                return Err(crate::OxmlError::InvalidValue(
                    "revision source-fact alias is ambiguous".into(),
                ));
            }
            legacy.adopt_source_fact_aliases(&after.revisions[index].2)?;
        }
        Ok(())
    }

    fn adopt_source_fact_aliases(&mut self, current: &Self) -> crate::Result<()> {
        if self.raw_xml != current.raw_xml
            || self.nested_revisions.len() != current.nested_revisions.len()
        {
            return Err(crate::OxmlError::InvalidValue(
                "revision source-fact cache occurrence changed".into(),
            ));
        }
        self.content_paragraph = current.content_paragraph.clone();
        for ((old_slot, old), (new_slot, new)) in self
            .nested_revisions
            .iter_mut()
            .zip(&current.nested_revisions)
        {
            if old_slot != new_slot {
                return Err(crate::OxmlError::InvalidValue(
                    "revision source-fact cache slot changed".into(),
                ));
            }
            old.adopt_source_fact_aliases(new)?;
        }
        Ok(())
    }

    fn refresh_accepted_content(&mut self) -> crate::Result<()> {
        let paragraph = self.content_paragraph.as_deref().ok_or_else(|| {
            crate::OxmlError::MissingElement("accepted revision paragraph".to_owned())
        })?;
        let paragraph_xml = crate::text::ordered_paragraph_xml(paragraph)?;
        let (content_start, content_end) = revision_content_bounds(&self.raw_xml)?;
        let capacity = self
            .raw_xml
            .len()
            .checked_sub(content_end - content_start)
            .and_then(|bytes| bytes.checked_add(paragraph_xml.len()))
            .ok_or_else(|| {
                crate::OxmlError::InvalidValue("revision refresh source byte overflow".into())
            })?;
        let mut updated = Vec::with_capacity(capacity);
        updated.extend_from_slice(&self.raw_xml[..content_start]);
        updated.extend_from_slice(&paragraph_xml);
        updated.extend_from_slice(&self.raw_xml[content_end..]);
        let prefixes = self.parse_word_prefixes.clone();
        let require_author = self.require_author;
        let mut normalized = Self::try_from_raw(updated, &prefixes, require_author)?;
        let parsed_source = normalized.content_paragraph.clone();
        if let Some(current) = normalized.content_paragraph.as_mut() {
            crate::text::retain_ordered_field_source_identities(paragraph, Arc::make_mut(current))?;
        }
        if let Some(source) = parsed_source {
            normalized.synchronize_source_fact_aliases(&source)?;
        }
        // Rebinding follows current source slots and does not create a new edit.
        normalized.accepted_fields_pending = false;
        normalized.accepted_payload_pending = false;
        *self = normalized;
        Ok(())
    }

    pub(crate) fn write_xml<W: std::io::Write>(
        &self,
        writer: &mut quick_xml::Writer<W>,
    ) -> crate::Result<()> {
        writer.get_mut().write_all(&self.raw_xml)?;
        Ok(())
    }

    pub(crate) fn write_xml_with_word_override<W: std::io::Write>(
        &self,
        writer: &mut quick_xml::Writer<W>,
        foreign_word_namespace: Option<&str>,
    ) -> crate::Result<()> {
        crate::text::write_raw_with_word_override(writer, &self.raw_xml, foreign_word_namespace)
    }
}

fn validate_revision_nesting_depth(raw_xml: &[u8], word_prefixes: &[String]) -> crate::Result<()> {
    let mut reader = Reader::from_reader(raw_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes = vec![(word_prefixes.to_vec(), false)];
    let mut revision_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let inherited = &scopes.last().expect("root namespace scope").0;
                let prefixes = word_prefixes_at(&start, inherited)?;
                let is_revision = revision_kind(start.name().as_ref(), &prefixes).is_some();
                if is_revision {
                    revision_depth += 1;
                    if revision_depth > MAX_REVISION_NESTING_DEPTH {
                        return Err(crate::OxmlError::InvalidValue(
                            "revision nesting depth exceeds reader limit".to_owned(),
                        ));
                    }
                }
                scopes.push((prefixes, is_revision));
            }
            Event::Empty(start) => {
                let inherited = &scopes.last().expect("root namespace scope").0;
                let prefixes = word_prefixes_at(&start, inherited)?;
                if revision_kind(start.name().as_ref(), &prefixes).is_some()
                    && revision_depth + 1 > MAX_REVISION_NESTING_DEPTH
                {
                    return Err(crate::OxmlError::InvalidValue(
                        "revision nesting depth exceeds reader limit".to_owned(),
                    ));
                }
            }
            Event::End(_) => {
                let (_, was_revision) = scopes.pop().expect("balanced namespace scope");
                if was_revision {
                    revision_depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(())
}

impl CT_Document {
    /// Return every valid modeled main-document revision in document order.
    pub fn revisions(&self) -> Vec<&CT_Revision> {
        let mut revisions = Vec::new();
        for content in &self.body.content {
            collect_body_content(content, &mut revisions);
        }
        if let Some(section) = &self.body.sect_pr {
            collect_section_properties(section, &mut revisions);
        }
        revisions
    }
}

fn collect_body_content<'a>(content: &'a BodyContent, revisions: &mut Vec<&'a CT_Revision>) {
    match content {
        BodyContent::Paragraph(paragraph) => collect_paragraph(paragraph, revisions),
        BodyContent::Table(table) => collect_table(table, revisions),
        BodyContent::ContentControl(control) => collect_control(control, revisions),
        BodyContent::RawXml(_) => {}
    }
}

fn collect_paragraph<'a>(paragraph: &'a CT_P, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(properties) = &paragraph.properties {
        if let Some(revision) = &properties.numbering_revision {
            revisions.push(revision);
        }
        if let Some(run_properties) = &properties.rpr {
            collect_run_properties(run_properties, revisions);
        }
        if let Some(section) = &properties.sect_pr {
            collect_section_properties(section, revisions);
        }
        if let Some(change) = &properties.change {
            revisions.push(change);
        }
    }

    for boundary in 0..=paragraph.runs.len() {
        let raw_count = paragraph
            .extra_xml
            .iter()
            .filter(|(at, _)| *at == boundary)
            .count();
        for raw_index in 0..=raw_count {
            for (_, _, _, control) in
                paragraph
                    .content_controls
                    .iter()
                    .filter(|(at, raw_before, _, _)| {
                        *at == boundary && (*raw_before).min(raw_count) == raw_index
                    })
            {
                collect_control(control, revisions);
            }
            for (_, _, revision) in paragraph.revisions.iter().filter(|(at, raw_before, _)| {
                *at == boundary
                    && hyperlink_revision_index(*raw_before).is_none()
                    && (*raw_before).min(raw_count) == raw_index
            }) {
                collect_revision(revision, revisions);
            }
            for (hyperlink_index, _hyperlink) in
                paragraph
                    .hyperlinks
                    .iter()
                    .enumerate()
                    .filter(|(_, hyperlink)| {
                        hyperlink.run_start == boundary
                            && hyperlink.run_end == boundary
                            && hyperlink.preserved_raw_before == Some(raw_index)
                    })
            {
                for (_, _, revision) in paragraph.revisions.iter().filter(|(at, slot, _)| {
                    *at == boundary && hyperlink_revision_index(*slot) == Some(hyperlink_index)
                }) {
                    collect_revision(revision, revisions);
                }
            }
        }
        for (_, _, revision) in paragraph.revisions.iter().filter(|(at, slot, _)| {
            *at == boundary
                && hyperlink_revision_index(*slot).is_some_and(|index| {
                    paragraph
                        .hyperlinks
                        .get(index)
                        .is_none_or(|hyperlink| hyperlink.preserved_raw_before.is_none())
                })
        }) {
            collect_revision(revision, revisions);
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            collect_run(run, revisions);
        }
    }
}

fn collect_run<'a>(run: &'a CT_R, revisions: &mut Vec<&'a CT_Revision>) {
    for content in &run.content {
        if let crate::text::RunContent::Field(field) = content {
            collect_field(field, revisions);
        }
    }
    if let Some(properties) = &run.properties {
        collect_run_properties(properties, revisions);
    }
}

fn collect_field<'a>(field: &'a crate::text::Field, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(paragraph) = field.stored_ordered_result() {
        collect_paragraph(paragraph, revisions);
    } else {
        for child in field.cached_fields_in_source_order() {
            collect_field(child, revisions);
        }
    }
}

fn collect_run_properties<'a>(properties: &'a CT_RPr, revisions: &mut Vec<&'a CT_Revision>) {
    revisions.extend(properties.revision_markers.iter());
    if let Some(change) = &properties.change {
        revisions.push(change);
    }
}

fn collect_section_properties<'a>(properties: &'a CT_SectPr, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(change) = &properties.change {
        revisions.push(change);
    }
}

fn collect_table<'a>(table: &'a CT_Tbl, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(change) = table
        .properties
        .as_ref()
        .and_then(|properties| properties.change.as_ref())
    {
        revisions.push(change);
    }
    for boundary in 0..=table.rows.len() {
        collect_controls_at_table_boundary(table, boundary, revisions);
        if let Some(row) = table.rows.get(boundary) {
            collect_row(row, revisions);
        }
    }
}

fn collect_controls_at_table_boundary<'a>(
    table: &'a CT_Tbl,
    boundary: usize,
    revisions: &mut Vec<&'a CT_Revision>,
) {
    let raw_count = table
        .extra_xml
        .iter()
        .filter(|(at, _)| *at == boundary)
        .count();
    for raw_index in 0..=raw_count {
        for (_, _, control) in table.content_controls.iter().filter(|(at, raw_before, _)| {
            *at == boundary && (*raw_before).min(raw_count) == raw_index
        }) {
            collect_control(control, revisions);
        }
    }
}

fn collect_row<'a>(row: &'a CT_Row, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(properties) = &row.properties {
        revisions.extend(properties.revision_markers.iter());
    }
    for boundary in 0..=row.cells.len() {
        let raw_count = row
            .extra_xml
            .iter()
            .filter(|(at, _)| *at == boundary)
            .count();
        for raw_index in 0..=raw_count {
            for (_, _, control) in row.content_controls.iter().filter(|(at, raw_before, _)| {
                *at == boundary && (*raw_before).min(raw_count) == raw_index
            }) {
                collect_control(control, revisions);
            }
        }
        if let Some(cell) = row.cells.get(boundary) {
            collect_cell(cell, revisions);
        }
    }
}

fn collect_cell<'a>(cell: &'a CT_Tc, revisions: &mut Vec<&'a CT_Revision>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => collect_paragraph(paragraph, revisions),
            CellContent::Table(table) => collect_table(table, revisions),
            CellContent::ContentControl(control) => collect_control(control, revisions),
        }
    }
}

fn collect_control<'a>(control: &'a CT_Sdt, revisions: &mut Vec<&'a CT_Revision>) {
    for (content_index, content) in control.content.iter().enumerate() {
        for (_, revision) in control
            .revisions
            .iter()
            .filter(|(at, _)| *at == content_index)
        {
            collect_revision(revision, revisions);
        }
        match content {
            SdtContent::Paragraph(paragraph) => collect_paragraph(paragraph, revisions),
            SdtContent::Table(table) => collect_table(table, revisions),
            SdtContent::Row(row) => collect_row(row, revisions),
            SdtContent::Cell(cell) => collect_cell(cell, revisions),
            SdtContent::Run(run) => collect_run(run, revisions),
            SdtContent::ContentControl(control) => collect_control(control, revisions),
            SdtContent::RawXml(_) => {}
        }
    }
}

fn collect_revision<'a>(revision: &'a CT_Revision, revisions: &mut Vec<&'a CT_Revision>) {
    revisions.push(revision);
    if let Some(paragraph) = revision.content_paragraph() {
        collect_paragraph(paragraph, revisions);
        return;
    }
    if let RevisionContent::Runs(runs) = revision.content() {
        for boundary in 0..=runs.len() {
            for (_, nested) in revision
                .nested_revisions
                .iter()
                .filter(|(at, _)| *at == boundary)
            {
                collect_revision(nested, revisions);
            }
            if let Some(run) = runs.get(boundary) {
                collect_run(run, revisions);
            }
        }
    } else {
        for (_, nested) in &revision.nested_revisions {
            collect_revision(nested, revisions);
        }
    }
}

fn parse_accepted_revision_content(
    raw_xml: &[u8],
    word_prefixes: &[String],
) -> crate::Result<CT_P> {
    let copy_bound = crate::document::admit_word_xml(raw_xml, word_prefixes, None)?;
    parse_revision_content_with_facts(raw_xml, word_prefixes, copy_bound)
        .map(|(paragraph, _)| paragraph)
}

fn parse_revision_content_with_facts(
    raw_xml: &[u8],
    word_prefixes: &[String],
    copy_bound: (usize, usize),
) -> crate::Result<(CT_P, RevisionSourceFacts)> {
    let (content_start, content_end) = revision_content_bounds(raw_xml)?;
    let mut paragraph_xml = Vec::from(b"<w:p>".as_slice());
    paragraph_xml.extend_from_slice(&raw_xml[content_start..content_end]);
    paragraph_xml.extend_from_slice(b"</w:p>");
    let mut paragraph_reader = Reader::from_reader(paragraph_xml.as_slice());
    let mut paragraph_buffer = Vec::new();
    match paragraph_reader.read_event_into(&mut paragraph_buffer)? {
        Event::Start(start) => {
            let mut facts = RevisionSourceFacts::with_copy_bound(copy_bound);
            let paragraph = CT_P::from_xml_with_revision_facts(
                &mut paragraph_reader,
                word_prefixes,
                Some(&start),
                Some(&mut facts),
            )?;
            Ok((paragraph, facts))
        }
        _ => Err(crate::OxmlError::MissingElement(
            "insertion content".to_owned(),
        )),
    }
}

fn revision_content_bounds(raw_xml: &[u8]) -> crate::Result<(usize, usize)> {
    let mut reader = Reader::from_reader(raw_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let content_start = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) => break reader.buffer_position() as usize,
            Event::Eof => {
                return Err(crate::OxmlError::MissingElement(
                    "insertion element".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    };
    let content_end = raw_xml
        .iter()
        .enumerate()
        .rev()
        .find_map(|(index, byte)| (*byte == b'<').then_some(index))
        .ok_or_else(|| crate::OxmlError::MissingElement("insertion close tag".to_owned()))?;
    Ok((content_start, content_end))
}

fn ensure_revision_word_prefix_is_mutable(raw_xml: &[u8]) -> crate::Result<()> {
    let mut reader = Reader::from_reader(raw_xml);
    let mut buffer = Vec::new();
    let root = match reader.read_event_into(&mut buffer)? {
        Event::Start(root) => root,
        _ => {
            return Err(crate::OxmlError::MissingElement(
                "revision start".to_owned(),
            ));
        }
    };
    for attribute in root.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() != b"xmlns:w" {
            continue;
        }
        let namespace =
            attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, root.decoder())?;
        if namespace.as_ref() != W_NS {
            return Err(crate::OxmlError::InvalidValue(
                "revision binds the reserved w prefix to a foreign namespace".to_owned(),
            ));
        }
    }
    Ok(())
}

pub(crate) fn revision_kind(name: &[u8], prefixes: &[String]) -> Option<RevisionKind> {
    [
        (b"ins".as_slice(), RevisionKind::Insertion),
        (b"del".as_slice(), RevisionKind::Deletion),
        (b"moveFrom".as_slice(), RevisionKind::MoveFrom),
        (b"moveTo".as_slice(), RevisionKind::MoveTo),
        (b"rPrChange".as_slice(), RevisionKind::RunPropertyChange),
        (
            b"pPrChange".as_slice(),
            RevisionKind::ParagraphPropertyChange,
        ),
        (b"tblPrChange".as_slice(), RevisionKind::TablePropertyChange),
        (
            b"sectPrChange".as_slice(),
            RevisionKind::SectionPropertyChange,
        ),
    ]
    .into_iter()
    .find_map(|(local, kind)| is_word_element(name, local, prefixes).then_some(kind))
}

fn parse_content(
    reader: &mut Reader<&[u8]>,
    kind: RevisionKind,
    word_prefixes: &[String],
) -> crate::Result<(RevisionContent, Vec<(usize, CT_Revision)>)> {
    let mut runs = Vec::new();
    let mut nested_revisions = Vec::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                let name = start.name();
                if is_word_element(name.as_ref(), b"r", &prefixes) {
                    runs.push(CT_R::from_xml_with_prefixes_and_root(
                        reader,
                        &prefixes,
                        Some(&start),
                    )?);
                } else if revision_kind(name.as_ref(), &prefixes).is_some() {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                } else if is_word_element(name.as_ref(), b"hyperlink", &prefixes) {
                    parse_run_content_container(
                        reader,
                        b"hyperlink",
                        &prefixes,
                        &mut runs,
                        &mut nested_revisions,
                    )?;
                } else if is_word_element(name.as_ref(), b"rPr", &prefixes)
                    && kind == RevisionKind::RunPropertyChange
                {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    return Ok((
                        RevisionContent::PriorRunProperties(Box::new(parse_scoped_rpr(
                            &raw,
                            word_prefixes,
                        )?)),
                        nested_revisions,
                    ));
                } else if is_word_element(name.as_ref(), b"pPr", &prefixes)
                    && kind == RevisionKind::ParagraphPropertyChange
                {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    return Ok((
                        RevisionContent::PriorParagraphProperties(Box::new(parse_scoped_ppr(
                            &raw,
                            word_prefixes,
                        )?)),
                        nested_revisions,
                    ));
                } else if is_word_element(name.as_ref(), b"tblPr", &prefixes)
                    && kind == RevisionKind::TablePropertyChange
                {
                    let owner_bindings =
                        crate::numbering::local_namespace_overrides(&start, word_prefixes)?;
                    return Ok((
                        RevisionContent::PriorTableProperties(Box::new(
                            CT_TblPr::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &owner_bindings,
                            )?,
                        )),
                        nested_revisions,
                    ));
                } else if is_word_element(name.as_ref(), b"sectPr", &prefixes)
                    && kind == RevisionKind::SectionPropertyChange
                {
                    let owner_bindings =
                        crate::numbering::local_namespace_overrides(&start, word_prefixes)?;
                    return Ok((
                        RevisionContent::PriorSectionProperties(Box::new(
                            CT_SectPr::from_xml_with_prefixes_owner_bindings_and_root(
                                reader,
                                &prefixes,
                                &owner_bindings,
                                Some(&start),
                            )?,
                        )),
                        nested_revisions,
                    ));
                } else {
                    reader.read_to_end_into(name, &mut Vec::new())?;
                }
            }
            Event::Empty(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                if revision_kind(start.name().as_ref(), &prefixes).is_some() {
                    let raw = capture_empty_element(&start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                }
            }
            Event::End(_) | Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if runs.is_empty() {
        Ok((RevisionContent::Marker, nested_revisions))
    } else {
        Ok((RevisionContent::Runs(runs), nested_revisions))
    }
}

fn parse_run_content_container(
    reader: &mut Reader<&[u8]>,
    closing_local: &[u8],
    word_prefixes: &[String],
    runs: &mut Vec<CT_R>,
    nested_revisions: &mut Vec<(usize, CT_Revision)>,
) -> crate::Result<()> {
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                let name = start.name();
                if is_word_element(name.as_ref(), b"r", &prefixes) {
                    runs.push(CT_R::from_xml_with_prefixes_and_root(
                        reader,
                        &prefixes,
                        Some(&start),
                    )?);
                } else if revision_kind(name.as_ref(), &prefixes).is_some() {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                } else if is_word_element(name.as_ref(), b"hyperlink", &prefixes) {
                    parse_run_content_container(
                        reader,
                        b"hyperlink",
                        &prefixes,
                        runs,
                        nested_revisions,
                    )?;
                } else {
                    reader.read_to_end_into(name, &mut Vec::new())?;
                }
            }
            Event::Empty(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                if revision_kind(start.name().as_ref(), &prefixes).is_some() {
                    let raw = capture_empty_element(&start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                }
            }
            Event::End(end)
                if is_word_element(end.name().as_ref(), closing_local, word_prefixes) =>
            {
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn required_word_attribute(
    start: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> crate::Result<String> {
    optional_word_attribute(start, local, prefixes)?.ok_or_else(|| {
        crate::OxmlError::MissingElement(format!("w:{} attribute", String::from_utf8_lossy(local)))
    })
}

fn optional_word_attribute(
    start: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> crate::Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let Some(separator) = key.iter().position(|byte| *byte == b':') else {
            continue;
        };
        if key.get(separator + 1..) == Some(local)
            && prefixes
                .iter()
                .any(|prefix| prefix.as_bytes() == &key[..separator])
        {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::namespace::W_NS;
    use crate::text::RunContent;

    fn revision_identity_context() -> Vec<String> {
        vec![
            crate::numbering::namespace_binding("w", W_NS),
            "w".to_owned(),
            crate::numbering::namespace_binding("v", "urn:producer"),
            crate::numbering::namespace_binding("", ""),
        ]
    }

    #[test]
    fn revision_identity_reordered_context_preserves_marker_and_nested_projection() {
        let context = revision_identity_context();
        let mut reordered = context.clone();
        reordered.rotate_left(2);
        let mut outcomes = Vec::new();
        for (label, raw) in [
            ("empty marker", br#"<w:ins w:id="7" w:author="Ada"/>"#.as_slice()),
            ("nested projection", br#"<w:ins w:id="7" w:author="Ada"><w:r><w:t>kept</w:t><v:opaque v:flag="preserved"/></w:r><w:del w:id="8" w:author="Ben"><w:r><w:delText>old</w:delText></w:r></w:del></w:ins>"#.as_slice()),
        ] {
            let left = CT_Revision::from_raw(raw.to_vec(), &context).unwrap();
            let right = CT_Revision::from_raw(raw.to_vec(), &reordered).unwrap();
            assert_eq!(left.parse_word_prefixes, context);
            assert_eq!(right.parse_word_prefixes, reordered);
            assert_eq!(left.raw_xml, raw);
            assert_eq!(right.raw_xml, raw);
            let mut writer = quick_xml::Writer::new(Vec::new());
            left.write_xml(&mut writer).unwrap();
            assert_eq!(writer.into_inner(), raw);
            if label == "nested projection" {
                assert_eq!(left.nested_revisions().len(), 1);
                assert_eq!(left.nested_revisions()[0].1.author(), "Ben");
                assert_eq!(left.content_paragraph().unwrap().runs[0].text(), "kept");
            }
            outcomes.push((label, left == right));
        }
        assert!(outcomes.iter().all(|(_, equal)| *equal), "{outcomes:?}");
    }

    #[test]
    fn revision_identity_keeps_complete_namespace_values_and_multiplicity() {
        let context = revision_identity_context();
        let raw = br#"<w:ins w:id="7" w:author="Ada"><w:r><w:t>kept</w:t><v:opaque v:flag="preserved"/></w:r></w:ins>"#;
        let baseline = CT_Revision::from_raw(raw.to_vec(), &context).unwrap();
        for namespace in ["urn:other-producer", W_NS] {
            let mut changed = context.clone();
            changed[2] = crate::numbering::namespace_binding("v", namespace);
            if namespace == W_NS {
                changed.push("v".to_owned());
            }
            let changed = CT_Revision::from_raw(raw.to_vec(), &changed).unwrap();
            assert_eq!(changed.raw_xml, baseline.raw_xml);
            assert_ne!(changed, baseline);
        }
        for default in [Some("urn:default"), None] {
            let mut changed = context.clone();
            changed.pop();
            if let Some(default) = default {
                changed.push(crate::numbering::namespace_binding("", default));
            }
            assert_ne!(
                CT_Revision::from_raw(raw.to_vec(), &changed).unwrap(),
                baseline
            );
        }
        let mut duplicated = context.clone();
        duplicated.push(context[2].clone());
        assert_ne!(
            CT_Revision::from_raw(raw.to_vec(), &duplicated).unwrap(),
            baseline
        );
    }

    #[test]
    fn revision_identity_preserves_physical_projection_edit_and_parse_policy() {
        let context = revision_identity_context();
        let raw =
            br#"<w:del w:id="7" w:author="Ada"><w:r><w:delText>old</w:delText></w:r></w:del>"#;
        let baseline = CT_Revision::from_raw(raw.to_vec(), &context).unwrap();
        let mut physical = baseline.clone();
        physical
            .physical_content_paragraph_mut()
            .unwrap()
            .unwrap()
            .runs[0] = CT_R::new("physical");
        assert_eq!(physical.raw_xml, baseline.raw_xml);
        assert_eq!(physical.parse_word_prefixes, context);
        assert_ne!(physical, baseline);
        assert_ne!(
            CT_Revision::from_raw_content(raw.to_vec(), &context).unwrap(),
            baseline
        );
        let insertion = br#"<w:ins w:id="7" w:author="Ada"><w:r><w:t>kept</w:t><v:opaque v:flag="preserved"/></w:r></w:ins>"#;
        let original = CT_Revision::from_raw(insertion.to_vec(), &context).unwrap();
        let mut edited = original.clone();
        assert!(
            edited
                .edit_accepted_paragraph(|paragraph| {
                    paragraph.runs[0] = CT_R::new("edited");
                    Ok(true)
                })
                .unwrap()
        );
        assert_eq!(edited.parse_word_prefixes, context);
        assert_eq!(edited.content_paragraph().unwrap().runs[0].text(), "edited");
        assert_ne!(edited, original);
        assert_eq!(original.raw_xml, insertion);
        let changed_opaque = String::from_utf8(insertion.to_vec())
            .unwrap()
            .replace("preserved", "changed");
        assert_ne!(
            CT_Revision::from_raw(changed_opaque.into_bytes(), &context).unwrap(),
            original
        );
    }

    fn nested_insertions(depth: usize) -> Vec<u8> {
        let mut xml = format!(r#"<w:ins xmlns:w="{W_NS}" w:id="0" w:author="Ada">"#);
        for id in 1..depth {
            xml.push_str(&format!(r#"<w:ins w:id="{id}" w:author="Ada">"#));
        }
        xml.push_str("<w:r><w:t>visible</w:t></w:r>");
        for _ in 0..depth {
            xml.push_str("</w:ins>");
        }
        xml.into_bytes()
    }

    #[test]
    fn revision_nesting_depth_is_bounded_before_recursive_projection() {
        let prefixes = ["w".to_owned()];
        assert!(
            validate_revision_nesting_depth(
                &nested_insertions(MAX_REVISION_NESTING_DEPTH),
                &prefixes,
            )
            .is_ok()
        );
        let too_deep = nested_insertions(MAX_REVISION_NESTING_DEPTH + 1);
        assert!(validate_revision_nesting_depth(&too_deep, &prefixes).is_err());
        assert!(CT_Revision::from_raw(too_deep, &prefixes).is_none());
    }

    #[test]
    fn revision_attributes_are_prefix_tolerant_and_namespace_checked() {
        let raw = format!(
            r#"<word:ins xmlns:word="{W_NS}" word:id="7" word:author="Ada"><word:r><word:t>added</word:t></word:r></word:ins>"#
        )
        .into_bytes();
        let revision = CT_Revision::from_raw(raw, &["word".to_owned()])
            .expect("aliased WordprocessingML revision should be projected");
        assert_eq!(revision.id(), 7);
        assert_eq!(revision.author(), "Ada");
        assert_eq!(revision.timestamp(), None);
        assert_eq!(revision.kind(), RevisionKind::Insertion);

        let escaped = format!(
            r#"<word:ins xmlns:word="{W_NS}" word:id="8" word:author="A &amp; B" word:date="2026-08-17T09:00:00+01:00 &amp; later"/>"#
        )
        .into_bytes();
        let escaped = CT_Revision::from_raw(escaped, &["word".to_owned()])
            .expect("escaped revision metadata should parse");
        assert_eq!(escaped.author(), "A & B");
        assert_eq!(
            escaped.timestamp(),
            Some("2026-08-17T09:00:00+01:00 & later")
        );

        let foreign = br#"<x:ins xmlns:x="urn:not-word" x:id="9" x:author="Eve"/>"#.to_vec();
        assert!(CT_Revision::from_raw(foreign, &["word".to_owned()]).is_none());

        let missing_author = format!(r#"<word:del xmlns:word="{W_NS}" word:id="9"/>"#).into_bytes();
        assert!(CT_Revision::from_raw(missing_author, &["word".to_owned()]).is_none());

        let document_xml = format!(
            r#"<word:document xmlns:word="{W_NS}" xmlns:x="urn:not-word"><word:body><word:p><word:pPr><word:pPrChange word:id="10" word:author="Gia"><word:pPr><word:jc word:val="right"/></word:pPr></word:pPrChange><x:pPrChange x:id="11" x:author="No"/><word:pPrChange word:id="12"/></word:pPr></word:p></word:body></word:document>"#
        );
        let document = CT_Document::from_xml(document_xml.as_bytes()).expect("document parses");
        assert_eq!(document.revisions().len(), 1);
        assert_eq!(document.revisions()[0].id(), 10);
        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        assert!(output.contains(r#"<word:pPrChange word:id="10" word:author="Gia">"#));
        assert!(output.contains(r#"<x:pPrChange x:id="11" x:author="No"/>"#));
        assert!(output.contains(r#"<word:pPrChange word:id="12"/>"#));

        let collision_xml = format!(
            r#"<word:document xmlns:word="{W_NS}" xmlns:x="urn:not-word"><word:body>
<word:tbl><word:tblPr><word:tblPrChange word:id="20" word:author="Tab"><word:tblPr><x:jc x:val="center"/><word:jc x:val="right"/><word:tblBorders><x:top x:val="single" word:sz="invalid"/><word:bottom x:val="single"/></word:tblBorders><word:tblCellMar><x:top x:w="120"/><word:bottom x:w="240"/></word:tblCellMar></word:tblPr></word:tblPrChange></word:tblPr><word:tblGrid/><word:tr><word:tc><word:p/></word:tc></word:tr></word:tbl>
<word:sectPr><word:sectPrChange word:id="21" word:author="Sec"><word:sectPr><x:titlePg/><word:pgSz x:w="12240" x:h="15840"/><word:pgMar x:producer="not-a-number" word:top="720"/></word:sectPr></word:sectPrChange></word:sectPr>
</word:body></word:document>"#
        );
        let collision = CT_Document::from_xml(collision_xml.as_bytes()).expect("document parses");
        let collision_revisions = collision.revisions();
        assert!(matches!(
            collision_revisions[0].content(),
            RevisionContent::PriorTableProperties(properties)
                if properties.jc.is_none()
                    && properties
                        .borders
                        .as_ref()
                        .is_some_and(|borders| borders.top.is_none()
                            && borders.bottom.as_ref().is_some_and(|edge| {
                                edge.val == crate::shared::ST_Border::None
                            }))
                    && properties
                        .cell_margin
                        .as_ref()
                        .is_some_and(|margin| margin.top.is_none() && margin.bottom.is_none())
        ));
        assert!(matches!(
            collision_revisions[1].content(),
            RevisionContent::PriorSectionProperties(properties)
                if properties.title_pg.is_none()
                    && properties.page_width.is_none()
                    && properties.page_height.is_none()
                    && properties.margin_top == Some(crate::units::Twips(720))
        ));

        let paragraph_collision_xml = format!(
            r#"<word:document xmlns:word="{W_NS}" xmlns:x="urn:not-word"><word:body><word:p><word:pPr><word:pPrChange word:id="22" word:author="Para"><word:pPr><word:pBdr><x:top x:val="single" word:sz="invalid"/><word:bottom x:val="single"/></word:pBdr></word:pPr></word:pPrChange></word:pPr></word:p></word:body></word:document>"#
        );
        let paragraph_collision =
            CT_Document::from_xml(paragraph_collision_xml.as_bytes()).expect("document parses");
        assert!(matches!(
            paragraph_collision.revisions()[0].content(),
            RevisionContent::PriorParagraphProperties(properties)
                if properties.borders.as_ref().is_some_and(|borders| {
                    borders.top.is_none()
                        && borders.bottom.as_ref().is_some_and(|edge| {
                            edge.val == crate::shared::ST_Border::None
                        })
                })
        ));

        let control_xml = format!(
            r#"<word:document xmlns:word="{W_NS}"><word:body><word:p><word:sdt><word:sdtContent><word:ins word:id="30" word:author="Sdt"><word:r><word:t>inside</word:t></word:r></word:ins></word:sdtContent></word:sdt></word:p></word:body></word:document>"#
        );
        let control = CT_Document::from_xml(control_xml.as_bytes()).expect("document parses");
        assert_eq!(
            control
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            vec![30]
        );
    }

    #[test]
    fn insertion_content_retains_inline_paragraph_structure() {
        let raw = format!(
            r#"<w:ins xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:ins>"#
        )
        .into_bytes();
        let revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("insertion parses");
        let paragraph = revision
            .content_paragraph()
            .expect("insertion exposes paragraph content");

        assert_eq!(paragraph.text(), "beforelinked");
        assert_eq!(paragraph.hyperlinks.len(), 1);
        assert_eq!(paragraph.hyperlinks[0].rel_id.as_deref(), Some("rId9"));
    }

    #[test]
    fn accepted_content_revisions_retain_inline_content_controls() {
        for kind in ["ins", "moveTo"] {
            let raw = format!(
                r#"<w:{kind} xmlns:w="{W_NS}" w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt></w:{kind}>"#
            )
            .into_bytes();
            let revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("revision parses");
            let paragraph = revision
                .content_paragraph()
                .expect("accepted revision exposes paragraph content");

            assert_eq!(paragraph.text(), "beforecontrol");
            assert_eq!(paragraph.content_controls.len(), 1);
            let crate::content_control::SdtContent::Run(run) =
                &paragraph.content_controls[0].3.content[0]
            else {
                panic!("expected controlled run");
            };
            assert_eq!(run.text(), "control");
        }
    }

    #[test]
    fn accepted_revision_mutation_rejects_a_shadowed_reserved_word_prefix() {
        let raw = format!(
            r#"<wx:ins xmlns:wx="{W_NS}" xmlns:w="urn:foreign" wx:id="7" wx:author="Ada"><wx:r><wx:t>before</wx:t></wx:r></wx:ins>"#
        )
        .into_bytes();
        let mut revision =
            CT_Revision::from_raw(raw, &["wx".to_owned()]).expect("insertion parses");
        let error = revision
            .replace_accepted_run_segments(&[AcceptedRunPathSegment::Run(0)], CT_R::new("after"))
            .expect_err("shadowed reserved prefix must fail closed");

        assert_eq!(
            error.to_string(),
            "invalid value: revision binds the reserved w prefix to a foreign namespace"
        );
        assert_eq!(
            revision
                .content_paragraph()
                .expect("accepted content remains available")
                .text(),
            "before"
        );
    }

    #[test]
    fn accepted_revision_refresh_changes_one_field_of_a_shared_run() {
        let raw = format!(
            concat!(
                r#"<w:ins xmlns:w="{}" w:id="7" w:author="Ada"><w:r>"#,
                r#"<w:fldChar w:fldCharType="begin"/><w:instrText>DATE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>first</w:t><w:fldChar w:fldCharType="end"/>"#,
                r#"<w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>second</w:t><w:fldChar w:fldCharType="end"/>"#,
                r#"</w:r></w:ins>"#,
            ),
            W_NS
        )
        .into_bytes();
        let mut revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("insertion parses");
        let paragraph = revision
            .content_paragraph()
            .expect("accepted content is projected");
        let page_index = paragraph
            .runs
            .iter()
            .position(|run| {
                run.content.iter().any(|content| {
                    matches!(content, crate::text::RunContent::Field(field) if field.instruction.name == "PAGE")
                })
            })
            .expect("PAGE field is projected");
        let mut replacement = paragraph.runs[page_index].clone();
        let field = replacement
            .content
            .iter_mut()
            .find_map(|content| match content {
                crate::text::RunContent::Field(field) => Some(field),
                _ => None,
            })
            .expect("replacement contains a field");
        field.cached_result = "changed".to_owned();

        assert!(
            revision
                .replace_accepted_run_segments(
                    &[AcceptedRunPathSegment::Run(page_index)],
                    replacement
                )
                .expect("one field of a shared physical run can change")
        );

        assert_eq!(
            revision
                .content_paragraph()
                .expect("accepted content remains")
                .text(),
            "firstchanged"
        );
        let raw = std::str::from_utf8(&revision.raw_xml).unwrap();
        assert!(raw.contains("<w:instrText>DATE</w:instrText>"), "{raw}");
        assert!(raw.contains("<w:t>first</w:t>"), "{raw}");
        assert!(!raw.contains("second"), "{raw}");
    }

    #[test]
    fn insertion_content_retains_nested_revision_projection() {
        let raw = format!(
            r#"<w:ins xmlns:w="{W_NS}" w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:del w:id="8" w:author="Ben"><w:r><w:delText>removed</w:delText></w:r></w:del><w:ins w:id="9" w:author="Cy"><w:r><w:t>nested</w:t></w:r></w:ins><w:r><w:t>after</w:t></w:r></w:ins>"#
        )
        .into_bytes();
        let revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("insertion parses");

        let RevisionContent::Runs(runs) = revision.content() else {
            panic!("insertion exposes direct runs");
        };
        assert_eq!(
            runs.iter().map(CT_R::text).collect::<String>(),
            "beforeafter"
        );
        assert_eq!(
            revision
                .nested_revisions()
                .iter()
                .map(|(boundary, nested)| (*boundary, nested.kind()))
                .collect::<Vec<_>>(),
            vec![(1, RevisionKind::Deletion), (1, RevisionKind::Insertion),]
        );
        let paragraph = revision
            .content_paragraph()
            .expect("insertion exposes paragraph content");
        assert_eq!(paragraph.text(), "beforeafter");
        assert_eq!(
            paragraph
                .revisions
                .iter()
                .map(|(_, _, nested)| nested.kind())
                .collect::<Vec<_>>(),
            vec![RevisionKind::Deletion, RevisionKind::Insertion]
        );
    }

    #[test]
    fn property_changes_write_in_their_schema_final_slots() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body>
<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>
<w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Ben"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="3" w:author="Cy"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="4" w:author="Dee"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");

        for change in ["pPrChange", "rPrChange", "tblPrChange", "sectPrChange"] {
            assert_eq!(
                output.matches(&format!("<w:{change} ")).count(),
                1,
                "{output}"
            );
        }
        assert!(
            output.find("<w:jc w:val=\"center\"").unwrap() < output.find("<w:pPrChange ").unwrap()
        );
        assert!(output.find("<w:b").unwrap() < output.find("<w:rPrChange ").unwrap());
        assert!(output.find("<w:tblW").unwrap() < output.find("<w:tblPrChange ").unwrap());
        assert!(output.find("<w:pgSz").unwrap() < output.find("<w:sectPrChange ").unwrap());
    }

    #[test]
    fn run_property_raw_positions_cannot_follow_the_schema_final_change() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body><w:p><w:r><w:rPr><x:raw/><w:rPrChange w:id="1" w:author="Ada"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        );
        let mut document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        {
            let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
                panic!("expected paragraph");
            };
            paragraph.runs[0]
                .properties
                .as_mut()
                .expect("run properties")
                .revision_xml_positions[0] = (41, 9);
        }

        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        assert!(output.find("<x:raw").unwrap() < output.find("<w:rPrChange").unwrap());

        {
            let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
                panic!("expected paragraph");
            };
            paragraph.runs[0]
                .properties
                .as_mut()
                .expect("run properties")
                .revision_xml_positions
                .clear();
        }
        let output = String::from_utf8(document.to_xml().expect("fallback writes")).unwrap();
        assert!(output.find("<x:raw").unwrap() < output.find("<w:rPrChange").unwrap());

        let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
            panic!("expected paragraph");
        };
        paragraph.runs[0]
            .properties
            .as_mut()
            .expect("run properties")
            .revision_xml_positions
            .extend([(41, 0), (41, 1)]);
        let output = String::from_utf8(document.to_xml().expect("mismatch writes")).unwrap();
        assert!(output.find("<x:raw").unwrap() < output.find("<w:rPrChange").unwrap());
    }

    #[test]
    fn retained_run_property_children_keep_owner_local_namespace_bindings() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:rPr xmlns:x="urn:producer" xmlns:wa="{W_NS}"><x:raw/><wa:outline><wa:opaque/></wa:outline><wa:rPrChange wa:id="7" wa:author="Ada"><wa:rPr><wa:b/></wa:rPr></wa:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        assert_eq!(document.revisions()[0].id(), 7);

        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        assert_eq!(
            output.matches(r#"<x:raw xmlns:x="urn:producer"/>"#).count(),
            1
        );
        assert_eq!(
            output
                .matches(&format!(r#"<wa:outline xmlns:wa="{W_NS}">"#))
                .count(),
            1
        );
        assert_eq!(
            output
                .matches(&format!(
                    r#"<wa:rPrChange wa:id="7" wa:author="Ada" xmlns:wa="{W_NS}">"#
                ))
                .count(),
            1
        );

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(reopened.revisions()[0].id(), 7);
        let rewritten = String::from_utf8(reopened.to_xml().expect("output rewrites")).unwrap();
        assert_eq!(rewritten.matches(r#"xmlns:x="urn:producer""#).count(), 1);
        assert_eq!(
            rewritten.matches(&format!(r#"xmlns:wa="{W_NS}""#)).count(),
            2
        );
    }

    #[test]
    fn duplicate_property_changes_are_retained_without_replacing_the_first_projection() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body>
<w:p><w:pPr><w:numPr><w:ins w:id="501" w:author="Ada"/><x:ins x:mark="num"/><w:ins w:id="502" w:author="Ada"/></w:numPr><w:pPrChange w:id="101" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange><x:pPrChange x:mark="paragraph"/><w:pPrChange w:id="102" w:author="Ada"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="201" w:author="Ada"><w:rPr><w:b/></w:rPr></w:rPrChange><x:rPrChange x:mark="run"/><w:rPrChange w:id="202" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblPrChange w:id="301" w:author="Ada"><w:tblPr><w:jc w:val="left"/></w:tblPr></w:tblPrChange><x:tblPrChange x:mark="table"/><w:tblPrChange w:id="302" w:author="Ada"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr><w:sectPrChange w:id="401" w:author="Ada"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange><x:sectPrChange x:mark="section"/><w:sectPrChange w:id="402" w:author="Ada"><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:sectPrChange></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let mut ids = document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        ids.sort_unstable();
        assert_eq!(ids, [102, 202, 302, 402, 502]);

        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        for id in [101, 102, 201, 202, 301, 302, 401, 402, 501, 502] {
            assert_eq!(output.matches(&format!(r#"w:id="{id}""#)).count(), 1);
        }
        for owner in ["pPr", "rPr", "tblPr", "sectPr"] {
            let first = output
                .find(&format!(
                    r#"w:{owner}Change w:id="{}""#,
                    match owner {
                        "pPr" => 101,
                        "rPr" => 201,
                        "tblPr" => 301,
                        _ => 401,
                    }
                ))
                .unwrap();
            let typed = output
                .find(&format!(
                    r#"w:{owner}Change w:id="{}""#,
                    match owner {
                        "pPr" => 102,
                        "rPr" => 202,
                        "tblPr" => 302,
                        _ => 402,
                    }
                ))
                .unwrap();
            assert!(
                first < typed,
                "duplicate must precede typed schema-final change"
            );
            let foreign = output
                .find(&format!(r#"<x:{owner}Change x:mark="#))
                .unwrap();
            assert!(first < foreign && foreign < typed);
        }
        let numbering_positions = [
            output.find(r#"w:id="501""#).unwrap(),
            output.find("<x:ins ").unwrap(),
            output.find(r#"w:id="502""#).unwrap(),
        ];
        assert!(output.contains(r#"x:mark="num""#));
        assert!(numbering_positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        let mut reopened_ids = reopened
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        reopened_ids.sort_unstable();
        assert_eq!(reopened_ids, [102, 202, 302, 402, 502]);
    }

    #[test]
    fn every_property_revision_keeps_owner_local_aliases() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body>
<w:p><w:pPr xmlns:pa="{W_NS}" xmlns:px="urn:paragraph"><w:numPr xmlns:na="{W_NS}" xmlns:nx="urn:numbering"><na:ins na:id="501" na:author="Ada"/><nx:ins nx:mark="raw"/><na:ins na:id="502" na:author="Ada"/></w:numPr><pa:pPrChange pa:id="101" pa:author="Ada"><pa:pPr/></pa:pPrChange><px:pPrChange px:mark="raw"/><pa:pPrChange pa:id="102" pa:author="Ada"><pa:pPr/></pa:pPrChange></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
<w:tbl><w:tblPr xmlns:ta="{W_NS}" xmlns:tx="urn:table"><ta:tblPrChange ta:id="301" ta:author="Ada"><ta:tblPr/></ta:tblPrChange><tx:tblPrChange tx:mark="raw"/><ta:tblPrChange ta:id="302" ta:author="Ada"><ta:tblPr/></ta:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr xmlns:sa="{W_NS}" xmlns:sx="urn:section"><sa:sectPrChange sa:id="401" sa:author="Ada"><sa:sectPr/></sa:sectPrChange><sx:sectPrChange sx:mark="raw"/><sa:sectPrChange sa:id="402" sa:author="Ada"><sa:sectPr/></sa:sectPrChange></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        for (prefix, local) in [
            ("pa", "pPrChange"),
            ("na", "ins"),
            ("ta", "tblPrChange"),
            ("sa", "sectPrChange"),
        ] {
            assert_eq!(output.matches(&format!("<{prefix}:{local} ")).count(), 2);
            assert_eq!(
                output
                    .matches(&format!(r#"xmlns:{prefix}="{W_NS}""#))
                    .count(),
                if prefix == "na" { 1 } else { 2 }
            );
        }
        for (prefix, namespace, local) in [
            ("px", "urn:paragraph", "pPrChange"),
            ("nx", "urn:numbering", "ins"),
            ("tx", "urn:table", "tblPrChange"),
            ("sx", "urn:section", "sectPrChange"),
        ] {
            assert_eq!(output.matches(&format!("<{prefix}:{local} ")).count(), 1);
            assert_eq!(
                output
                    .matches(&format!(r#"xmlns:{prefix}="{namespace}""#))
                    .count(),
                1
            );
        }
        for (first, foreign, second) in [
            (101, "<px:pPrChange", 102),
            (501, "<nx:ins", 502),
            (301, "<tx:tblPrChange", 302),
            (401, "<sx:sectPrChange", 402),
        ] {
            let positions = [
                output.find(&format!(r#"id="{first}""#)).unwrap(),
                output.find(foreign).unwrap(),
                output.find(&format!(r#"id="{second}""#)).unwrap(),
            ];
            assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
        }

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        let mut ids = reopened
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        ids.sort_unstable();
        assert_eq!(ids, [102, 302, 402, 502]);
        let rewritten = String::from_utf8(reopened.to_xml().expect("output rewrites")).unwrap();
        for prefix in ["pa", "na", "ta", "sa"] {
            assert_eq!(
                rewritten
                    .matches(&format!(r#"xmlns:{prefix}="{W_NS}""#))
                    .count(),
                if prefix == "na" { 1 } else { 2 }
            );
        }
        for (prefix, namespace) in [
            ("px", "urn:paragraph"),
            ("nx", "urn:numbering"),
            ("tx", "urn:table"),
            ("sx", "urn:section"),
        ] {
            assert_eq!(
                rewritten
                    .matches(&format!(r#"xmlns:{prefix}="{namespace}""#))
                    .count(),
                1
            );
        }
    }

    #[test]
    fn a_section_root_retains_no_namespace_declaration_its_attributes_do_not_use() {
        // F-X128 retained every attribute on a modeled root, declarations
        // included, so a binding used only by a child was written twice, once
        // on `w:sectPr` and once on the child that already carried it.
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p/>
<w:sectPr xmlns:sa="{W_NS}" xmlns:sx="urn:section"><sa:sectPrChange sa:id="401" sa:author="Ada"><sa:sectPr/></sa:sectPrChange><sx:sectPrChange sx:mark="raw"/></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();

        // The declaration reaches the one child that uses it, and no more.
        assert_eq!(
            output.matches(&format!(r#"xmlns:sa="{W_NS}""#)).count(),
            1,
            "{output}"
        );
        assert_eq!(
            output.matches(r#"xmlns:sx="urn:section""#).count(),
            1,
            "{output}"
        );
        assert!(
            !output.contains(&format!(r#"<w:sectPr xmlns:sa="{W_NS}""#)),
            "the section root kept a declaration none of its attributes use: {output}"
        );
        assert!(output.contains(r#"sa:id="401""#), "{output}");
        assert!(output.contains(r#"sx:mark="raw""#), "{output}");
    }

    #[test]
    fn revision_elements_round_trip_unchanged_and_report_metadata() {
        let revisions = [
            r#"<w:ins w:id="1" w:author="Ada" w:date="2026-08-01T10:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>"#,
            r#"<w:del w:id="2" w:author="Ben"><w:r><w:delText xml:space="preserve"> gone </w:delText></w:r></w:del>"#,
            r#"<w:moveFrom w:id="3" w:author="Cy"><w:r><w:t>from</w:t></w:r></w:moveFrom>"#,
            r#"<w:moveTo w:id="4" w:author="Dee"><w:r><w:t>to</w:t></w:r></w:moveTo>"#,
        ];
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p>{}<w:r><w:rPr><w:ins w:id="5" w:author="Eve"/><w:rPrChange w:id="6" w:author="Fox"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>end</w:t></w:r><x:ins xmlns:x="urn:not-word" x:id="99" x:author="No"/></w:p><w:tbl><w:tblPr><w:tblPrChange w:id="7" w:author="Gia"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:trPr><w:del w:id="8" w:author="Hal"/></w:trPr><w:tc><w:p><w:pPr><w:pPrChange w:id="9" w:author="Ivy"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:sectPrChange w:id="10" w:author="Jay"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>"#,
            revisions.concat()
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let reported = document.revisions();
        assert_eq!(
            reported
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            (1..=10).collect::<Vec<_>>()
        );
        assert_eq!(reported[0].author(), "Ada");
        assert_eq!(reported[0].timestamp(), Some("2026-08-01T10:00:00Z"));
        assert_eq!(reported[1].timestamp(), None);
        assert_eq!(reported[1].kind(), RevisionKind::Deletion);
        let RevisionContent::Runs(runs) = reported[1].content() else {
            panic!("deletion should project runs");
        };
        assert!(matches!(
            runs[0].content.as_slice(),
            [RunContent::DeletedText(text)] if text.text == " gone " && text.preserve_space
        ));
        assert!(matches!(
            reported[5].content(),
            RevisionContent::PriorRunProperties(properties) if properties.italic == Some(true)
        ));
        assert!(matches!(
            reported[6].content(),
            RevisionContent::PriorTableProperties(properties) if properties.jc.is_some()
        ));
        assert!(matches!(
            reported[8].content(),
            RevisionContent::PriorParagraphProperties(properties) if properties.jc.is_some()
        ));
        assert!(matches!(
            reported[9].content(),
            RevisionContent::PriorSectionProperties(properties) if properties.title_pg == Some(true)
        ));

        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        for raw in revisions {
            assert!(
                output.contains(raw),
                "missing exact subtree {raw} in {output}"
            );
        }
        for raw in [
            r#"<w:rPrChange w:id="6" w:author="Fox"><w:rPr><w:i/></w:rPr></w:rPrChange>"#,
            r#"<w:tblPrChange w:id="7" w:author="Gia"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:tblPrChange>"#,
            r#"<w:pPrChange w:id="9" w:author="Ivy"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>"#,
            r#"<w:sectPrChange w:id="10" w:author="Jay"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange>"#,
        ] {
            assert!(
                output.contains(raw),
                "missing exact subtree {raw} in {output}"
            );
        }
        assert!(output.contains(r#"<x:ins xmlns:x="urn:not-word" x:id="99" x:author="No"/>"#));
        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(reopened.revisions().len(), 10);
    }

    #[test]
    fn hyperlink_and_nested_content_revisions_round_trip_and_report_in_order() {
        let nested = r#"<w:ins w:id="11" w:author="Ada"><w:del w:id="12" w:author="Ben"><w:r><w:delText>nested</w:delText></w:r></w:del></w:ins>"#;
        let hyperlink_before = r#"<w:hyperlink><w:ins w:id="21" w:author="Cy"><w:r><w:t>before</w:t></w:r></w:ins></w:hyperlink>"#;
        let direct_after =
            r#"<w:del w:id="22" w:author="Dee"><w:r><w:delText>after</w:delText></w:r></w:del>"#;
        let direct_before =
            r#"<w:ins w:id="23" w:author="Eve"><w:r><w:t>before</w:t></w:r></w:ins>"#;
        let hyperlink_after = r#"<w:hyperlink><w:del w:id="24" w:author="Fox"><w:r><w:delText>after</w:delText></w:r></w:del></w:hyperlink>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rId5"><w:r><w:t xml:space="preserve">before </w:t></w:r>{nested}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:hyperlink></w:p><w:p>{hyperlink_before}{direct_after}{direct_before}{hyperlink_after}</w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        assert_eq!(
            document
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            [11, 12, 21, 22, 23, 24]
        );
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        assert!(output.contains(nested), "missing exact subtree in {output}");
        for raw in [
            hyperlink_before,
            direct_after,
            direct_before,
            hyperlink_after,
        ] {
            assert!(
                output.contains(raw),
                "missing exact subtree {raw} in {output}"
            );
        }
        let positions = [21, 22, 23, 24].map(|id| {
            output
                .find(&format!(r#"w:id="{id}""#))
                .expect("revision id remains present")
        });
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(
            reopened
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            [11, 12, 21, 22, 23, 24]
        );
    }

    #[test]
    fn modeled_hyperlinks_preserve_unreported_raw_children_and_foreign_owners() {
        let malformed = r#"<w:ins w:id="bad"><w:r><w:t>raw revision</w:t></w:r></w:ins>"#;
        let foreign_child = r#"<x:opaque x:flag="1"/>"#;
        let foreign_hyperlink = r#"<x:hyperlink xmlns:x="urn:not-word"><w:r><w:t>foreign owner</w:t></w:r></x:hyperlink>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rId5" xmlns:x="urn:opaque"><w:r><w:t>before</w:t></w:r>{malformed}<w:ins w:id="31" w:author="Ada"><w:r><w:t>reported</w:t></w:r></w:ins>{foreign_child}<w:r><w:t>after</w:t></w:r></w:hyperlink>{foreign_hyperlink}</w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        assert_eq!(
            document
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            [31]
        );
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        assert!(
            output.contains(malformed),
            "missing malformed subtree in {output}"
        );
        assert!(
            output.contains(foreign_child),
            "missing foreign child in {output}"
        );
        assert!(
            output.contains(foreign_hyperlink),
            "foreign owner was promoted or lost in {output}"
        );
        let positions = [
            output.find("<w:t>before</w:t>").unwrap(),
            output.find(malformed).unwrap(),
            output.find(r#"w:id="31""#).unwrap(),
            output.find(foreign_child).unwrap(),
            output.find("<w:t>after</w:t>").unwrap(),
        ];
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(reopened.revisions().len(), 1);
        assert_eq!(reopened.revisions()[0].id(), 31);
    }

    #[test]
    fn aliased_hyperlink_runs_survive_a_locally_shadowed_word_prefix() {
        let foreign_run_properties = r#"<w:rPr/>"#;
        let foreign_text = r#"<w:t>foreign</w:t>"#;
        let foreign_drawing = r#"<w:drawing><w:opaque/></w:drawing>"#;
        let foreign_after = r#"<w:after/>"#;
        let lexical_literal = r#"<f:literal xmlns:f="urn:foreign" f:value="w: >">w:</f:literal>"#;
        let nested_word =
            r#"<f:box xmlns:f="urn:foreign" xmlns:w2="urn:two" f:value=">"><w:nested/></f:box>"#;
        let raw_at_boundary = r#"<w:boundary/>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:wx="{W_NS}"><w:body><w:p><wx:hyperlink xmlns:w="urn:foreign"><wx:r>{foreign_run_properties}<wx:rPr><wx:rStyle><wx:opaque/></wx:rStyle><w:pre/><wx:b wx:val="0"/><w:betweenDuplicate/><wx:b/><w:b/><w:between/><wx:i/><w:beforeChange/><wx:rPrChange wx:id="62" wx:author="Ada"><wx:rPr><wx:b/></wx:rPr></wx:rPrChange><w:afterChange/></wx:rPr>{foreign_text}<wx:t>before</wx:t>{foreign_drawing}<wx:t>middle</wx:t>{foreign_after}{lexical_literal}{nested_word}</wx:r><wx:ins wx:id="61" wx:author="Ada"><wx:r><wx:t>reported</wx:t></wx:r></wx:ins>{raw_at_boundary}<wx:r><wx:t>after</wx:t></wx:r></wx:hyperlink></w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        assert_eq!(
            document.body.paragraphs().next().unwrap().text(),
            "beforemiddleafter"
        );
        assert_eq!(
            document
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            vec![62, 61]
        );
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        assert!(output.contains(r#"xmlns:w="urn:foreign""#));
        let foreign_run_properties = r#"<w:rPr xmlns:w="urn:foreign"/>"#;
        let foreign_property = r#"<w:b xmlns:w="urn:foreign"/>"#;
        let foreign_before_property = r#"<w:pre xmlns:w="urn:foreign"/>"#;
        let foreign_between_duplicate = r#"<w:betweenDuplicate xmlns:w="urn:foreign"/>"#;
        let foreign_between_property = r#"<w:between xmlns:w="urn:foreign"/>"#;
        let foreign_before_change = r#"<w:beforeChange xmlns:w="urn:foreign"/>"#;
        let foreign_after_change = r#"<w:afterChange xmlns:w="urn:foreign"/>"#;
        let foreign_text = r#"<w:t xmlns:w="urn:foreign">foreign</w:t>"#;
        let foreign_drawing = r#"<w:drawing xmlns:w="urn:foreign"><w:opaque/></w:drawing>"#;
        let foreign_after = r#"<w:after xmlns:w="urn:foreign"/>"#;
        for raw in [
            foreign_run_properties,
            foreign_before_property,
            foreign_between_duplicate,
            foreign_property,
            foreign_between_property,
            foreign_before_change,
            foreign_after_change,
            foreign_text,
            foreign_drawing,
            foreign_after,
        ] {
            assert_eq!(output.matches(raw).count(), 1, "raw child {raw}");
        }
        assert_eq!(output.matches(lexical_literal).count(), 1);
        assert!(!output.contains(r#"<f:literal xmlns:w="urn:foreign""#));
        let nested_word = r#"<f:box xmlns:f="urn:foreign" xmlns:w2="urn:two" f:value=">" xmlns:w="urn:foreign"><w:nested/></f:box>"#;
        assert_eq!(output.matches(nested_word).count(), 1);
        assert!(output.contains(raw_at_boundary));
        assert!(output.contains(&format!(r#"<w:r xmlns:w="{W_NS}">"#)));
        let positions = [
            output.find(foreign_run_properties).unwrap(),
            output.find("<wx:rStyle>").unwrap(),
            output.find(foreign_before_property).unwrap(),
            output.find(foreign_between_duplicate).unwrap(),
            output.find("<w:b/>").unwrap(),
            output.find(foreign_property).unwrap(),
            output.find(foreign_between_property).unwrap(),
            output.find("<w:i/>").unwrap(),
            output.find(foreign_before_change).unwrap(),
            output.find(foreign_after_change).unwrap(),
            output.find(r#"wx:id="62""#).unwrap(),
            output.find(foreign_text).unwrap(),
            output.find("<w:t>before</w:t>").unwrap(),
            output.find(foreign_drawing).unwrap(),
            output.find("<w:t>middle</w:t>").unwrap(),
            output.find(foreign_after).unwrap(),
            output.find(r#"wx:id="61""#).unwrap(),
            output.find(raw_at_boundary).unwrap(),
            output.find("<w:t>after</w:t>").unwrap(),
        ];
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(
            reopened.body.paragraphs().next().unwrap().text(),
            "beforemiddleafter"
        );
        assert_eq!(
            reopened
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            vec![62, 61]
        );
    }

    #[test]
    fn issue425_ordered_result_current_wrappers_and_bounded_construction() {
        for (depth, siblings) in [(1, 1), (2, 1), (4, 1), (6, 1), (12, 1), (4, 3)] {
            for prefix in ["w:", "q:", ""] {
                let mut payload = format!("<{prefix}r><{prefix}t>cache</{prefix}t></{prefix}r>");
                for level in 0..depth {
                    payload = format!(
                        r#"<{prefix}fldSimple {prefix}instr="DATE"><{prefix}ins {prefix}id="{level}" {prefix}author="Prior"><{prefix}sdt><{prefix}sdtPr/><{prefix}sdtContent>{payload}</{prefix}sdtContent></{prefix}sdt></{prefix}ins></{prefix}fldSimple>"#
                    );
                }
                // Attributes require a namespace prefix even for default-Word roots.
                if prefix.is_empty() {
                    payload = payload
                        .replace(" instr=", " q:instr=")
                        .replace(" id=", " q:id=")
                        .replace(" author=", " q:author=");
                }
                let declarations = if prefix == "w:" {
                    format!(r#"xmlns:w="{W_NS}""#)
                } else {
                    format!(r#"xmlns:q="{W_NS}" xmlns="{W_NS}" xmlns:w="urn:foreign""#)
                };
                let xml = format!(
                    "<{prefix}document {declarations}><{prefix}body><{prefix}p>{}</{prefix}p></{prefix}body></{prefix}document>",
                    payload.repeat(siblings)
                );
                crate::text::ORDERED_DIAGNOSTIC_SIMPLE_PARSE_CALLS.set(0);
                ORDERED_DIAGNOSTIC_REVISION_PARSE_CALLS.set(0);
                ORDERED_DIAGNOSTIC_REVISION_PARSE_BYTES.set(0);
                let model = CT_Document::from_xml(xml.as_bytes()).unwrap();
                let fields = crate::text::ORDERED_DIAGNOSTIC_SIMPLE_PARSE_CALLS.get();
                let revisions = ORDERED_DIAGNOSTIC_REVISION_PARSE_CALLS.get();
                let revision_bytes = ORDERED_DIAGNOSTIC_REVISION_PARSE_BYTES.get();
                assert_eq!(
                    fields,
                    depth * siblings,
                    "one actual recursive Field construction per source owner: {depth}/{siblings}/{prefix:?}"
                );
                assert_eq!(
                    revisions,
                    depth * siblings,
                    "legacy cache must derive from the canonical revision construction"
                );
                assert!(revision_bytes <= depth.checked_mul(xml.len()).unwrap());
                let current = model.revisions();
                assert_eq!(current.len(), depth * siblings);
                assert_eq!(
                    current
                        .iter()
                        .map(|revision| revision.id())
                        .collect::<Vec<_>>(),
                    (0..siblings)
                        .flat_map(|_| (0..depth).rev().map(|level| level as i32))
                        .collect::<Vec<_>>()
                );
                assert!(current.iter().all(|revision| {
                    revision.content_paragraph().unwrap().accepted_text() == "cache"
                }));
                let mut secondary_nodes = 0usize;
                let mut pending = current.to_vec();
                while let Some(revision) = pending.pop() {
                    secondary_nodes = secondary_nodes.checked_add(1).unwrap();
                    pending.extend(revision.nested_revisions().iter().map(|(_, nested)| nested));
                }
                assert!(
                    secondary_nodes
                        <= depth
                            .checked_mul(depth)
                            .unwrap()
                            .checked_mul(siblings)
                            .unwrap(),
                    "stored legacy caches must have a finite ancestry bound, not a duplicate canonical subtree"
                );
                let cloned = (*current[0]).clone();
                assert!(Arc::ptr_eq(
                    current[0].content_paragraph.as_ref().unwrap(),
                    cloned.content_paragraph.as_ref().unwrap()
                ));
                model.to_xml().unwrap();
                assert_eq!(
                    crate::text::ORDERED_DIAGNOSTIC_SIMPLE_PARSE_CALLS.get(),
                    fields,
                    "unchanged current serialization does not reconstruct source owners"
                );
                assert_eq!(ORDERED_DIAGNOSTIC_REVISION_PARSE_CALLS.get(), revisions);
            }
        }
        let prior =
            r#"<w:rPrChange w:id="82" w:author="Prior"><w:rPr><w:i/></w:rPr></w:rPrChange>"#;
        let inner = format!(
            r#"<w:ins w:id="81" w:author="Prior"><w:r><w:rPr>{prior}</w:rPr><w:t>copy</w:t></w:r></w:ins>"#
        );
        let source = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:ins w:id="80" w:author="Prior">{inner}</w:ins></w:p></w:body></w:document>"#
        );
        let copies_before = ORDERED_DIAGNOSTIC_SECONDARY_COPY_NODES.get();
        let visible_before = ORDERED_DIAGNOSTIC_SECONDARY_VISIBLE_RAW_BYTES.get();
        let charged_before = ORDERED_DIAGNOSTIC_SECONDARY_CHARGED_SOURCE_BYTES.get();
        let model = CT_Document::from_xml(source.as_bytes()).unwrap();
        assert_eq!(
            ORDERED_DIAGNOSTIC_SECONDARY_COPY_NODES.get() - copies_before,
            3,
            "one physical run fact plus the inner revision and its legacy run cache"
        );
        assert_eq!(
            ORDERED_DIAGNOSTIC_SECONDARY_VISIBLE_RAW_BYTES.get() - visible_before,
            inner.len() + prior.len(),
            "visible copied buffers include run-property revision raw, beyond the wrapper raw"
        );
        assert!(
            ORDERED_DIAGNOSTIC_SECONDARY_CHARGED_SOURCE_BYTES.get() - charged_before
                >= ORDERED_DIAGNOSTIC_SECONDARY_VISIBLE_RAW_BYTES.get() - visible_before
        );
        assert!(
            String::from_utf8(model.to_xml().unwrap())
                .unwrap()
                .contains(prior)
        );
        let nested = format!(
            "{}<w:r/><w:r/>{}",
            (0..33)
                .map(|id| format!(r#"<w:ins w:id="{id}" w:author="Prior">"#))
                .collect::<String>(),
            "</w:ins>".repeat(33)
        );
        assert!(
            CT_Revision::try_from_raw(nested.into_bytes(), &["w".into()], true).is_err(),
            "the existing qualified revision nesting admission remains binding"
        );
    }
}
