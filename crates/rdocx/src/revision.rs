//! Native tracked-revision inspection and atomic resolution.

use std::collections::{HashMap, HashSet};

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
pub use rdocx_oxml::RevisionKind;
use rdocx_oxml::{CT_Document, CT_Revision};

use crate::{Document, Error, ParagraphRef, Result};

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type NamespaceScope = HashMap<String, String>;
type NamespaceDeclarations = Vec<(String, String)>;

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
        self.inner
            .content_paragraph()
            .map(|inner| ParagraphRef { inner })
    }
}

#[derive(Clone, Copy)]
enum Resolution {
    Accept,
    Reject,
}

#[derive(Clone)]
enum RevisionScope<'a> {
    All,
    Author(&'a str),
    Id(i32),
    DateRange { start: Instant, end: Instant },
}

#[derive(Clone, PartialEq, Eq)]
struct RevisionMetadata {
    kind: RevisionKind,
    id: i32,
    author: String,
    timestamp: Option<String>,
}

struct XmlElement {
    start: usize,
    open_end: usize,
    close_start: usize,
    end: usize,
    name: String,
    local: String,
    word: bool,
    modeled: bool,
    empty: bool,
    parent: Option<usize>,
    children: Vec<usize>,
    revision: Option<RevisionMetadata>,
    namespace_declarations: NamespaceDeclarations,
}

struct XmlTree<'a> {
    source: &'a [u8],
    elements: Vec<XmlElement>,
    root: usize,
}

struct RenderState<'a> {
    resolution: Resolution,
    scope: RevisionScope<'a>,
    resolved: HashSet<usize>,
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

    fn resolve_revisions(
        &mut self,
        resolution: Resolution,
        scope: RevisionScope<'_>,
    ) -> Result<usize> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let source = candidate.document.to_xml()?;
        let mut tree = XmlTree::parse(&source)?;
        if let Some(packaged_xml) = candidate.package.get_part(&candidate.doc_part_name) {
            let packaged_tree = XmlTree::parse(packaged_xml)?;
            tree.recover_property_owner_namespaces(&packaged_tree);
        }
        let mut state = RenderState {
            resolution,
            scope: scope.clone(),
            resolved: HashSet::new(),
        };
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

        let story_parts = crate::comparison::related_story_part_names(&candidate)?;
        for part_name in story_parts {
            let part = candidate
                .package
                .get_part(&part_name)
                .ok_or_else(|| Error::Other(format!("missing revision story part {part_name}")))?
                .to_vec();
            let (updated, count) = resolve_story_xml(&part, resolution, scope.clone())?;
            if count > 0 {
                candidate.package.set_part(&part_name, updated);
                resolved = resolved
                    .checked_add(count)
                    .ok_or_else(|| Error::Other("resolved revision count overflowed".to_owned()))?;
            }
        }
        if resolved == 0 {
            return Ok(0);
        }
        candidate = crate::comparison::reopen_staged(candidate)?;
        self.commit_staged_mutation(candidate);
        Ok(resolved)
    }
}

fn resolve_story_xml(
    source: &[u8],
    resolution: Resolution,
    scope: RevisionScope<'_>,
) -> Result<(Vec<u8>, usize)> {
    let tree = XmlTree::parse(source)?;
    let mut state = RenderState {
        resolution,
        scope,
        resolved: HashSet::new(),
    };
    let mut output = Vec::with_capacity(source.len());
    output.extend_from_slice(&source[..tree.elements[tree.root].start]);
    output.extend_from_slice(&tree.render(tree.root, &mut state, false)?);
    output.extend_from_slice(&source[tree.elements[tree.root].end..]);
    Ok((output, state.resolved.len()))
}

pub(crate) fn modeled_revision_count(source: &[u8]) -> Result<usize> {
    let tree = XmlTree::parse(source)?;
    Ok(tree
        .elements
        .iter()
        .filter(|element| element.revision.is_some())
        .count())
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
    fn matches(&self, metadata: &RevisionMetadata) -> bool {
        match self {
            Self::All => true,
            Self::Author(author) => metadata.author == *author,
            Self::Id(id) => metadata.id == *id,
            Self::DateRange { start, end } => metadata
                .timestamp
                .as_deref()
                .and_then(|value| Instant::parse(value).ok())
                .is_some_and(|instant| *start <= instant && instant <= *end),
        }
    }
}

impl<'a> XmlTree<'a> {
    fn parse(source: &'a [u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        let mut elements: Vec<XmlElement> = Vec::new();
        let mut stack = Vec::new();
        let mut scopes = vec![NamespaceScope::new()];
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
        })
    }

    fn render(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
    ) -> Result<Vec<u8>> {
        self.render_with_namespaces(index, state, convert_deleted_text, &[])
    }

    fn recover_property_owner_namespaces(&mut self, packaged: &XmlTree<'_>) {
        let packaged_revisions = packaged
            .elements
            .iter()
            .filter_map(|element| {
                let metadata = element.revision.as_ref()?;
                let parent = element.parent?;
                is_property_change(metadata.kind).then(|| {
                    (
                        metadata.clone(),
                        packaged.elements[parent].namespace_declarations.clone(),
                    )
                })
            })
            .collect::<Vec<_>>();
        let mut used = vec![false; packaged_revisions.len()];

        for index in 0..self.elements.len() {
            let Some(metadata) = self.elements[index].revision.as_ref() else {
                continue;
            };
            if !is_property_change(metadata.kind) {
                continue;
            }
            let Some(parent) = self.elements[index].parent else {
                continue;
            };
            let Some((matched, (_, declarations))) = packaged_revisions.iter().enumerate().find(
                |(candidate, (packaged_metadata, _))| {
                    !used[*candidate] && packaged_metadata == metadata
                },
            ) else {
                continue;
            };
            used[matched] = true;
            self.elements[parent].namespace_declarations =
                merged_namespaces(&self.elements[parent].namespace_declarations, declarations);
        }
    }

    fn render_with_namespaces(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[index];
        if let Some(metadata) = &element.revision
            && state.scope.matches(metadata)
        {
            state.resolved.insert(index);
            return self.render_selected_revision(index, metadata.kind, state, promoted_namespaces);
        }

        if matches!(state.resolution, Resolution::Reject) {
            let changes = self.selected_property_changes(index, state)?;
            if let Some((change, prior)) = changes.first().copied() {
                for (selected, _) in &changes {
                    state.resolved.insert(*selected);
                    self.validate_selected_descendants(*selected, state)?;
                }
                let namespaces = merged_namespaces(
                    &merged_namespaces(promoted_namespaces, &element.namespace_declarations),
                    &self.elements[change].namespace_declarations,
                );
                return self.render_rejected_property_owner(
                    index,
                    prior,
                    &changes,
                    state,
                    &namespaces,
                );
            }
        }

        let markers = self.selected_owner_markers(index, state);
        if markers.iter().any(|marker| {
            let kind = self.elements[*marker]
                .revision
                .as_ref()
                .expect("selected marker is a revision")
                .kind;
            removes_content(state.resolution, kind)
        }) {
            if element.local == "numPr" {
                return self.render_rejected_numbering_owner(
                    index,
                    &markers,
                    state,
                    promoted_namespaces,
                );
            }
            self.validate_selected_descendants(index, state)?;
            return Ok(Vec::new());
        }

        self.render_ordinary(index, state, convert_deleted_text, promoted_namespaces)
    }

    fn render_selected_revision(
        &self,
        index: usize,
        kind: RevisionKind,
        state: &mut RenderState<'_>,
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        if is_property_change(kind) {
            let prior = self.prior_property(index, kind)?;
            if matches!(state.resolution, Resolution::Accept) {
                return Ok(Vec::new());
            }
            let namespaces = merged_namespaces(
                promoted_namespaces,
                &self.elements[index].namespace_declarations,
            );
            return self.render_with_namespaces(prior, state, false, &namespaces);
        }

        if self.is_contextual_marker(index) {
            return Ok(Vec::new());
        }

        if keeps_content(state.resolution, kind) {
            let namespaces = merged_namespaces(
                promoted_namespaces,
                &self.elements[index].namespace_declarations,
            );
            self.render_inner_with_namespaces(
                index,
                state,
                matches!(state.resolution, Resolution::Reject)
                    && matches!(kind, RevisionKind::Deletion | RevisionKind::MoveFrom),
                &namespaces,
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
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[index];
        if element.empty {
            let raw = &self.source[element.start..element.end];
            let raw = inject_namespace_declarations(
                raw,
                &element.namespace_declarations,
                promoted_namespaces,
            );
            return Ok(
                if convert_deleted_text && element.word && element.local == "delText" {
                    rename_element(&raw, &element.name, "t")
                } else {
                    raw
                },
            );
        }

        let resolves_paragraph_property_change = element.word
            && element.local == "pPr"
            && !self.selected_property_changes(index, state)?.is_empty();
        let removes_empty_property_shell = element.word
            && matches!(element.local.as_str(), "pPr" | "rPr" | "trPr" | "numPr")
            && (resolves_paragraph_property_change
                || self.subtree_contains_selected_contextual_marker(index, state));
        let inner = self.render_inner_with_namespaces(
            index,
            state,
            convert_deleted_text,
            promoted_namespaces,
        )?;
        if element.word && element.local == "tbl" && self.owned_rows_all_remove(index, state) {
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
            promoted_namespaces,
        );
        if convert_deleted_text && element.word && element.local == "delText" {
            output.extend_from_slice(&rename_element(&open, &element.name, "t"));
        } else {
            output.extend_from_slice(&open);
        }
        output.extend_from_slice(&inner);
        let close = &self.source[element.close_start..element.end];
        if convert_deleted_text && element.word && element.local == "delText" {
            output.extend_from_slice(&rename_element(close, &element.name, "t"));
        } else {
            output.extend_from_slice(close);
        }
        Ok(output)
    }

    fn render_inner_with_namespaces(
        &self,
        index: usize,
        state: &mut RenderState<'_>,
        convert_deleted_text: bool,
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[index];
        let mut output = Vec::new();
        let mut cursor = element.open_end;
        let mut child_index = 0;
        while child_index < element.children.len() {
            let child = element.children[child_index];
            let child_element = &self.elements[child];
            output.extend_from_slice(&self.source[cursor..child_element.start]);
            if child_element.word
                && child_element.local == "p"
                && self.paragraph_mark_removes(child, state)
            {
                let mut paragraphs = vec![child];
                let mut next_index = child_index + 1;
                loop {
                    let next = element.children.get(next_index).copied().ok_or_else(|| {
                        Error::Other("cannot remove the final paragraph mark".to_owned())
                    })?;
                    if !self.elements[next].word || self.elements[next].local != "p" {
                        return Err(Error::Other(
                            "paragraph mark removal requires an adjacent paragraph".to_owned(),
                        ));
                    }
                    paragraphs.push(next);
                    next_index += 1;
                    if !self.paragraph_mark_removes(next, state) {
                        break;
                    }
                }
                for pair in paragraphs.windows(2) {
                    output.extend_from_slice(
                        &self.source[self.elements[pair[0]].end..self.elements[pair[1]].start],
                    );
                }
                output.extend_from_slice(&self.render_merged_paragraphs(&paragraphs, state)?);
                cursor = self.elements[*paragraphs.last().expect("paragraph chain exists")].end;
                child_index = next_index;
                continue;
            }
            output.extend_from_slice(&self.render_with_namespaces(
                child,
                state,
                convert_deleted_text,
                promoted_namespaces,
            )?);
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
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let prior_xml = self.render_with_namespaces(prior, state, false, promoted_namespaces)?;
        let selected = selected
            .iter()
            .map(|(change, _)| *change)
            .collect::<HashSet<_>>();
        let retained =
            self.render_retained_owner_children(owner, &selected, state, promoted_namespaces)?;
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
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let selected = selected.iter().copied().collect::<HashSet<_>>();
        for marker in &selected {
            state.resolved.insert(*marker);
            self.validate_selected_descendants(*marker, state)?;
        }
        let retained =
            self.render_retained_owner_children(owner, &selected, state, promoted_namespaces)?;
        if retained.is_empty() {
            return Ok(Vec::new());
        }
        let element = &self.elements[owner];
        let mut output = self.source[element.start..element.open_end].to_vec();
        output.extend_from_slice(&retained);
        output.extend_from_slice(&self.source[element.close_start..element.end]);
        Ok(output)
    }

    fn render_retained_owner_children(
        &self,
        owner: usize,
        selected: &HashSet<usize>,
        state: &mut RenderState<'_>,
        promoted_namespaces: &[(String, String)],
    ) -> Result<Vec<u8>> {
        let element = &self.elements[owner];
        let child_namespaces =
            merged_namespaces(promoted_namespaces, &element.namespace_declarations);
        let mut output = Vec::new();
        for child in &element.children {
            let child_element = &self.elements[*child];
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
                    &child_namespaces,
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
        for child in &self.elements[index].children {
            self.render(*child, state, false)?;
        }
        Ok(())
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

    fn render_merged_paragraphs(
        &self,
        paragraphs: &[usize],
        state: &mut RenderState<'_>,
    ) -> Result<Vec<u8>> {
        let current = paragraphs[0];
        let next = *paragraphs.last().expect("paragraph chain exists");
        let current_element = &self.elements[current];
        let next_element = &self.elements[next];
        let mut output = self.source[current_element.start..current_element.open_end].to_vec();
        if let Some(properties) = next_element
            .children
            .iter()
            .find(|child| self.elements[**child].word && self.elements[**child].local == "pPr")
        {
            output.extend_from_slice(&self.render(*properties, state, false)?);
        }
        for paragraph in paragraphs {
            for child in &self.elements[*paragraph].children {
                let child_element = &self.elements[*child];
                if child_element.word && child_element.local == "pPr" {
                    if *paragraph != next {
                        self.validate_selected_descendants(*child, state)?;
                    }
                } else {
                    output.extend_from_slice(&self.render(*child, state, false)?);
                }
            }
        }
        output.extend_from_slice(&self.source[current_element.close_start..current_element.end]);
        Ok(output)
    }
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
    let revision = modeled
        .then(|| revision_metadata(start, &local, scope))
        .transpose()?
        .flatten();
    if modeled && is_revision_local(&local) && revision.is_none() {
        modeled = false;
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
struct Instant {
    seconds: i64,
    fraction: String,
}

impl Instant {
    fn parse(value: &str) -> Result<Self> {
        parse_rfc3339(value)
            .ok_or_else(|| Error::Other(format!("invalid RFC 3339 revision timestamp: {value}")))
    }
}

fn parse_rfc3339(value: &str) -> Option<Instant> {
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
