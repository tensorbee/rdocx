use std::collections::{BTreeMap, HashMap};
use std::hash::{DefaultHasher, Hash, Hasher};
use std::path::PathBuf;

use oxml_py_support::{PathSeg, RevisionCounter, StaleElementError};
use pyo3::exceptions::{
    PyFileNotFoundError, PyIndexError, PyNotADirectoryError, PyOverflowError, PyTypeError,
    PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBytes, PyList, PyTuple};
use smallvec::smallvec;

use rdocx_oxml::{ST_PageOrientation, ST_SectionType};

use crate::paragraph::{
    ParagraphLocation, PyParagraph, PyParagraphCollection, defined_style, style_id_of_type,
};
use crate::rdocx_to_pyerr;
use crate::story::{PyHeaderFooter, PySettings, SectionLayout, StorySlot};
use crate::table::{PyTable, PyTableCollection};

#[pyclass(name = "RunPosition", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct PyRunPosition {
    pub body_index: usize,
    pub run_index: usize,
}

#[pymethods]
impl PyRunPosition {
    #[new]
    #[pyo3(signature = (*, body_index, run_index))]
    fn new(body_index: usize, run_index: usize) -> Self {
        Self {
            body_index,
            run_index,
        }
    }
}

impl From<PyRunPosition> for rdocx::RunPosition {
    fn from(value: PyRunPosition) -> Self {
        Self {
            body_index: value.body_index,
            run_index: value.run_index,
        }
    }
}

#[pyclass(name = "RunRange", frozen, eq, skip_from_py_object)]
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct PyRunRange {
    start: PyRunPosition,
    end: PyRunPosition,
}

#[pymethods]
impl PyRunRange {
    #[new]
    #[pyo3(signature = (*, start, end))]
    fn new(start: PyRef<'_, PyRunPosition>, end: PyRef<'_, PyRunPosition>) -> Self {
        Self {
            start: *start,
            end: *end,
        }
    }

    #[getter]
    fn start(&self) -> PyRunPosition {
        self.start
    }

    #[getter]
    fn end(&self) -> PyRunPosition {
        self.end
    }
}

impl From<PyRunRange> for rdocx::RunRange {
    fn from(value: PyRunRange) -> Self {
        Self {
            start: value.start.into(),
            end: value.end.into(),
        }
    }
}

impl From<rdocx::RunRange> for PyRunRange {
    fn from(value: rdocx::RunRange) -> Self {
        let position = |position: rdocx::RunPosition| PyRunPosition {
            body_index: position.body_index,
            run_index: position.run_index,
        };
        Self {
            start: position(value.start),
            end: position(value.end),
        }
    }
}

#[pyclass(name = "Bookmark", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyBookmark {
    id: Option<i32>,
    name: Option<String>,
    range: Option<PyRunRange>,
    direct_range: Option<PyRunRange>,
    text: String,
    issue: Option<String>,
}

#[pymethods]
impl PyBookmark {
    #[new]
    #[pyo3(signature = (*, id, name, range, direct_range, text, issue))]
    fn new(
        id: Option<i32>,
        name: Option<String>,
        range: Option<PyRef<'_, PyRunRange>>,
        direct_range: Option<PyRef<'_, PyRunRange>>,
        text: String,
        issue: Option<String>,
    ) -> Self {
        Self {
            id,
            name,
            range: range.map(|range| *range),
            direct_range: direct_range.map(|range| *range),
            text,
            issue,
        }
    }

    #[getter]
    fn id(&self) -> Option<i32> {
        self.id
    }

    #[getter]
    fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    #[getter]
    fn range(&self) -> Option<PyRunRange> {
        self.range
    }

    #[getter]
    fn direct_range(&self) -> Option<PyRunRange> {
        self.direct_range
    }

    #[getter]
    fn text(&self) -> &str {
        &self.text
    }

    #[getter]
    fn issue(&self) -> Option<&str> {
        self.issue.as_deref()
    }
}

#[pyclass(name = "Comment", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone)]
pub struct PyComment {
    pub id: i32,
    pub author: Option<String>,
    pub initials: Option<String>,
    pub date: Option<String>,
    pub text: String,
    pub parent_id: Option<i32>,
    pub resolved: bool,
    pub anchor_text: Option<String>,
    pub anchor: Option<PyStoryRunRange>,
}

// Preserve the established metadata value equality. Derived anchor snapshots
// carry document revisions and do not change the identity of a comment record.
impl PartialEq for PyComment {
    fn eq(&self, other: &Self) -> bool {
        self.id == other.id
            && self.author == other.author
            && self.initials == other.initials
            && self.date == other.date
            && self.text == other.text
            && self.parent_id == other.parent_id
            && self.resolved == other.resolved
    }
}

impl Eq for PyComment {}

#[pymethods]
impl PyComment {
    #[new]
    #[pyo3(signature = (*, id, author, initials, date, text, parent_id, resolved, anchor_text=None, anchor=None))]
    #[allow(clippy::too_many_arguments)] // Preserve seven original fields plus optional anchor snapshots.
    fn new(
        id: i32,
        author: Option<String>,
        initials: Option<String>,
        date: Option<String>,
        text: String,
        parent_id: Option<i32>,
        resolved: bool,
        anchor_text: Option<String>,
        anchor: Option<PyRef<'_, PyStoryRunRange>>,
    ) -> Self {
        Self {
            anchor_text,
            anchor: anchor.map(|range| (*range).clone()),
            id,
            author,
            initials,
            date,
            text,
            parent_id,
            resolved,
        }
    }
}

#[pyclass(
    name = "ComparisonDiagnostic",
    frozen,
    get_all,
    eq,
    skip_from_py_object
)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyComparisonDiagnostic {
    pub location: String,
    pub message: String,
}

#[pymethods]
impl PyComparisonDiagnostic {
    #[new]
    #[pyo3(signature = (*, location, message))]
    fn new(location: String, message: String) -> Self {
        Self { location, message }
    }
}

#[pyclass(name = "SvgDiagnostic", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PySvgDiagnostic {
    pub path: String,
    pub message: String,
}

#[pymethods]
impl PySvgDiagnostic {
    #[new]
    #[pyo3(signature = (*, path, message))]
    fn new(path: String, message: String) -> Self {
        Self { path, message }
    }
}

#[pyclass(name = "SvgRenderResult", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PySvgRenderResult {
    svg: String,
    diagnostics: Vec<PySvgDiagnostic>,
}

#[pymethods]
impl PySvgRenderResult {
    #[new]
    #[pyo3(signature = (*, svg, diagnostics))]
    fn new(svg: String, diagnostics: Vec<PyRef<'_, PySvgDiagnostic>>) -> Self {
        Self {
            svg,
            diagnostics: diagnostics
                .iter()
                .map(|diagnostic| (**diagnostic).clone())
                .collect(),
        }
    }

    #[getter]
    fn svg(&self) -> &str {
        &self.svg
    }

    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.diagnostics.iter().cloned())
    }
}

#[pyclass(name = "BoundingBox", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, Copy, PartialEq)]
pub struct PyBoundingBox {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

#[pymethods]
impl PyBoundingBox {
    #[new]
    #[pyo3(signature = (*, x, y, width, height))]
    fn new(x: f64, y: f64, width: f64, height: f64) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }
}

#[pyclass(name = "LayoutFragment", frozen, eq, skip_from_py_object)]
#[derive(Clone, Copy, PartialEq)]
pub struct PyLayoutFragment {
    body_index: usize,
    physical_page: usize,
    displayed_page: usize,
    bounds: PyBoundingBox,
}

#[pymethods]
impl PyLayoutFragment {
    #[new]
    #[pyo3(signature = (*, body_index, physical_page, displayed_page, bounds))]
    fn new(
        body_index: usize,
        physical_page: usize,
        displayed_page: usize,
        bounds: PyRef<'_, PyBoundingBox>,
    ) -> Self {
        Self {
            body_index,
            physical_page,
            displayed_page,
            bounds: *bounds,
        }
    }

    #[getter]
    fn body_index(&self) -> usize {
        self.body_index
    }

    #[getter]
    fn physical_page(&self) -> usize {
        self.physical_page
    }

    #[getter]
    fn displayed_page(&self) -> usize {
        self.displayed_page
    }

    #[getter]
    fn bounds(&self) -> PyBoundingBox {
        self.bounds
    }
}

#[pyclass(name = "LayoutPage", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, Copy, PartialEq)]
pub struct PyLayoutPage {
    pub page_number: usize,
    pub displayed_page_number: usize,
    pub width: f64,
    pub height: f64,
}

#[pymethods]
impl PyLayoutPage {
    #[new]
    #[pyo3(signature = (*, page_number, displayed_page_number, width, height))]
    fn new(page_number: usize, displayed_page_number: usize, width: f64, height: f64) -> Self {
        Self {
            page_number,
            displayed_page_number,
            width,
            height,
        }
    }
}

#[pyclass(name = "TocRebuildReport", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyTocRebuildReport {
    entry_count: usize,
    bookmark_count: usize,
    diagnostics: Vec<String>,
}

#[pyclass(
    name = "LayoutBackedFieldUpdateReport",
    frozen,
    eq,
    skip_from_py_object
)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyLayoutBackedFieldUpdateReport {
    page_fields: usize,
    num_pages_fields: usize,
    page_reference_fields: usize,
    section_fields: usize,
    section_pages_fields: usize,
    diagnostics: Vec<String>,
}

#[pyclass(name = "Revision", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyRevision {
    pub id: i32,
    pub author: String,
    pub timestamp: Option<String>,
    pub kind: String,
    pub story: Option<PyStory>,
}

#[pymethods]
impl PyRevision {
    #[new]
    #[pyo3(signature = (*, id, author, timestamp, kind, story=None))]
    fn new(
        id: i32,
        author: String,
        timestamp: Option<String>,
        kind: String,
        story: Option<PyRef<'_, PyStory>>,
    ) -> Self {
        Self {
            id,
            author,
            timestamp,
            kind,
            story: story.map(|story| story.clone()),
        }
    }
}

fn revision_kind_name(kind: rdocx::RevisionKind) -> &'static str {
    match kind {
        rdocx::RevisionKind::Insertion => "insertion",
        rdocx::RevisionKind::Deletion => "deletion",
        rdocx::RevisionKind::MoveFrom => "move_from",
        rdocx::RevisionKind::MoveTo => "move_to",
        rdocx::RevisionKind::RunPropertyChange => "run_property_change",
        rdocx::RevisionKind::ParagraphPropertyChange => "paragraph_property_change",
        rdocx::RevisionKind::TablePropertyChange => "table_property_change",
        rdocx::RevisionKind::SectionPropertyChange => "section_property_change",
    }
}

#[pyclass(name = "Story", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyStory {
    pub kind: String,
    pub part_name: String,
    pub owner_index: usize,
}

#[pymethods]
impl PyStory {
    #[new]
    #[pyo3(signature = (*, kind, part_name, owner_index))]
    fn new(kind: String, part_name: String, owner_index: usize) -> Self {
        Self {
            kind,
            part_name,
            owner_index,
        }
    }
}

#[pyclass(name = "StoryItem", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyStoryItem {
    story: PyStory,
    kind: String,
    index_path: Vec<usize>,
    direct_body_index: Option<usize>,
    text: Option<String>,
    xml: Vec<u8>,
    revision: u64,
}

#[pymethods]
impl PyStoryItem {
    #[new]
    #[pyo3(signature = (*, story, kind, index_path, text, xml=None, direct_body_index=None, revision=0))]
    fn new(
        story: PyRef<'_, PyStory>,
        kind: String,
        index_path: Vec<usize>,
        text: Option<String>,
        xml: Option<&[u8]>,
        direct_body_index: Option<usize>,
        revision: u64,
    ) -> Self {
        Self {
            story: story.clone(),
            kind,
            index_path,
            direct_body_index,
            text,
            xml: xml.unwrap_or_default().to_vec(),
            revision,
        }
    }

    #[getter]
    fn story(&self) -> PyStory {
        self.story.clone()
    }

    #[getter]
    fn kind(&self) -> &str {
        &self.kind
    }

    #[getter]
    fn index_path<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.index_path.iter().copied())
    }

    #[getter]
    fn direct_body_index(&self) -> Option<usize> {
        self.direct_body_index
    }

    #[getter]
    fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    #[getter]
    fn xml<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.xml)
    }

    #[getter]
    fn revision(&self) -> u64 {
        self.revision
    }
}

#[pyclass(name = "StoryRunPosition", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyStoryRunPosition {
    item: PyStoryItem,
    run_index: usize,
}

#[pymethods]
impl PyStoryRunPosition {
    /// Take a `StoryItem`, or a `Paragraph` handle, which also reaches a
    /// paragraph inside a block content control.
    #[new]
    #[pyo3(signature = (*, item = None, run_index, paragraph = None))]
    fn new(
        py: Python<'_>,
        item: Option<PyRef<'_, PyStoryItem>>,
        run_index: usize,
        paragraph: Option<PyRef<'_, PyParagraph>>,
    ) -> PyResult<Self> {
        let item = match (item, paragraph) {
            (Some(item), None) => item.clone(),
            (None, Some(paragraph)) => PyDocument::paragraph_story_item(py, &paragraph)?,
            _ => {
                return Err(PyTypeError::new_err(
                    "StoryRunPosition takes exactly one of item and paragraph",
                ));
            }
        };
        Ok(Self { item, run_index })
    }

    #[getter]
    fn item(&self) -> PyStoryItem {
        self.item.clone()
    }

    #[getter]
    fn run_index(&self) -> usize {
        self.run_index
    }
}

#[pyclass(name = "StoryRunRange", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyStoryRunRange {
    start: PyStoryRunPosition,
    end: PyStoryRunPosition,
}

#[pymethods]
impl PyStoryRunRange {
    #[new]
    #[pyo3(signature = (*, start, end))]
    fn new(start: PyRef<'_, PyStoryRunPosition>, end: PyRef<'_, PyStoryRunPosition>) -> Self {
        Self {
            start: start.clone(),
            end: end.clone(),
        }
    }

    #[getter]
    fn start(&self) -> PyStoryRunPosition {
        self.start.clone()
    }

    #[getter]
    fn end(&self) -> PyStoryRunPosition {
        self.end.clone()
    }
}

#[pyclass(name = "ContentFragment", frozen, skip_from_py_object)]
pub struct PyContentFragment {
    inner: rdocx::ContentFragment,
}

#[pymethods]
impl PyContentFragment {
    #[getter]
    fn kind(&self) -> &'static str {
        story_item_kind_name(self.inner.kind())
    }
}

#[pyclass(name = "Hyperlink", frozen, eq, skip_from_py_object)]
#[derive(Clone)]
pub struct PyHyperlink {
    story: PyStory,
    index_path: Vec<usize>,
    text: String,
    url: Option<String>,
    anchor: Option<String>,
    relationship_id: Option<String>,
    /// The position among the native `story_links` of `story`, which
    /// selects one of several identical links, and the story's link layout
    /// when the snapshot was taken. A constructed record has neither.
    position: Option<(usize, u64)>,
}

/// Equality covers the public fields only, so a snapshot equals a record
/// rebuilt from them.
impl PartialEq for PyHyperlink {
    fn eq(&self, other: &Self) -> bool {
        self.story == other.story
            && self.index_path == other.index_path
            && self.text == other.text
            && self.url == other.url
            && self.anchor == other.anchor
            && self.relationship_id == other.relationship_id
    }
}

impl Eq for PyHyperlink {}

/// A digest of one story's link layout: the item path and text of every link
/// in order. Retargeting a link keeps it, and adding or removing one changes it.
fn story_link_layout<'a>(links: impl IntoIterator<Item = (&'a [usize], &'a str)>) -> u64 {
    let mut hasher = DefaultHasher::new();
    for (index_path, text) in links {
        index_path.hash(&mut hasher);
        text.hash(&mut hasher);
    }
    hasher.finish()
}

#[pymethods]
impl PyHyperlink {
    #[new]
    #[pyo3(signature = (*, story, index_path, text, url, anchor, relationship_id))]
    fn new(
        story: PyRef<'_, PyStory>,
        index_path: Vec<usize>,
        text: String,
        url: Option<String>,
        anchor: Option<String>,
        relationship_id: Option<String>,
    ) -> Self {
        Self {
            story: story.clone(),
            index_path,
            text,
            url,
            anchor,
            relationship_id,
            position: None,
        }
    }

    #[getter]
    fn story(&self) -> PyStory {
        self.story.clone()
    }

    #[getter]
    fn index_path<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.index_path.iter().copied())
    }

    #[getter]
    fn text(&self) -> &str {
        &self.text
    }

    #[getter]
    fn url(&self) -> Option<&str> {
        self.url.as_deref()
    }

    #[getter]
    fn anchor(&self) -> Option<&str> {
        self.anchor.as_deref()
    }

    #[getter]
    fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
}

#[pyclass(name = "HeaderFooterVariant", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyHeaderFooterVariant {
    section_index: usize,
    kind: String,
    variant: String,
    story: Option<PyStory>,
    source_section: Option<usize>,
    inherited: bool,
}

#[pymethods]
impl PyHeaderFooterVariant {
    #[new]
    #[pyo3(signature = (*, section_index, kind, variant, story, source_section, inherited))]
    fn new(
        section_index: usize,
        kind: String,
        variant: String,
        story: Option<PyRef<'_, PyStory>>,
        source_section: Option<usize>,
        inherited: bool,
    ) -> Self {
        Self {
            section_index,
            kind,
            variant,
            story: story.map(|value| value.clone()),
            source_section,
            inherited,
        }
    }

    #[getter]
    fn section_index(&self) -> usize {
        self.section_index
    }

    #[getter]
    fn kind(&self) -> &str {
        &self.kind
    }

    #[getter]
    fn variant(&self) -> &str {
        &self.variant
    }

    #[getter]
    fn story(&self) -> Option<PyStory> {
        self.story.clone()
    }

    #[getter]
    fn source_section(&self) -> Option<usize> {
        self.source_section
    }

    #[getter]
    fn inherited(&self) -> bool {
        self.inherited
    }
}

#[pyclass(name = "Section", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PySection {
    #[pyo3(get)]
    pub ordinal: usize,
    #[pyo3(get)]
    pub is_final: bool,
    #[pyo3(get)]
    pub orientation: Option<String>,
    #[pyo3(get)]
    pub page_width: Option<i64>,
    #[pyo3(get)]
    pub page_height: Option<i64>,
    #[pyo3(get)]
    pub margin_top: Option<i64>,
    #[pyo3(get)]
    pub margin_right: Option<i64>,
    #[pyo3(get)]
    pub margin_bottom: Option<i64>,
    #[pyo3(get)]
    pub margin_left: Option<i64>,
    #[pyo3(get)]
    pub gutter: Option<i64>,
    #[pyo3(get)]
    pub column_count: Option<u32>,
    #[pyo3(get)]
    pub column_spacing: Option<i64>,
    #[pyo3(get)]
    pub page_number_start: Option<u32>,
    #[pyo3(get)]
    pub header_distance: Option<i64>,
    #[pyo3(get)]
    pub footer_distance: Option<i64>,
    #[pyo3(get)]
    pub different_first_page: Option<bool>,
    #[pyo3(get)]
    pub break_type: Option<String>,
    /// The document a section read from `Document.sections` belongs to,
    /// which its header and footer accessors edit.
    pub(crate) owner: SectionOwner,
}

/// The document of a section snapshot and its section layout when read,
/// ignored when snapshots are compared.
#[derive(Default)]
pub(crate) struct SectionOwner(Option<(Py<PyDocument>, SectionLayout)>);

impl SectionOwner {
    pub(crate) fn of(document: Py<PyDocument>, layout: SectionLayout) -> Self {
        Self(Some((document, layout)))
    }
}

impl Clone for SectionOwner {
    fn clone(&self) -> Self {
        Self(
            self.0
                .as_ref()
                .map(|(document, layout)| (Python::attach(|py| document.clone_ref(py)), *layout)),
        )
    }
}

impl PartialEq for SectionOwner {
    fn eq(&self, _: &Self) -> bool {
        true
    }
}

impl Eq for SectionOwner {}

#[pymethods]
impl PySection {
    #[new]
    #[pyo3(signature = (*, ordinal, is_final, orientation, page_width, page_height, margin_top, margin_right, margin_bottom, margin_left, gutter, column_count, column_spacing, page_number_start, header_distance, footer_distance, different_first_page, break_type))]
    #[allow(clippy::too_many_arguments)]
    fn new(
        ordinal: usize,
        is_final: bool,
        orientation: Option<String>,
        page_width: Option<i64>,
        page_height: Option<i64>,
        margin_top: Option<i64>,
        margin_right: Option<i64>,
        margin_bottom: Option<i64>,
        margin_left: Option<i64>,
        gutter: Option<i64>,
        column_count: Option<u32>,
        column_spacing: Option<i64>,
        page_number_start: Option<u32>,
        header_distance: Option<i64>,
        footer_distance: Option<i64>,
        different_first_page: Option<bool>,
        break_type: Option<String>,
    ) -> Self {
        Self {
            ordinal,
            is_final,
            orientation,
            page_width,
            page_height,
            margin_top,
            margin_right,
            margin_bottom,
            margin_left,
            gutter,
            column_count,
            column_spacing,
            page_number_start,
            header_distance,
            footer_distance,
            different_first_page,
            break_type,
            owner: SectionOwner::default(),
        }
    }

    // ---- Headers, footers and python-docx section names ----

    #[getter]
    fn header(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.header_footer(
            py,
            rdocx::HeaderFooterKind::Header,
            rdocx::HdrFtrType::Default,
        )
    }

    #[getter]
    fn footer(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.header_footer(
            py,
            rdocx::HeaderFooterKind::Footer,
            rdocx::HdrFtrType::Default,
        )
    }

    #[getter]
    fn first_page_header(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.header_footer(
            py,
            rdocx::HeaderFooterKind::Header,
            rdocx::HdrFtrType::First,
        )
    }

    #[getter]
    fn first_page_footer(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.header_footer(
            py,
            rdocx::HeaderFooterKind::Footer,
            rdocx::HdrFtrType::First,
        )
    }

    #[getter]
    fn even_page_header(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.header_footer(py, rdocx::HeaderFooterKind::Header, rdocx::HdrFtrType::Even)
    }

    #[getter]
    fn even_page_footer(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.header_footer(py, rdocx::HeaderFooterKind::Footer, rdocx::HdrFtrType::Even)
    }

    /// Whether the section's first page shows its first-page header and
    /// footer, read from and written to the document.
    #[getter]
    fn different_first_page_header_footer(&self, py: Python<'_>) -> PyResult<bool> {
        if self.owner.0.is_none() {
            return Ok(self.different_first_page.unwrap_or(false));
        }
        let document = self.owned_document(py)?.borrow(py);
        let section = document
            .inner
            .section(self.ordinal)
            .ok_or_else(|| PyIndexError::new_err("section index out of range"))?;
        Ok(section.different_first_page().unwrap_or(false))
    }

    // A section is a snapshot: only the python-docx header and footer switch
    // writes through, and every other name says how to change it.
    fn __setattr__(&self, py: Python<'_>, name: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        if name == "different_first_page_header_footer" {
            let enabled = value.extract::<bool>()?;
            let document = self.owned_document(py)?;
            let mut document = document.borrow_mut(py);
            document
                .inner
                .section_mut(self.ordinal)
                .ok_or_else(|| PyIndexError::new_err("section index out of range"))?
                .set_different_first_page(enabled);
            return Ok(());
        }
        let field = match name {
            "left_margin" => "margin_left",
            "right_margin" => "margin_right",
            "top_margin" => "margin_top",
            "bottom_margin" => "margin_bottom",
            "start_type" => "break_type",
            other => other,
        };
        Err(pyo3::exceptions::PyAttributeError::new_err(format!(
            "Section is a snapshot and {name} cannot be set on it, use \
             document.update_section({}, {field}=...) with a Length such as Inches(1)",
            self.ordinal
        )))
    }
}

impl PySection {
    /// The document this snapshot was read from, while its sections are
    /// still the ones the snapshot saw.
    fn owned_document(&self, py: Python<'_>) -> PyResult<&Py<PyDocument>> {
        let (document, layout) = self.owner.0.as_ref().ok_or_else(|| {
            PyValueError::new_err(
                "this Section was built by hand, read one from document.sections to reach its \
                 headers and footers",
            )
        })?;
        layout.check(py, &document.borrow(py), "section", self.ordinal, "")?;
        Ok(document)
    }

    fn header_footer(
        &self,
        py: Python<'_>,
        kind: rdocx::HeaderFooterKind,
        variant: rdocx::HdrFtrType,
    ) -> PyResult<Py<PyHeaderFooter>> {
        let document = self.owned_document(py)?.clone_ref(py);
        let layout = self
            .owner
            .0
            .as_ref()
            .map(|(_, layout)| *layout)
            .expect("owned");
        Py::new(
            py,
            PyHeaderFooter::new(
                document,
                StorySlot::new(self.ordinal, kind, variant),
                layout,
            ),
        )
    }
}

#[pyclass(name = "Style", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyStyle {
    pub style_id: String,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub style_type: String,
    pub linked_style: Option<String>,
    pub next_style: Option<String>,
    pub priority: Option<u32>,
    pub auto_redefine: Option<bool>,
    pub hidden: Option<bool>,
    pub semi_hidden: Option<bool>,
    pub unhide_when_used: Option<bool>,
    pub quick_format: Option<bool>,
    pub locked: Option<bool>,
    pub is_default: bool,
}

#[pymethods]
impl PyStyle {
    #[new]
    #[pyo3(signature = (*, style_id, name, based_on, style_type, linked_style, next_style, priority, auto_redefine, hidden, semi_hidden, unhide_when_used, quick_format, locked, is_default))]
    #[allow(clippy::too_many_arguments)]
    fn new(
        style_id: String,
        name: Option<String>,
        based_on: Option<String>,
        style_type: String,
        linked_style: Option<String>,
        next_style: Option<String>,
        priority: Option<u32>,
        auto_redefine: Option<bool>,
        hidden: Option<bool>,
        semi_hidden: Option<bool>,
        unhide_when_used: Option<bool>,
        quick_format: Option<bool>,
        locked: Option<bool>,
        is_default: bool,
    ) -> Self {
        Self {
            style_id,
            name,
            based_on,
            style_type,
            linked_style,
            next_style,
            priority,
            auto_redefine,
            hidden,
            semi_hidden,
            unhide_when_used,
            quick_format,
            locked,
            is_default,
        }
    }
}

/// One level of a numbering definition for `Document.add_numbering_definition`.
///
/// `format` is a Word `w:numFmt` name such as `decimal`, `lowerLetter` or
/// `bullet`. `text` is the level text, `%1.` or a bullet glyph by default.
/// The indents are EMU, as `Length` values are, and default to half an inch
/// per level with a quarter-inch hanging indent.
#[pyclass(name = "ListLevel", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyListLevel {
    pub format: String,
    pub text: Option<String>,
    pub start: Option<u32>,
    pub left_indent: Option<i64>,
    pub hanging_indent: Option<i64>,
}

#[pymethods]
impl PyListLevel {
    #[new]
    #[pyo3(signature = (*, format = "decimal", text = None, start = None, left_indent = None, hanging_indent = None))]
    fn new(
        format: &str,
        text: Option<String>,
        start: Option<u32>,
        left_indent: Option<i64>,
        hanging_indent: Option<i64>,
    ) -> PyResult<Self> {
        if let rdocx::ListNumberFormat::Other(_) = rdocx::ListNumberFormat::from_name(format) {
            return Err(PyValueError::new_err(format!(
                "'{format}' is not a Word numbering format"
            )));
        }
        if hanging_indent.is_some_and(|value| value < 0) {
            return Err(PyValueError::new_err("hanging_indent cannot be negative"));
        }
        Ok(Self {
            format: format.to_owned(),
            text,
            start,
            left_indent,
            hanging_indent,
        })
    }
}

impl PyListLevel {
    fn native(&self) -> rdocx::ListLevel {
        let twips = |value: Option<i64>| value.map(|emu| rdocx::Length::emu(emu).as_twips());
        let mut level = rdocx::ListLevel::new(rdocx::ListNumberFormat::from_name(&self.format))
            .indentation(twips(self.left_indent), twips(self.hanging_indent), None);
        level.start = self.start;
        match &self.text {
            Some(text) => level.level_text(text.as_str()),
            None => level,
        }
    }
}

/// The formatting `Document.add_style` gives a new style. Lengths and the
/// font size are EMU, as `Length` values are, and `None` leaves a property to
/// the style's base.
struct StyleFormatting {
    font_name: Option<String>,
    font_size: Option<i64>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<(u8, u8, u8)>,
    space_before: Option<i64>,
    space_after: Option<i64>,
    left_indent: Option<i64>,
    right_indent: Option<i64>,
    first_line_indent: Option<i64>,
}

impl StyleFormatting {
    /// The run properties, written as the `Font` setters write them on a run.
    fn run_properties(&self) -> Option<rdocx::CT_RPr> {
        let font = || self.font_name.clone();
        let size = self
            .font_size
            .map(|emu| rdocx::HalfPoint::from_pt(rdocx::Length::emu(emu).to_pt()));
        let properties = rdocx::CT_RPr {
            font_ascii: font(),
            font_hansi: font(),
            font_east_asia: font(),
            font_cs: font(),
            bold: self.bold,
            bold_cs: self.bold,
            italic: self.italic,
            italic_cs: self.italic,
            sz: size,
            sz_cs: size,
            color: self
                .color
                .map(|(red, green, blue)| format!("{red:02X}{green:02X}{blue:02X}")),
            ..rdocx::CT_RPr::default()
        };
        (properties != rdocx::CT_RPr::default()).then_some(properties)
    }

    /// The paragraph properties, written as the `ParagraphFormat` setters
    /// write them, a negative first-line indent becoming a hanging one.
    fn paragraph_properties(&self) -> Option<rdocx::CT_PPr> {
        let twips = |value: Option<i64>| value.map(|emu| rdocx::Length::emu(emu).as_twips());
        let first_line = twips(self.first_line_indent);
        let properties = rdocx::CT_PPr {
            space_before: twips(self.space_before),
            space_after: twips(self.space_after),
            ind_left: twips(self.left_indent),
            ind_right: twips(self.right_indent),
            ind_first_line: first_line.filter(|value| value.0 >= 0),
            ind_hanging: first_line
                .filter(|value| value.0 < 0)
                .map(|value| rdocx::Twips(value.0.saturating_abs())),
            ..rdocx::CT_PPr::default()
        };
        (properties != rdocx::CT_PPr::default()).then_some(properties)
    }
}

/// The style ID Word derives from a style name: the name's ASCII letters,
/// digits and hyphens, so "Q&A" gives `QA`. A name with none of them, such as
/// a Japanese one, gets the first of `a`, `a0`, `a1` and so on that `document`
/// does not use, as Word numbers them. As in python-docx, Word's lowercase
/// built-in names `caption` and `heading 1` to `heading 9` keep their
/// capitalised IDs.
fn style_id_from_name(document: &rdocx::Document, name: &str) -> String {
    match (name, name.strip_prefix("heading ")) {
        ("caption", _) => return "Caption".to_owned(),
        (_, Some(level)) if matches!(level.as_bytes(), [b'1'..=b'9']) => {
            return format!("Heading{level}");
        }
        _ => {}
    }
    let kept = name
        .chars()
        .filter(|character| character.is_ascii_alphanumeric() || *character == '-')
        .collect::<String>();
    if !kept.is_empty() {
        return kept;
    }
    let mut fallback = "a".to_owned();
    let mut number = 0;
    while document.style(&fallback).is_some() {
        fallback = format!("a{number}");
        number += 1;
    }
    fallback
}

#[pymethods]
impl PyTocRebuildReport {
    #[new]
    #[pyo3(signature = (*, entry_count, bookmark_count, diagnostics))]
    fn new(entry_count: usize, bookmark_count: usize, diagnostics: Vec<String>) -> Self {
        Self {
            entry_count,
            bookmark_count,
            diagnostics,
        }
    }

    #[getter]
    fn entry_count(&self) -> usize {
        self.entry_count
    }

    #[getter]
    fn bookmark_count(&self) -> usize {
        self.bookmark_count
    }

    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.diagnostics)
    }

    #[getter]
    fn diagnostic_count(&self) -> usize {
        self.diagnostics.len()
    }
}

#[pymethods]
impl PyLayoutBackedFieldUpdateReport {
    #[new]
    #[pyo3(signature = (*, page_fields, num_pages_fields, page_reference_fields, diagnostics, section_fields=0, section_pages_fields=0))]
    fn new(
        page_fields: usize,
        num_pages_fields: usize,
        page_reference_fields: usize,
        diagnostics: Vec<String>,
        section_fields: usize,
        section_pages_fields: usize,
    ) -> Self {
        Self {
            page_fields,
            num_pages_fields,
            page_reference_fields,
            section_fields,
            section_pages_fields,
            diagnostics,
        }
    }

    #[getter]
    fn page_fields(&self) -> usize {
        self.page_fields
    }

    #[getter]
    fn num_pages_fields(&self) -> usize {
        self.num_pages_fields
    }

    #[getter]
    fn page_reference_fields(&self) -> usize {
        self.page_reference_fields
    }

    #[getter]
    fn section_fields(&self) -> usize {
        self.section_fields
    }

    #[getter]
    fn section_pages_fields(&self) -> usize {
        self.section_pages_fields
    }

    #[getter]
    fn updated_count(&self) -> usize {
        self.page_fields
            + self.num_pages_fields
            + self.page_reference_fields
            + self.section_fields
            + self.section_pages_fields
    }

    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.diagnostics)
    }

    #[getter]
    fn diagnostic_count(&self) -> usize {
        self.diagnostics.len()
    }
}

/// One text field of the native core-properties model.
type CoreField = fn(&mut rdocx::CoreProperties) -> &mut Option<String>;

/// The longest text python-docx accepts for a core property.
const CORE_TEXT_LIMIT: usize = 255;

/// Read a W3CDTF date as python-docx does: a date and time, a date, a year and
/// month, or a year in the first nineteen characters, then an optional
/// `+hh:mm` or `-hh:mm` offset. Anything after the time that is not an offset,
/// such as `Z` or fractional seconds, is ignored.
fn w3cdtf_fields(value: &str) -> Option<([u32; 6], i32)> {
    let split = value
        .char_indices()
        .nth(19)
        .map_or(value.len(), |(index, _)| index);
    let (stamp, offset) = value.split_at(split);
    let bytes = stamp.as_bytes();
    let number = |start: usize, end: usize| {
        let digits = bytes.get(start..end)?;
        digits
            .iter()
            .all(u8::is_ascii_digit)
            .then(|| stamp[start..end].parse().ok())
            .flatten()
    };
    let separators = |expected: &[(usize, u8)]| {
        expected
            .iter()
            .all(|(index, byte)| bytes.get(*index) == Some(byte))
    };
    let fields = match bytes.len() {
        19 if separators(&[(4, b'-'), (7, b'-'), (10, b'T'), (13, b':'), (16, b':')]) => [
            number(0, 4)?,
            number(5, 7)?,
            number(8, 10)?,
            number(11, 13)?,
            number(14, 16)?,
            number(17, 19)?,
        ],
        10 if separators(&[(4, b'-'), (7, b'-')]) => {
            [number(0, 4)?, number(5, 7)?, number(8, 10)?, 0, 0, 0]
        }
        7 if separators(&[(4, b'-')]) => [number(0, 4)?, number(5, 7)?, 1, 0, 0, 0],
        4 => [number(0, 4)?, 1, 1, 0, 0, 0],
        _ => return None,
    };
    let minutes = if offset.len() == 6 {
        let offset = offset.as_bytes();
        let sign = match offset[0] {
            b'+' => 1,
            b'-' => -1,
            _ => return None,
        };
        if offset[3] != b':'
            || ![1, 2, 4, 5]
                .iter()
                .all(|index| offset[*index].is_ascii_digit())
        {
            return None;
        }
        let digit = |index: usize| i32::from(offset[index] - b'0');
        sign * ((digit(1) * 10 + digit(2)) * 60 + digit(4) * 10 + digit(5))
    } else {
        0
    };
    Some((fields, minutes))
}

/// Convert a stored W3CDTF date to an aware UTC `datetime`, or `None` when it
/// cannot be read, has an offset of a day or more, or falls outside the
/// `datetime` range once converted to UTC.
fn w3cdtf_to_datetime(py: Python<'_>, value: &str) -> PyResult<Option<Py<PyAny>>> {
    let Some(([year, month, day, hour, minute, second], offset)) = w3cdtf_fields(value) else {
        return Ok(None);
    };
    let datetime = py.import("datetime")?;
    let timezone = datetime.getattr("timezone")?;
    let utc = timezone.getattr("utc")?;
    let delta = datetime
        .getattr("timedelta")?
        .call((0, i64::from(offset) * 60), None)?;
    let constructor = datetime.getattr("datetime")?;
    let stamp = timezone
        .call1((delta,))
        .and_then(|zone| constructor.call1((year, month, day, hour, minute, second, 0, zone)))
        .and_then(|stamp| stamp.call_method1("astimezone", (utc,)));
    match stamp {
        Ok(stamp) => Ok(Some(stamp.unbind())),
        Err(error)
            if error.is_instance_of::<PyValueError>(py)
                || error.is_instance_of::<PyOverflowError>(py) =>
        {
            Ok(None)
        }
        Err(error) => Err(error),
    }
}

/// Write a `datetime` as a UTC W3CDTF date. A naive value is taken as UTC, as
/// python-docx does, and an aware one is converted to UTC first.
fn datetime_to_w3cdtf(value: &Bound<'_, PyAny>) -> PyResult<String> {
    let datetime = value.py().import("datetime")?;
    if !value.is_instance(&datetime.getattr("datetime")?)? {
        return Err(PyTypeError::new_err(
            "a core property date must be a datetime.datetime",
        ));
    }
    let value = if value.call_method0("utcoffset")?.is_none() {
        value.clone()
    } else {
        value.call_method1(
            "astimezone",
            (datetime.getattr("timezone")?.getattr("utc")?,),
        )?
    };
    let field = |name: &str| value.getattr(name)?.extract::<u32>();
    Ok(format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        field("year")?,
        field("month")?,
        field("day")?,
        field("hour")?,
        field("minute")?,
        field("second")?,
    ))
}

/// The package core properties (`docProps/core.xml`) under python-docx's
/// attribute names.
///
/// Text properties read as an empty string when absent, `revision` as zero,
/// and dates as `None`. Assigning `None` or empty text removes a property.
/// A write replaces the native model and creates the part, its package
/// relationship and its content type when the document has none. It changes
/// no content, so handles stay valid.
#[pyclass(name = "CoreProperties")]
pub struct PyCoreProperties {
    document: Py<PyDocument>,
}

impl PyCoreProperties {
    fn read(&self, py: Python<'_>, field: CoreField) -> Option<String> {
        let document = self.document.borrow(py);
        let mut properties = document.inner.core_properties()?.clone();
        field(&mut properties).take()
    }

    fn write(&self, py: Python<'_>, field: CoreField, value: Option<String>) -> PyResult<()> {
        let mut document = self.document.borrow_mut(py);
        let mut properties = document
            .inner
            .core_properties()
            .cloned()
            .unwrap_or_default();
        let value = value.filter(|value| !value.is_empty());
        if *field(&mut properties) == value {
            return Ok(());
        }
        *field(&mut properties) = value;
        document
            .inner
            .set_core_properties(properties)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn text(&self, py: Python<'_>, field: CoreField) -> String {
        self.read(py, field).unwrap_or_default()
    }

    fn set_text(&self, py: Python<'_>, field: CoreField, value: Option<String>) -> PyResult<()> {
        if value
            .as_deref()
            .is_some_and(|value| value.chars().count() > CORE_TEXT_LIMIT)
        {
            return Err(PyValueError::new_err(format!(
                "a core property holds at most {CORE_TEXT_LIMIT} characters"
            )));
        }
        self.write(py, field, value)
    }

    fn date(&self, py: Python<'_>, field: CoreField) -> PyResult<Option<Py<PyAny>>> {
        match self.read(py, field) {
            Some(value) => w3cdtf_to_datetime(py, &value),
            None => Ok(None),
        }
    }

    fn set_date(
        &self,
        py: Python<'_>,
        field: CoreField,
        value: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let value = value
            .filter(|value| !value.is_none())
            .map(datetime_to_w3cdtf)
            .transpose()?;
        self.write(py, field, value)
    }
}

#[pymethods]
impl PyCoreProperties {
    #[getter]
    fn author(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.creator)
    }
    #[setter]
    fn set_author(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.creator, value)
    }
    #[getter]
    fn category(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.category)
    }
    #[setter]
    fn set_category(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.category, value)
    }
    #[getter]
    fn comments(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.description)
    }
    #[setter]
    fn set_comments(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.description, value)
    }
    #[getter]
    fn content_status(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.content_status)
    }
    #[setter]
    fn set_content_status(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.content_status, value)
    }
    #[getter]
    fn created(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.date(py, |properties| &mut properties.created)
    }
    #[setter]
    fn set_created(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.set_date(py, |properties| &mut properties.created, value)
    }
    #[getter]
    fn identifier(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.identifier)
    }
    #[setter]
    fn set_identifier(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.identifier, value)
    }
    #[getter]
    fn keywords(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.keywords)
    }
    #[setter]
    fn set_keywords(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.keywords, value)
    }
    #[getter]
    fn language(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.language)
    }
    #[setter]
    fn set_language(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.language, value)
    }
    #[getter]
    fn last_modified_by(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.last_modified_by)
    }
    #[setter]
    fn set_last_modified_by(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.last_modified_by, value)
    }
    #[getter]
    fn last_printed(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.date(py, |properties| &mut properties.last_printed)
    }
    #[setter]
    fn set_last_printed(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.set_date(py, |properties| &mut properties.last_printed, value)
    }
    #[getter]
    fn modified(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.date(py, |properties| &mut properties.modified)
    }
    #[setter]
    fn set_modified(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.set_date(py, |properties| &mut properties.modified, value)
    }
    #[getter]
    fn revision(&self, py: Python<'_>) -> i64 {
        self.read(py, |properties| &mut properties.revision)
            .and_then(|value| value.trim().parse::<i64>().ok())
            .map_or(0, |value| value.max(0))
    }
    #[setter]
    fn set_revision(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        if value.is_some_and(|value| value < 1) {
            return Err(PyValueError::new_err("revision must be a positive integer"));
        }
        self.write(
            py,
            |properties| &mut properties.revision,
            value.map(|value| value.to_string()),
        )
    }
    #[getter]
    fn subject(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.subject)
    }
    #[setter]
    fn set_subject(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.subject, value)
    }
    #[getter]
    fn title(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.title)
    }
    #[setter]
    fn set_title(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.title, value)
    }
    #[getter]
    fn version(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.version)
    }
    #[setter]
    fn set_version(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.version, value)
    }
}

#[pyclass(name = "Document")]
pub struct PyDocument {
    pub(crate) inner: rdocx::Document,
    pub(crate) revisions: RevisionCounter,
    /// Advanced when a section is inserted or removed, which renumbers the
    /// sections that header and footer handles name.
    pub(crate) section_epoch: u64,
}

impl PyDocument {
    fn from_document(inner: rdocx::Document) -> Self {
        Self {
            inner,
            revisions: RevisionCounter::new(),
            section_epoch: 0,
        }
    }

    /// The live native owner that a `Story` snapshot names.
    ///
    /// A snapshot carries no fingerprint, so it resolves by kind, part name
    /// and owner index against the document as it is now.
    fn native_story(&self, py: Python<'_>, story: &PyStory) -> PyResult<rdocx::StoryId> {
        self.inner
            .stories()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .find(|candidate| story_snapshot(candidate) == *story)
            .ok_or_else(|| {
                rdocx_to_pyerr(
                    py,
                    rdocx::Error::Other(format!(
                        "document has no {} story {} at owner index {}",
                        story.kind, story.part_name, story.owner_index
                    )),
                )
            })
    }

    /// The live story of a `Hyperlink` record and the position of its link.
    ///
    /// A snapshot resolves at its recorded position when the story's link
    /// layout is unchanged and the link there still has the snapshot's
    /// fields. A constructed record resolves when exactly one link matches.
    fn native_hyperlink(
        &self,
        py: Python<'_>,
        hyperlink: &PyHyperlink,
    ) -> PyResult<(rdocx::StoryId, usize)> {
        let story = self.native_story(py, &hyperlink.story)?;
        let links = self
            .inner
            .story_links(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let matches = |(location, link): &(rdocx::ContentLocation, rdocx::LinkInfo)| {
            location.index_path() == hyperlink.index_path.as_slice()
                && link.text == hyperlink.text
                && link.url == hyperlink.url
                && link.anchor == hyperlink.anchor
                && link.rel_id == hyperlink.relationship_id
        };
        let index =
            match hyperlink.position {
                Some((index, layout)) => links
                    .get(index)
                    .filter(|entry| {
                        matches(entry)
                            && story_link_layout(links.iter().map(|(location, link)| {
                                (location.index_path(), link.text.as_str())
                            })) == layout
                    })
                    .map(|_| index),
                None => {
                    let mut found = links
                        .iter()
                        .enumerate()
                        .filter(|(_, entry)| matches(entry))
                        .map(|(index, _)| index);
                    match (found.next(), found.next()) {
                        (Some(index), None) => Some(index),
                        _ => None,
                    }
                }
            };
        index.map(|index| (story, index)).ok_or_else(|| {
            rdocx_to_pyerr(
                py,
                rdocx::Error::Other(
                    "the hyperlink does not match exactly one link of the document, re-fetch it with document.hyperlinks"
                        .to_owned(),
                ),
            )
        })
    }

    fn body_story(&self, py: Python<'_>) -> PyResult<rdocx::StoryId> {
        self.inner
            .stories()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .find(|story| story.kind() == rdocx::StoryKind::Body)
            .ok_or_else(|| {
                rdocx_to_pyerr(
                    py,
                    rdocx::Error::Other("document body story is missing".to_owned()),
                )
            })
    }

    fn native_location(
        &self,
        py: Python<'_>,
        item: &PyStoryItem,
    ) -> PyResult<rdocx::ContentLocation> {
        let story = self.native_story(py, &item.story)?;
        if item.revision != self.revisions.current() {
            return Err(crate::stale_to_pyerr(
                py,
                StaleElementError {
                    element_kind: "story item".to_owned(),
                    captured_revision: item.revision,
                    current_revision: self.revisions.current(),
                    recovery_hint: "Re-fetch it with document.story_items.".to_owned(),
                },
            ));
        }
        Ok(rdocx::ContentLocation::new(
            story,
            story_item_kind_from_name(&item.kind)?,
            item.index_path.clone(),
        ))
    }

    fn body_location(&self, py: Python<'_>, index: usize) -> PyResult<rdocx::ContentLocation> {
        self.body_locations(py, &[index])?
            .pop()
            .ok_or_else(|| PyIndexError::new_err("content index out of range"))
    }

    fn body_locations(
        &self,
        py: Python<'_>,
        indices: &[usize],
    ) -> PyResult<Vec<rdocx::ContentLocation>> {
        let content_count = self.inner.content_count();
        if indices.iter().any(|index| *index > content_count) {
            return Err(PyIndexError::new_err("content index out of range"));
        }
        let story = self.body_story(py)?;
        let snapshots = if indices.iter().all(|index| *index == content_count) {
            Vec::new()
        } else {
            self.inner
                .story_item_snapshots()
                .map_err(|error| rdocx_to_pyerr(py, error))?
        };
        let mut locations = Vec::with_capacity(indices.len());
        for index in indices {
            if *index == content_count {
                locations.push(rdocx::ContentLocation::end(story.clone()));
                continue;
            }
            let location = snapshots
                .iter()
                .find(|item| {
                    item.location().story() == &story && item.direct_body_index() == Some(*index)
                })
                .map(|item| item.location().clone())
                .ok_or_else(|| {
                    rdocx_to_pyerr(
                        py,
                        rdocx::Error::Other(format!(
                            "direct body content at index {index} has no checked location"
                        )),
                    )
                })?;
            locations.push(location);
        }
        Ok(locations)
    }

    fn comment_story_position(
        &self,
        py: Python<'_>,
        position: &rdocx::StoryRunPosition,
        items: &mut HashMap<rdocx::ContentLocation, rdocx::StoryItemSnapshot>,
    ) -> PyResult<PyStoryRunPosition> {
        if !items.contains_key(&position.location) {
            let snapshot = self
                .inner
                .story_range_paragraph_snapshot(&position.location)
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            items.insert(position.location.clone(), snapshot);
        }
        let item = &items[&position.location];
        Ok(PyStoryRunPosition {
            item: PyStoryItem {
                story: story_snapshot(item.location().story()),
                kind: story_item_kind_name(item.location().item_kind()).to_owned(),
                index_path: item.location().index_path().to_vec(),
                direct_body_index: item.direct_body_index(),
                text: item.text().map(str::to_owned),
                xml: item.xml().to_vec(),
                revision: self.revisions.current(),
            },
            run_index: position.run_index,
        })
    }

    fn story_item_snapshot(
        &self,
        py: Python<'_>,
        location: &rdocx::ContentLocation,
    ) -> PyResult<PyStoryItem> {
        let item = self
            .inner
            .story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .find(|item| item.location() == location)
            .ok_or_else(|| PyIndexError::new_err("inserted story item was not found"))?;
        Ok(PyStoryItem {
            story: story_snapshot(item.location().story()),
            kind: story_item_kind_name(item.location().item_kind()).to_owned(),
            index_path: item.location().index_path().to_vec(),
            direct_body_index: item.direct_body_index(),
            text: item.text().map(str::to_owned),
            xml: item.xml().to_vec(),
            revision: self.revisions.current(),
        })
    }

    /// Snapshot a checked body paragraph, including a paragraph inside a block control.
    fn paragraph_story_item(py: Python<'_>, paragraph: &PyParagraph) -> PyResult<PyStoryItem> {
        let paragraph_index = match paragraph.validate(py)? {
            ParagraphLocation::Body(index) => index,
            ParagraphLocation::Cell { .. } => {
                return Err(PyValueError::new_err(
                    "StoryRunPosition does not accept a table cell paragraph handle",
                ));
            }
            ParagraphLocation::Story { .. } => {
                return Err(PyValueError::new_err(
                    "StoryRunPosition does not accept a header or footer paragraph handle, \
                     build one from a StoryItem of document.story_items",
                ));
            }
        };
        let document = paragraph.document.borrow(py);
        let location = document
            .inner
            .paragraph_story_location(paragraph_index)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
        let [_, _] = location.index_path() else {
            return document.story_item_snapshot(py, &location);
        };
        let item = document
            .inner
            .story_range_paragraph_snapshot(&location)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(PyStoryItem {
            story: story_snapshot(item.location().story()),
            kind: story_item_kind_name(item.location().item_kind()).to_owned(),
            index_path: item.location().index_path().to_vec(),
            direct_body_index: item.direct_body_index(),
            text: item.text().map(str::to_owned),
            xml: item.xml().to_vec(),
            revision: document.revisions.current(),
        })
    }

    fn direct_content_index(
        slf: &Py<Self>,
        py: Python<'_>,
        content: &Bound<'_, PyAny>,
        argument: &str,
    ) -> PyResult<usize> {
        if let Ok(paragraph) = content.cast::<PyParagraph>() {
            let paragraph = paragraph.borrow();
            if !paragraph.belongs_to(py, slf) {
                return Err(PyValueError::new_err(
                    "content handle belongs to a different document",
                ));
            }
            let paragraph_index = match paragraph.validate(py)? {
                crate::paragraph::ParagraphLocation::Body(index) => index,
                crate::paragraph::ParagraphLocation::Cell { .. }
                | crate::paragraph::ParagraphLocation::Story { .. } => {
                    return Err(PyValueError::new_err(
                        "content handle is not a direct body child",
                    ));
                }
            };
            return slf
                .borrow(py)
                .inner
                .content_index_of_paragraph(paragraph_index)
                .ok_or_else(|| PyValueError::new_err("content handle is not a direct body child"));
        }
        if let Ok(table) = content.cast::<PyTable>() {
            let table = table.borrow();
            if !table.belongs_to(py, slf) {
                return Err(PyValueError::new_err(
                    "content handle belongs to a different document",
                ));
            }
            let table_index = table.validate(py)?;
            return slf
                .borrow(py)
                .inner
                .content_index_of_table(table_index)
                .ok_or_else(|| PyValueError::new_err("content handle is not a direct body child"));
        }
        Err(PyTypeError::new_err(format!(
            "{argument} must be a Paragraph or Table handle"
        )))
    }

    /// Split a run of the direct body paragraph at `body_index`. The revision
    /// advances only when a continuation run is created.
    fn split_body_run(
        &mut self,
        py: Python<'_>,
        body_index: usize,
        run_index: usize,
        character_offset: usize,
    ) -> PyResult<usize> {
        let paragraph_index = self.inner.paragraph_index_of_content(body_index);
        let run_count = |document: &rdocx::Document| {
            paragraph_index
                .and_then(|index| document.paragraph(index))
                .map(|paragraph| paragraph.run_count())
        };
        let before = run_count(&self.inner);
        let boundary = self
            .inner
            .split_run(body_index, run_index, character_offset)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if run_count(&self.inner) != before {
            self.revisions.bump();
        }
        Ok(boundary)
    }

    /// Run a native mutation that reports how many things it changed.
    ///
    /// The GIL is released while it runs, and live handles are staled only
    /// when the count is nonzero.
    fn counted_mutation<F>(&mut self, py: Python<'_>, mutation: F) -> PyResult<usize>
    where
        F: FnOnce(&mut rdocx::Document) -> rdocx::Result<usize> + Send,
    {
        let count = py
            .detach(|| mutation(&mut self.inner))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if count > 0 {
            self.revisions.bump();
        }
        Ok(count)
    }

    /// Run an ordered replacement batch that publishes only when every
    /// expected count holds, and stale live handles only when it replaced
    /// something.
    fn expected_replacements(
        &mut self,
        py: Python<'_>,
        pairs: &[(&str, &str, Option<usize>)],
        batch: bool,
    ) -> PyResult<Vec<usize>> {
        let counts = py
            .detach(|| self.inner.try_replace_all_expected(pairs))
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .map_err(|mismatch| crate::replacement_count_to_pyerr(py, &mismatch, batch))?;
        if counts.iter().any(|count| *count > 0) {
            self.revisions.bump();
        }
        Ok(counts)
    }

    // Both checked item and physical cell routes share singular count/error and
    // publication policy. The concrete closures are the two actual consumers.
    pub(crate) fn scoped_replacement<F>(&mut self, py: Python<'_>, mutation: F) -> PyResult<usize>
    where
        F: FnOnce(
                &mut rdocx::Document,
            ) -> rdocx::Result<Result<usize, rdocx::ReplacementCountMismatch>>
            + Send,
    {
        let count = py
            .detach(|| mutation(&mut self.inner))
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .map_err(|mismatch| crate::replacement_count_to_pyerr(py, &mismatch, false))?;
        if count > 0 {
            self.revisions.bump();
        }
        Ok(count)
    }

    /// Check the names and the section of one header or footer variant.
    fn section_story_target(
        &self,
        section_index: usize,
        kind: &str,
        variant: &str,
    ) -> PyResult<(rdocx::HeaderFooterKind, rdocx::HdrFtrType)> {
        let kind = header_footer_kind_from_name(kind)?;
        let variant = header_footer_variant_from_name(variant)?;
        if section_index >= self.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        Ok((kind, variant))
    }
}

fn story_snapshot(story: &rdocx::StoryId) -> PyStory {
    PyStory {
        kind: match story.kind() {
            rdocx::StoryKind::Body => "body",
            rdocx::StoryKind::TableCell => "table_cell",
            rdocx::StoryKind::Header => "header",
            rdocx::StoryKind::Footer => "footer",
            rdocx::StoryKind::Footnote => "footnote",
            rdocx::StoryKind::Endnote => "endnote",
            rdocx::StoryKind::Comment => "comment",
            rdocx::StoryKind::TextBox => "text_box",
            _ => "unknown",
        }
        .to_owned(),
        part_name: story.part_name().to_owned(),
        owner_index: story.owner_index(),
    }
}

fn story_item_kind_name(kind: rdocx::StoryItemKind) -> &'static str {
    match kind {
        rdocx::StoryItemKind::Paragraph => "paragraph",
        rdocx::StoryItemKind::Table => "table",
        rdocx::StoryItemKind::ContentControl => "content_control",
        rdocx::StoryItemKind::Field => "field",
        rdocx::StoryItemKind::Drawing => "drawing",
        rdocx::StoryItemKind::PreservedNode => "preserved_node",
        _ => "unknown",
    }
}

fn story_item_kind_from_name(name: &str) -> PyResult<rdocx::StoryItemKind> {
    match name {
        "paragraph" => Ok(rdocx::StoryItemKind::Paragraph),
        "table" => Ok(rdocx::StoryItemKind::Table),
        "content_control" => Ok(rdocx::StoryItemKind::ContentControl),
        "field" => Ok(rdocx::StoryItemKind::Field),
        "drawing" => Ok(rdocx::StoryItemKind::Drawing),
        "preserved_node" => Ok(rdocx::StoryItemKind::PreservedNode),
        _ => Err(PyValueError::new_err(format!(
            "unsupported story item kind {name:?}"
        ))),
    }
}

fn header_footer_kind_from_name(name: &str) -> PyResult<rdocx::HeaderFooterKind> {
    match name {
        "header" => Ok(rdocx::HeaderFooterKind::Header),
        "footer" => Ok(rdocx::HeaderFooterKind::Footer),
        _ => Err(PyValueError::new_err(format!(
            "kind must be header or footer, not {name:?}"
        ))),
    }
}

fn header_footer_variant_from_name(name: &str) -> PyResult<rdocx::HdrFtrType> {
    // `HdrFtrType::from_str` reads every unknown name as the default variant.
    match name {
        "default" => Ok(rdocx::HdrFtrType::Default),
        "first" => Ok(rdocx::HdrFtrType::First),
        "even" => Ok(rdocx::HdrFtrType::Even),
        _ => Err(PyValueError::new_err(format!(
            "variant must be default, first or even, not {name:?}"
        ))),
    }
}

/// Extract a direct body index, naming the accepted forms on a type error.
fn body_index_argument(value: &Bound<'_, PyAny>, message: &'static str) -> PyResult<usize> {
    value.extract::<usize>().map_err(|error| {
        if error.is_instance_of::<PyTypeError>(value.py()) {
            PyTypeError::new_err(message)
        } else {
            error
        }
    })
}

fn section_snapshot(
    section: rdocx::SectionRef<'_>,
    owner: Option<(Py<PyDocument>, SectionLayout)>,
) -> PySection {
    let (page_width, page_height) = section.page_size().map_or((None, None), |(width, height)| {
        (Some(width.to_emu()), Some(height.to_emu()))
    });
    let (margin_top, margin_right, margin_bottom, margin_left) =
        section
            .margins()
            .map_or((None, None, None, None), |(top, right, bottom, left)| {
                (
                    Some(top.to_emu()),
                    Some(right.to_emu()),
                    Some(bottom.to_emu()),
                    Some(left.to_emu()),
                )
            });
    let (column_count, column_spacing) =
        section.columns().map_or((None, None), |(count, spacing)| {
            (Some(count), Some(spacing.to_emu()))
        });
    let (header_distance, footer_distance) = section
        .header_footer_distance()
        .map_or((None, None), |(header, footer)| {
            (Some(header.to_emu()), Some(footer.to_emu()))
        });
    PySection {
        ordinal: section.ordinal(),
        is_final: section.is_final(),
        orientation: section.orientation().map(|value| value.to_str().to_owned()),
        page_width,
        page_height,
        margin_top,
        margin_right,
        margin_bottom,
        margin_left,
        gutter: section.gutter().map(rdocx::Length::to_emu),
        column_count,
        column_spacing,
        page_number_start: section.page_number_start(),
        header_distance,
        footer_distance,
        different_first_page: section.different_first_page(),
        break_type: section.break_type().map(|value| value.to_str().to_owned()),
        owner: owner.map_or_else(SectionOwner::default, |(document, layout)| {
            SectionOwner::of(document, layout)
        }),
    }
}

/// A section length the caller gave, its explicit value, or its layout default.
fn given_or_current_length(
    given: Option<i64>,
    current: Option<rdocx::Twips>,
    default: rdocx::Twips,
) -> rdocx::Length {
    match (given, current) {
        (Some(emu), _) => rdocx::Length::emu(emu),
        (None, Some(twips)) => rdocx::Length::twips(twips.0),
        (None, None) => rdocx::Length::twips(default.0),
    }
}

fn style_snapshot(style: rdocx::style::Style<'_>) -> PyStyle {
    PyStyle {
        style_id: style.style_id().to_owned(),
        name: style.name().map(str::to_owned),
        based_on: style.based_on().map(str::to_owned),
        style_type: style.style_type().to_str().to_owned(),
        linked_style: style.linked_style().map(str::to_owned),
        next_style: style.next_style().map(str::to_owned),
        priority: style.priority(),
        auto_redefine: style.auto_redefine(),
        hidden: style.hidden(),
        semi_hidden: style.semi_hidden(),
        unhide_when_used: style.unhide_when_used(),
        quick_format: style.quick_format(),
        locked: style.locked(),
        is_default: style.is_default(),
    }
}

#[pymethods]
impl PyDocument {
    #[new]
    #[pyo3(signature = (path = None))]
    fn new(path: Option<PathBuf>, py: Python<'_>) -> PyResult<Self> {
        match path {
            Some(path) => rdocx::Document::open(path)
                .map(Self::from_document)
                .map_err(|error| rdocx_to_pyerr(py, error)),
            None => Ok(Self::from_document(rdocx::Document::new())),
        }
    }

    #[staticmethod]
    fn open(path: PathBuf, py: Python<'_>) -> PyResult<Self> {
        rdocx::Document::open(path)
            .map(Self::from_document)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[staticmethod]
    fn from_bytes(bytes: &[u8], py: Python<'_>) -> PyResult<Self> {
        rdocx::Document::from_bytes(bytes)
            .map(Self::from_document)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn save(&mut self, path: PathBuf, py: Python<'_>) -> PyResult<()> {
        py.detach(|| self.inner.save(path))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[pyo3(name = "to_bytes")]
    fn serialize<'py>(&mut self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        py.detach(|| self.inner.to_bytes())
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[getter]
    fn core_properties(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyCoreProperties>> {
        Py::new(py, PyCoreProperties { document: slf })
    }

    #[getter]
    fn update_fields_on_open(&self) -> Option<bool> {
        self.inner.update_fields_on_open()
    }

    #[setter]
    fn set_update_fields_on_open(&mut self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.inner
            .set_update_fields_on_open(value)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn image_data<'py>(
        &self,
        py: Python<'py>,
        relationship_id: &str,
    ) -> Option<Bound<'py, PyBytes>> {
        self.inner
            .image_data(relationship_id)
            .map(|bytes| PyBytes::new(py, &bytes))
    }

    fn replace_image(
        &mut self,
        py: Python<'_>,
        relationship_id: &str,
        data: &[u8],
    ) -> PyResult<()> {
        // No content moves, so live handles stay valid.
        self.inner
            .replace_image(relationship_id, data)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn replace_image_for_story(
        &mut self,
        py: Python<'_>,
        story: PyRef<'_, PyStory>,
        relationship_id: &str,
        data: &[u8],
    ) -> PyResult<()> {
        let story = self.native_story(py, &story)?;
        self.inner
            .replace_image_for_story(&story, relationship_id, data)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn set_picture_size(
        &mut self,
        py: Python<'_>,
        relationship_id: &str,
        width: i64,
        height: i64,
    ) -> PyResult<usize> {
        // No content moves, so live handles stay valid.
        py.detach(|| {
            self.inner.set_picture_size(
                relationship_id,
                rdocx::Length::emu(width),
                rdocx::Length::emu(height),
            )
        })
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn split_run(
        slf: Py<Self>,
        py: Python<'_>,
        body_index: &Bound<'_, PyAny>,
        run_index: usize,
        character_offset: usize,
    ) -> PyResult<usize> {
        let Ok(paragraph) = body_index.cast::<PyParagraph>() else {
            // An integer is the direct body index, as in `RunPosition`.
            let body_index = body_index.extract::<usize>().map_err(|error| {
                if error.is_instance_of::<PyTypeError>(py) {
                    PyTypeError::new_err("body_index must be an int or a Paragraph handle")
                } else {
                    error
                }
            })?;
            return slf
                .borrow_mut(py)
                .split_body_run(py, body_index, run_index, character_offset);
        };
        let paragraph = paragraph.borrow();
        if !paragraph.belongs_to(py, &slf) {
            return Err(PyValueError::new_err(
                "paragraph handle belongs to a different document",
            ));
        }
        // A cell handle counts paragraphs inside cell content controls, which
        // `Cell::paragraph_mut` does not, so it could name another paragraph.
        let ParagraphLocation::Body(paragraph_index) = paragraph.validate(py)? else {
            return Err(PyValueError::new_err(
                "split_run does not accept a table cell paragraph handle",
            ));
        };
        let mut document = slf.borrow_mut(py);
        if let Some(body_index) = document.inner.content_index_of_paragraph(paragraph_index) {
            return document.split_body_run(py, body_index, run_index, character_offset);
        }
        // A paragraph inside a block content control has no direct body index.
        let run_count = |document: &rdocx::Document| {
            document
                .paragraph(paragraph_index)
                .map(|paragraph| paragraph.run_count())
        };
        let before = run_count(&document.inner);
        let boundary = document
            .inner
            .paragraph_mut(paragraph_index)
            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
            .split_run(run_index, character_offset)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if run_count(&document.inner) != before {
            document.revisions.bump();
        }
        Ok(boundary)
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, revision_view = "accepted"))]
    fn to_pdf<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        revision_view: &str,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let options = parse_render_options(revision_view)?;
        if fonts.is_none() && font_dir.is_none() {
            return py
                .detach(|| self.inner.to_pdf_with_options(options))
                .map(|bytes| PyBytes::new(py, &bytes))
                .map_err(|error| rdocx_to_pyerr(py, error));
        }
        let mut font_files = fonts
            .unwrap_or_default()
            .into_iter()
            .map(|(family, data)| (family, data.as_bytes().to_vec()))
            .collect::<Vec<_>>();
        // The native loader reads a missing directory as one without fonts.
        if let Some(font_dir) = &font_dir
            && !font_dir.is_dir()
        {
            return Err(if font_dir.exists() {
                PyNotADirectoryError::new_err(format!(
                    "font directory {} is not a directory",
                    font_dir.display()
                ))
            } else {
                PyFileNotFoundError::new_err(format!(
                    "font directory {} does not exist",
                    font_dir.display()
                ))
            });
        }
        py.detach(|| {
            if let Some(font_dir) = &font_dir {
                font_files.extend(
                    rdocx::Document::load_fonts_from_dir(font_dir)
                        .into_iter()
                        .map(|font| (font.family, font.data)),
                );
            }
            let font_files = font_files
                .iter()
                .map(|(family, data)| (family.as_str(), data.as_slice()))
                .collect::<Vec<_>>();
            self.inner
                .to_pdf_with_fonts_and_options(&font_files, options)
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[pyo3(signature = (profile = "pdfa-2b"))]
    fn to_pdfa_deterministic<'py>(
        &self,
        py: Python<'py>,
        profile: &str,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let profile = match profile {
            "pdfa-2b" => rdocx::PdfConformance::PdfA2b,
            "pdfa-3b" => rdocx::PdfConformance::PdfA3b,
            _ => {
                return Err(PyValueError::new_err(format!(
                    "profile must be pdfa-2b or pdfa-3b, not {profile:?}"
                )));
            }
        };
        py.detach(|| self.inner.to_pdfa_deterministic(profile))
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[pyo3(signature = (page_index, dpi = 150.0, *, revision_view = "accepted"))]
    fn render_page_to_png<'py>(
        &self,
        py: Python<'py>,
        page_index: usize,
        dpi: f64,
        revision_view: &str,
    ) -> PyResult<Option<Bound<'py, PyBytes>>> {
        let options = parse_render_options(revision_view)?;
        py.detach(|| {
            self.inner
                .render_page_to_png_with_options(page_index, dpi, options)
        })
        .map(|bytes| bytes.map(|bytes| PyBytes::new(py, &bytes)))
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn render_page_to_svg(
        &self,
        py: Python<'_>,
        page_index: usize,
    ) -> PyResult<Option<PySvgRenderResult>> {
        let rendered = py
            .detach(|| self.inner.render_page_to_svg(page_index))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(rendered.map(|result| PySvgRenderResult {
            svg: result.svg,
            diagnostics: result
                .diagnostics
                .into_iter()
                .map(|diagnostic| PySvgDiagnostic {
                    path: diagnostic.path,
                    message: diagnostic.message,
                })
                .collect(),
        }))
    }

    #[pyo3(signature = (dpi = 150.0, *, revision_view = "accepted"))]
    fn render_all_pages<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        revision_view: &str,
    ) -> PyResult<Bound<'py, PyList>> {
        let options = parse_render_options(revision_view)?;
        let pages = py
            .detach(|| self.inner.render_all_pages_with_options(dpi, options))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        PyList::new(py, pages.iter().map(|page| PyBytes::new(py, page)))
    }

    #[pyo3(signature = (*, dpi = 150.0, format = "png", quality = 90, transparent = false, pages = None, revision_view = "accepted"))]
    #[allow(clippy::too_many_arguments)]
    fn render_pages(
        &self,
        py: Python<'_>,
        dpi: f64,
        format: &str,
        quality: u8,
        transparent: bool,
        pages: Option<Vec<usize>>,
        revision_view: &str,
    ) -> PyResult<Py<PyAny>> {
        let options = parse_render_options(revision_view)?;
        let rendered = py
            .detach(|| {
                let format = parse_raster_format(format, quality, transparent)?;
                let selected = match pages {
                    Some(pages) => pages,
                    None => {
                        (0..self.inner.layout_with_options(options)?.layout.pages.len()).collect()
                    }
                };
                self.inner.render_pages_with_options(
                    &selected,
                    rdocx::RasterOptions { dpi, format },
                    options,
                )
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        match rendered {
            rdocx::RasterOutput::SeparatePages(pages) => {
                let list = PyList::new(py, pages.iter().map(|page| PyBytes::new(py, page)))?;
                Ok(list.into_any().unbind())
            }
            rdocx::RasterOutput::MultiPageTiff(tiff) => {
                Ok(PyBytes::new(py, &tiff).into_any().unbind())
            }
        }
    }

    #[pyo3(signature = (
        edited,
        author,
        timestamp,
        *,
        granularity = "run",
        ignore_formatting = false,
        ignore_whitespace = false,
        ignore_fields = false,
        ignore_comments = false,
        ignored_stories = None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn compare<'py>(
        &mut self,
        edited: &PyDocument,
        author: &str,
        timestamp: &str,
        py: Python<'py>,
        granularity: &str,
        ignore_formatting: bool,
        ignore_whitespace: bool,
        ignore_fields: bool,
        ignore_comments: bool,
        ignored_stories: Option<Vec<String>>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let (diagnostics, changed) = py
            .detach(|| {
                let options = rdocx::ComparisonOptions {
                    granularity: parse_comparison_granularity(granularity)?,
                    ignore_formatting,
                    ignore_whitespace,
                    ignore_fields,
                    ignore_comments,
                    ignored_stories: ignored_stories
                        .iter()
                        .flatten()
                        .map(|name| parse_comparison_story(name))
                        .collect::<rdocx::Result<_>>()?,
                };
                let before = self.inner.to_bytes()?;
                let diagnostics =
                    self.inner
                        .compare_with_options(&edited.inner, author, timestamp, &options)?;
                let changed = self.inner.to_bytes()? != before;
                Ok::<_, rdocx::Error>((diagnostics, changed))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if changed {
            self.revisions.bump();
        }
        PyTuple::new(
            py,
            diagnostics.into_iter().map(|item| PyComparisonDiagnostic {
                location: item.location,
                message: item.message,
            }),
        )
    }

    #[getter]
    fn comments<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let mut anchors = self
            .inner
            .comment_anchor_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut items = self
            .inner
            .story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .map(|item| (item.location().clone(), item))
            .collect::<HashMap<_, _>>();
        let records = self
            .inner
            .comments()
            .into_iter()
            .map(|comment| {
                let (range, anchor_text) = anchors.remove(&comment.id()).ok_or_else(|| {
                    PyValueError::new_err(
                        "comment snapshot is missing its checked ownership record",
                    )
                })?;
                let anchor = range
                    .map(|range| -> PyResult<PyStoryRunRange> {
                        Ok(PyStoryRunRange {
                            start: self.comment_story_position(py, &range.start, &mut items)?,
                            end: self.comment_story_position(py, &range.end, &mut items)?,
                        })
                    })
                    .transpose()?;
                Ok(PyComment {
                    id: comment.id(),
                    author: comment.author().map(str::to_owned),
                    initials: comment.initials().map(str::to_owned),
                    date: comment.date().map(str::to_owned),
                    text: comment.text(),
                    parent_id: comment.parent_id(),
                    resolved: comment.resolved(),
                    anchor_text,
                    anchor,
                })
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, records)
    }

    #[getter]
    fn sections<'py>(slf: &Bound<'py, Self>) -> PyResult<Bound<'py, PyTuple>> {
        let py = slf.py();
        let this = slf.borrow();
        let layout = SectionLayout::of(&this);
        let sections = this
            .inner
            .sections()
            .map(|section| section_snapshot(section, Some((slf.clone().unbind(), layout))))
            .collect::<Vec<_>>();
        PyTuple::new(py, sections)
    }

    #[pyo3(signature = (
        index,
        *,
        orientation = None,
        page_width = None,
        page_height = None,
        margin_top = None,
        margin_right = None,
        margin_bottom = None,
        margin_left = None,
        gutter = None,
        column_count = None,
        column_spacing = None,
        page_number_start = None,
        header_distance = None,
        footer_distance = None,
        different_first_page = None,
        break_type = None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn update_section(
        mut slf: PyRefMut<'_, Self>,
        py: Python<'_>,
        index: usize,
        orientation: Option<&str>,
        page_width: Option<i64>,
        page_height: Option<i64>,
        margin_top: Option<i64>,
        margin_right: Option<i64>,
        margin_bottom: Option<i64>,
        margin_left: Option<i64>,
        gutter: Option<i64>,
        column_count: Option<u32>,
        column_spacing: Option<i64>,
        page_number_start: Option<u32>,
        header_distance: Option<i64>,
        footer_distance: Option<i64>,
        different_first_page: Option<bool>,
        break_type: Option<&str>,
    ) -> PyResult<PySection> {
        let orientation = orientation
            .map(|value| {
                ST_PageOrientation::from_str(value)
                    .map_err(|_| PyValueError::new_err("orientation must be portrait or landscape"))
            })
            .transpose()?;
        let break_type = break_type
            .map(|value| {
                ST_SectionType::from_str(value).map_err(|_| {
                    PyValueError::new_err(
                        "break type must be nextPage, continuous, evenPage, oddPage or nextColumn",
                    )
                })
            })
            .transpose()?;
        // Section lengths are EMU. A bare int meant as twips or points is
        // hundreds of times too small, so a page under Word's minimum of a
        // tenth of an inch, or a margin, gutter, distance or column gap under
        // a hundredth of an inch, is refused rather than written.
        for (name, value, minimum) in [
            ("page_width", page_width, 91_440),
            ("page_height", page_height, 91_440),
            ("margin_top", margin_top, 9_144),
            ("margin_right", margin_right, 9_144),
            ("margin_bottom", margin_bottom, 9_144),
            ("margin_left", margin_left, 9_144),
            ("gutter", gutter, 9_144),
            ("column_spacing", column_spacing, 9_144),
            ("header_distance", header_distance, 9_144),
            ("footer_distance", footer_distance, 9_144),
        ] {
            if let Some(value) = value
                && (1..minimum).contains(&value)
            {
                return Err(PyValueError::new_err(format!(
                    "{name}={value} is {:.4} inch: section lengths are EMU, pass a Length such \
                     as Inches(1), Pt(72) or Twips(1440)",
                    value as f64 / 914_400.0
                )));
            }
        }
        let mut section = slf
            .inner
            .section_mut(index)
            .ok_or_else(|| PyIndexError::new_err("section index out of range"))?;

        let current = section.properties();
        let defaults = rdocx_oxml::CT_SectPr::default_letter();
        let page_size = if page_width.is_some() || page_height.is_some() {
            Some((
                given_or_current_length(
                    page_width,
                    current.page_width,
                    defaults.page_width.unwrap(),
                ),
                given_or_current_length(
                    page_height,
                    current.page_height,
                    defaults.page_height.unwrap(),
                ),
            ))
        } else {
            None
        };
        let margins = if [margin_top, margin_right, margin_bottom, margin_left]
            .iter()
            .any(Option::is_some)
        {
            Some((
                given_or_current_length(
                    margin_top,
                    current.margin_top,
                    defaults.margin_top.unwrap(),
                ),
                given_or_current_length(
                    margin_right,
                    current.margin_right,
                    defaults.margin_right.unwrap(),
                ),
                given_or_current_length(
                    margin_bottom,
                    current.margin_bottom,
                    defaults.margin_bottom.unwrap(),
                ),
                given_or_current_length(
                    margin_left,
                    current.margin_left,
                    defaults.margin_left.unwrap(),
                ),
            ))
        } else {
            None
        };
        let columns = if column_count.is_some() || column_spacing.is_some() {
            let current = section.columns();
            if current.is_none()
                && section.properties().columns.is_some()
                && (column_count.is_none() || column_spacing.is_none())
            {
                return Err(PyValueError::new_err(
                    "unequal-width columns require both column_count and column_spacing",
                ));
            }
            let count = column_count
                .or(current.map(|(count, _)| count))
                .unwrap_or(1);
            let spacing = given_or_current_length(
                column_spacing,
                current.map(|(_, spacing)| spacing.as_twips()),
                rdocx::Twips(720),
            );
            Some((count, spacing))
        } else {
            None
        };
        let distances = if header_distance.is_some() || footer_distance.is_some() {
            Some((
                given_or_current_length(
                    header_distance,
                    current.header_distance,
                    defaults.header_distance.unwrap(),
                ),
                given_or_current_length(
                    footer_distance,
                    current.footer_distance,
                    defaults.footer_distance.unwrap(),
                ),
            ))
        } else {
            None
        };

        // The native setters check one value at a time, so restore the
        // section if a later value is rejected after an earlier one applied.
        let saved = section.properties().clone();
        let applied = (|| -> rdocx::Result<()> {
            if let Some((width, height)) = page_size {
                section.set_page_size(width, height)?;
            }
            if let Some(orientation) = orientation {
                section.set_orientation(orientation);
            }
            if let Some((top, right, bottom, left)) = margins {
                section.set_margins(top, right, bottom, left)?;
            }
            if let Some(gutter) = gutter {
                section.set_gutter(rdocx::Length::emu(gutter))?;
            }
            if let Some((count, spacing)) = columns {
                section.set_columns(count, spacing)?;
            }
            if let Some(start) = page_number_start {
                section.set_page_number_start(start)?;
            }
            if let Some((header, footer)) = distances {
                section.set_header_footer_distance(header, footer)?;
            }
            if let Some(enabled) = different_first_page {
                section.set_different_first_page(enabled);
            }
            if let Some(break_type) = break_type {
                section.set_break_type(break_type);
            }
            Ok(())
        })();
        if let Err(error) = applied {
            *section.properties_mut() = saved;
            return Err(rdocx_to_pyerr(py, error));
        }
        let mut updated = slf
            .inner
            .sections()
            .nth(index)
            .map(|section| section_snapshot(section, None))
            .expect("the updated section exists");
        let layout = SectionLayout::of(&slf);
        updated.owner = SectionOwner::of(slf.into(), layout);
        Ok(updated)
    }

    fn insert_section(&mut self, py: Python<'_>, index: usize) -> PyResult<()> {
        if index > self.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        self.inner
            .insert_section(index)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        self.section_epoch += 1;
        Ok(())
    }

    fn remove_section(&mut self, py: Python<'_>, index: usize) -> PyResult<()> {
        if index >= self.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        self.inner
            .remove_section(index)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        self.section_epoch += 1;
        Ok(())
    }

    #[getter]
    fn styles<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.inner.styles().into_iter().map(style_snapshot))
    }

    // Create a style as python-docx's `styles.add_style` does, deriving the
    // ID from the name unless one is given, with the formatting given as
    // keywords. Styles change no content, so handles stay valid.
    #[pyo3(signature = (
        name,
        style_type = "paragraph",
        *,
        style_id = None,
        based_on = None,
        next_style = None,
        font_name = None,
        font_size = None,
        bold = None,
        italic = None,
        color = None,
        space_before = None,
        space_after = None,
        left_indent = None,
        right_indent = None,
        first_line_indent = None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn add_style(
        &mut self,
        py: Python<'_>,
        name: &str,
        style_type: &str,
        style_id: Option<&str>,
        based_on: Option<&str>,
        next_style: Option<&str>,
        font_name: Option<String>,
        font_size: Option<i64>,
        bold: Option<bool>,
        italic: Option<bool>,
        color: Option<(u8, u8, u8)>,
        space_before: Option<i64>,
        space_after: Option<i64>,
        left_indent: Option<i64>,
        right_indent: Option<i64>,
        first_line_indent: Option<i64>,
    ) -> PyResult<PyStyle> {
        type NewStyle = fn(&str, &str) -> rdocx::StyleBuilder;
        let (native_type, new_style): (rdocx::StyleType, NewStyle) = match style_type {
            "paragraph" => (rdocx::StyleType::Paragraph, rdocx::StyleBuilder::paragraph),
            "character" => (rdocx::StyleType::Character, rdocx::StyleBuilder::character),
            "table" => (rdocx::StyleType::Table, rdocx::StyleBuilder::table),
            _ => {
                return Err(PyValueError::new_err(
                    "style_type must be 'paragraph', 'character' or 'table'",
                ));
            }
        };
        let style_id =
            style_id.map_or_else(|| style_id_from_name(&self.inner, name), str::to_owned);
        if name.trim().is_empty() || style_id.trim().is_empty() {
            return Err(PyValueError::new_err(
                "a style name and style ID cannot be blank",
            ));
        }
        if self.inner.style(&style_id).is_some() {
            return Err(PyValueError::new_err(format!(
                "a style with the ID '{style_id}' already exists"
            )));
        }
        let lowered = name.to_lowercase();
        if self.inner.styles().iter().any(|style| {
            style
                .name()
                .is_some_and(|name| name.to_lowercase() == lowered)
        }) {
            return Err(PyValueError::new_err(format!(
                "a style named '{name}' already exists"
            )));
        }
        let formatting = StyleFormatting {
            font_name,
            font_size,
            bold,
            italic,
            color,
            space_before,
            space_after,
            left_indent,
            right_indent,
            first_line_indent,
        };
        let mut builder = new_style(&style_id, name);
        if let Some(parent) = based_on {
            builder = builder.based_on(&style_id_of_type(&self.inner, parent, native_type)?);
        }
        if let Some(next) = next_style {
            if native_type != rdocx::StyleType::Paragraph {
                return Err(PyValueError::new_err(
                    "only a paragraph style has a next style",
                ));
            }
            builder = builder.next_style(&style_id_of_type(
                &self.inner,
                next,
                rdocx::StyleType::Paragraph,
            )?);
        }
        if let Some(properties) = formatting.paragraph_properties() {
            if native_type == rdocx::StyleType::Character {
                return Err(PyValueError::new_err(
                    "a character style takes no paragraph formatting",
                ));
            }
            builder = builder.paragraph_properties(properties);
        }
        if let Some(properties) = formatting.run_properties() {
            builder = builder.run_properties(properties);
        }
        self.inner
            .add_style(builder)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(style_snapshot(
            self.inner
                .style(&style_id)
                .expect("the style was just added"),
        ))
    }

    // Update only the supplied formatting on a defined style. The native
    // setter stages and validates the complete style graph before publishing.
    #[pyo3(signature = (
        style,
        *,
        based_on = None,
        next_style = None,
        font_name = None,
        font_size = None,
        bold = None,
        italic = None,
        color = None,
        space_before = None,
        space_after = None,
        left_indent = None,
        right_indent = None,
        first_line_indent = None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn set_style(
        &mut self,
        py: Python<'_>,
        style: &str,
        based_on: Option<&str>,
        next_style: Option<&str>,
        font_name: Option<String>,
        font_size: Option<i64>,
        bold: Option<bool>,
        italic: Option<bool>,
        color: Option<(u8, u8, u8)>,
        space_before: Option<i64>,
        space_after: Option<i64>,
        left_indent: Option<i64>,
        right_indent: Option<i64>,
        first_line_indent: Option<i64>,
    ) -> PyResult<PyStyle> {
        let (style_type, style_id) = defined_style(&self.inner, style)?;
        let existing = self.inner.style(&style_id).expect("the style was found");
        let name = existing.name().unwrap_or(&style_id).to_owned();
        let mut builder = match style_type {
            rdocx::StyleType::Paragraph => rdocx::StyleBuilder::paragraph(&style_id, &name),
            rdocx::StyleType::Character => rdocx::StyleBuilder::character(&style_id, &name),
            rdocx::StyleType::Table => rdocx::StyleBuilder::table(&style_id, &name),
            _ => return Err(PyValueError::new_err("unsupported style type")),
        };
        if let Some(parent) = based_on {
            builder = builder.based_on(&style_id_of_type(&self.inner, parent, style_type)?);
        }
        if let Some(next) = next_style {
            if style_type != rdocx::StyleType::Paragraph {
                return Err(PyValueError::new_err(
                    "only a paragraph style has a next style",
                ));
            }
            builder = builder.next_style(&style_id_of_type(
                &self.inner,
                next,
                rdocx::StyleType::Paragraph,
            )?);
        }
        let formatting = StyleFormatting {
            font_name,
            font_size,
            bold,
            italic,
            color,
            space_before,
            space_after,
            left_indent,
            right_indent,
            first_line_indent,
        };
        if let Some(properties) = formatting.paragraph_properties() {
            if style_type == rdocx::StyleType::Character {
                return Err(PyValueError::new_err(
                    "a character style takes no paragraph formatting",
                ));
            }
            builder = builder.paragraph_properties(properties);
        }
        if let Some(properties) = formatting.run_properties() {
            builder = builder.run_properties(properties);
        }
        self.inner
            .set_style(builder)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(style_snapshot(
            self.inner
                .style(&style_id)
                .expect("the style was just updated"),
        ))
    }

    // Return false when no style has the ID or name. The native call refuses
    // a style that content, another style or a numbering level still uses.
    fn remove_style(&mut self, py: Python<'_>, style: &str) -> PyResult<bool> {
        let Ok((_, style_id)) = defined_style(&self.inner, style) else {
            return Ok(false);
        };
        self.inner
            .remove_style(&style_id)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn set_default_style(&mut self, py: Python<'_>, style: &str) -> PyResult<()> {
        let (style_type, style_id) = defined_style(&self.inner, style)?;
        self.inner
            .set_default_style(style_type, &style_id)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn add_numbering_definition(
        &mut self,
        py: Python<'_>,
        levels: Vec<PyRef<'_, PyListLevel>>,
    ) -> PyResult<u32> {
        let levels = levels
            .iter()
            .map(|level| level.native())
            .collect::<Vec<_>>();
        self.inner
            .add_numbering_definition(&levels)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn add_numbering_instance(&mut self, py: Python<'_>, definition_id: u32) -> PyResult<u32> {
        self.inner
            .add_numbering_instance(definition_id, &[])
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn link_style_to_numbering(
        &mut self,
        py: Python<'_>,
        style: &str,
        num_id: u32,
        level: u32,
    ) -> PyResult<()> {
        let style_id = style_id_of_type(&self.inner, style, rdocx::StyleType::Paragraph)?;
        self.inner
            .link_style_to_numbering(&style_id, num_id, level)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[getter]
    fn stories<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let stories = self
            .inner
            .stories()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        PyTuple::new(py, stories.iter().map(story_snapshot))
    }

    #[getter]
    fn story_items<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let items = self
            .inner
            .story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let snapshots = items
            .into_iter()
            .map(|item| PyStoryItem {
                story: story_snapshot(item.location().story()),
                kind: story_item_kind_name(item.location().item_kind()).to_owned(),
                index_path: item.location().index_path().to_vec(),
                direct_body_index: item.direct_body_index(),
                text: item.text().map(str::to_owned),
                xml: item.xml().to_vec(),
                revision: self.revisions.current(),
            })
            .collect::<Vec<_>>();
        PyTuple::new(py, snapshots)
    }

    #[getter]
    fn header_footer_variants<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let mut snapshots = Vec::new();
        for section_index in 0..self.inner.section_count() {
            for (kind, kind_name) in [
                (rdocx::HeaderFooterKind::Header, "header"),
                (rdocx::HeaderFooterKind::Footer, "footer"),
            ] {
                for variant in [
                    rdocx::HdrFtrType::Default,
                    rdocx::HdrFtrType::First,
                    rdocx::HdrFtrType::Even,
                ] {
                    let resolved = self
                        .inner
                        .section_story(section_index, kind, variant)
                        .map_err(|error| rdocx_to_pyerr(py, error))?;
                    snapshots.push(PyHeaderFooterVariant {
                        section_index,
                        kind: kind_name.to_owned(),
                        variant: variant.to_str().to_owned(),
                        story: resolved.as_ref().map(|value| story_snapshot(value.story())),
                        source_section: resolved.as_ref().map(rdocx::SectionStory::source_section),
                        inherited: resolved.is_some_and(|value| value.is_inherited()),
                    });
                }
            }
        }
        PyTuple::new(py, snapshots)
    }

    // The section story operations publish a reopened package, so each one
    // advances the revision once, even when the variant already had a story.
    fn create_section_story(
        &mut self,
        py: Python<'_>,
        section_index: usize,
        kind: &str,
        variant: &str,
    ) -> PyResult<PyStory> {
        let (kind, variant) = self.section_story_target(section_index, kind, variant)?;
        let story = py
            .detach(|| {
                self.inner
                    .create_section_story(section_index, kind, variant)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(story_snapshot(&story))
    }

    fn link_section_story(
        &mut self,
        py: Python<'_>,
        section_index: usize,
        kind: &str,
        variant: &str,
        story: PyRef<'_, PyStory>,
    ) -> PyResult<PyStory> {
        let (kind, variant) = self.section_story_target(section_index, kind, variant)?;
        let story = self.native_story(py, &story)?;
        let linked = py
            .detach(|| {
                self.inner
                    .link_section_story(section_index, kind, variant, &story)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(story_snapshot(&linked))
    }

    fn unlink_section_story(
        &mut self,
        py: Python<'_>,
        section_index: usize,
        kind: &str,
        variant: &str,
    ) -> PyResult<PyStory> {
        let (kind, variant) = self.section_story_target(section_index, kind, variant)?;
        let story = py
            .detach(|| {
                self.inner
                    .unlink_section_story(section_index, kind, variant)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(story_snapshot(&story))
    }

    #[getter]
    fn hyperlinks<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let links = self
            .inner
            .story_link_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut story_links = HashMap::<_, Vec<_>>::new();
        for (location, link) in &links {
            story_links
                .entry(location.story().clone())
                .or_default()
                .push((location.index_path(), link.text.as_str()));
        }
        let mut layouts = story_links
            .into_iter()
            .map(|(story, links)| (story, (0usize, story_link_layout(links))))
            .collect::<HashMap<_, _>>();
        let snapshots = links
            .into_iter()
            .map(|(location, link)| {
                let (next, layout) = layouts
                    .get_mut(location.story())
                    .expect("every snapshot story has a layout");
                let position = Some((*next, *layout));
                *next += 1;
                PyHyperlink {
                    story: story_snapshot(location.story()),
                    index_path: location.index_path().to_vec(),
                    text: link.text,
                    url: link.url,
                    anchor: link.anchor,
                    relationship_id: link.rel_id,
                    position,
                }
            })
            .collect::<Vec<_>>();
        PyTuple::new(py, snapshots)
    }

    fn set_hyperlink_url(
        &mut self,
        py: Python<'_>,
        hyperlink: PyRef<'_, PyHyperlink>,
        url: &str,
    ) -> PyResult<()> {
        let (story, link_index) = self.native_hyperlink(py, &hyperlink)?;
        // No content moves, so live handles stay valid.
        py.detach(|| self.inner.set_hyperlink_url(&story, link_index, url))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn remove_hyperlink(
        &mut self,
        py: Python<'_>,
        hyperlink: PyRef<'_, PyHyperlink>,
    ) -> PyResult<()> {
        let (story, link_index) = self.native_hyperlink(py, &hyperlink)?;
        // The runs stay in their paragraph, so live handles stay valid.
        py.detach(|| self.inner.remove_hyperlink(&story, link_index))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn set_header(&mut self, py: Python<'_>, text: &str) -> PyResult<()> {
        py.detach(|| self.inner.try_set_header(text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    fn set_footer(&mut self, py: Python<'_>, text: &str) -> PyResult<()> {
        py.detach(|| self.inner.try_set_footer(text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    fn set_story_text(
        &mut self,
        py: Python<'_>,
        item: PyRef<'_, PyStoryItem>,
        text: &str,
    ) -> PyResult<()> {
        let location = self.native_location(py, &item)?;
        py.detach(|| self.inner.set_story_text(&location, text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    fn add_hyperlink_to_story(
        &mut self,
        py: Python<'_>,
        story: PyRef<'_, PyStory>,
        text: &str,
        url: &str,
    ) -> PyResult<()> {
        let story = self.native_story(py, &story)?;
        py.detach(|| self.inner.add_hyperlink_to_story(&story, text, url))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    #[pyo3(signature = (data, filename, width=None, height=None, *, after=None))]
    fn add_picture(
        &mut self,
        py: Python<'_>,
        data: &[u8],
        filename: &str,
        width: Option<i64>,
        height: Option<i64>,
        after: Option<PyRef<'_, PyStoryItem>>,
    ) -> PyResult<PyStoryItem> {
        let after = after
            .as_deref()
            .map(|item| self.native_location(py, item))
            .transpose()?;
        let story = match after.as_ref() {
            Some(location) => location.story().clone(),
            None => self.body_story(py)?,
        };
        let location = py
            .detach(|| {
                self.inner.insert_picture_to_story(
                    &story,
                    after.as_ref(),
                    data,
                    filename,
                    width.map(rdocx::Length::emu),
                    height.map(rdocx::Length::emu),
                )
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        self.story_item_snapshot(py, &location)
    }

    #[pyo3(signature = (range, *, author, text, initials = None, date = None))]
    fn add_comment(
        &mut self,
        range: &Bound<'_, PyAny>,
        author: &str,
        text: &str,
        initials: Option<&str>,
        date: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<i32> {
        let id = if let Ok(range) = range.cast::<PyRunRange>() {
            self.inner
                .add_comment_with_date((*range.borrow()).into(), author, initials, text, date)
        } else if let Ok(range) = range.cast::<PyStoryRunRange>() {
            let range = range.borrow();
            let start = rdocx::StoryRunPosition {
                location: self.native_location(py, &range.start.item)?,
                run_index: range.start.run_index,
            };
            let end = rdocx::StoryRunPosition {
                location: self.native_location(py, &range.end.item)?,
                run_index: range.end.run_index,
            };
            self.inner.add_story_comment_with_date(
                rdocx::StoryRunRange { start, end },
                author,
                initials,
                text,
                date,
            )
        } else {
            return Err(PyTypeError::new_err(
                "range must be a RunRange or StoryRunRange",
            ));
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(id)
    }

    /// Comment on the `occurrence`-th match of `anchor`, counted from zero,
    /// in the main story. Matching is case-sensitive, non-overlapping and
    /// within one paragraph. A match whose range would also show other text,
    /// such as a field result, raises.
    #[pyo3(signature = (anchor, *, author, text, occurrence = 0, initials = None, date = None))]
    #[allow(clippy::too_many_arguments)]
    fn add_comment_on_text(
        &mut self,
        anchor: &str,
        author: &str,
        text: &str,
        occurrence: usize,
        initials: Option<&str>,
        date: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<i32> {
        let id = self
            .inner
            .add_comment_on_text(anchor, occurrence, author, initials, text, date)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(id)
    }

    /// Move a root comment to a checked current story range atomically.
    fn move_comment(
        &mut self,
        id: i32,
        range: PyRef<'_, PyStoryRunRange>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let start = rdocx::StoryRunPosition {
            location: self.native_location(py, &range.start.item)?,
            run_index: range.start.run_index,
        };
        let end = rdocx::StoryRunPosition {
            location: self.native_location(py, &range.end.item)?,
            run_index: range.end.run_index,
        };
        py.detach(|| {
            self.inner
                .move_comment(id, rdocx::StoryRunRange { start, end })
        })
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    /// Move a root comment onto the selected literal main-story occurrence.
    #[pyo3(signature = (id, anchor, *, occurrence = 0))]
    fn move_comment_to_text(
        &mut self,
        id: i32,
        anchor: &str,
        occurrence: usize,
        py: Python<'_>,
    ) -> PyResult<()> {
        py.detach(|| self.inner.move_comment_to_text(id, anchor, occurrence))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    #[pyo3(signature = (parent_id, *, author, text, date = None))]
    fn reply_to(
        &mut self,
        parent_id: i32,
        author: &str,
        text: &str,
        date: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<i32> {
        let id = self
            .inner
            .reply_to_with_date(parent_id, author, text, date)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(id)
    }

    #[pyo3(signature = (id, *, resolved = true))]
    fn resolve_comment(&mut self, id: i32, resolved: bool, py: Python<'_>) -> PyResult<bool> {
        let updated = self
            .inner
            .resolve_comment(id, resolved)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if updated {
            self.revisions.bump();
        }
        Ok(updated)
    }

    fn remove_comment(&mut self, id: i32, py: Python<'_>) -> PyResult<bool> {
        let removed = self
            .inner
            .remove_comment(id)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if removed {
            self.revisions.bump();
        }
        Ok(removed)
    }

    #[getter]
    fn bookmarks<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.inner
                .bookmarks()
                .into_iter()
                .map(|bookmark| PyBookmark {
                    id: bookmark.id(),
                    name: bookmark.name().map(str::to_owned),
                    range: bookmark.range().map(PyRunRange::from),
                    direct_range: bookmark.direct_range().map(PyRunRange::from),
                    text: bookmark.text().to_owned(),
                    issue: bookmark.issue().map(str::to_owned),
                }),
        )
    }

    fn add_bookmark(
        &mut self,
        py: Python<'_>,
        name: &str,
        range: PyRef<'_, PyRunRange>,
    ) -> PyResult<i32> {
        // Bookmark markers sit between runs, so no content moves and live
        // handles stay valid.
        self.inner
            .add_bookmark(name, (*range).into())
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn layout<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let fragments = py
            .detach(|| {
                let layout = self.inner.layout_deterministic()?;
                let mut fragments = Vec::new();
                for body_index in 0..self.inner.content_count() {
                    let Some(body_fragments) = layout.body_layout_fragments(body_index) else {
                        continue;
                    };
                    fragments.extend(body_fragments.iter().map(|fragment| PyLayoutFragment {
                        body_index,
                        physical_page: fragment.physical_page,
                        displayed_page: fragment.displayed_page,
                        bounds: PyBoundingBox {
                            x: fragment.x,
                            y: fragment.y,
                            width: fragment.width,
                            height: fragment.height,
                        },
                    }));
                }
                Ok::<_, rdocx::Error>(fragments)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        PyTuple::new(py, fragments)
    }

    fn layout_page(&self, page_index: usize, py: Python<'_>) -> PyResult<Option<PyLayoutPage>> {
        py.detach(|| {
            let layout = self.inner.layout_deterministic()?;
            Ok(layout
                .layout
                .pages
                .get(page_index)
                .map(|page| PyLayoutPage {
                    page_number: page.page_number,
                    displayed_page_number: page.displayed_page_number,
                    width: page.width,
                    height: page.height,
                }))
        })
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn rebuild_toc(&mut self, py: Python<'_>) -> PyResult<PyTocRebuildReport> {
        let (report, changed) = py
            .detach(|| {
                let before = self.inner.to_bytes()?;
                let report = self.inner.rebuild_toc()?;
                let changed = self.inner.to_bytes()? != before;
                Ok::<_, rdocx::Error>((report, changed))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if changed {
            self.revisions.bump();
        }
        Ok(PyTocRebuildReport {
            entry_count: report.entry_count,
            bookmark_count: report.bookmark_count,
            diagnostics: report.diagnostics,
        })
    }

    #[getter(revisions)]
    fn revision_snapshots<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let revisions = self
            .inner
            .story_revisions()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        PyTuple::new(
            py,
            revisions.iter().map(|revision| PyRevision {
                id: revision.id(),
                author: revision.author().to_owned(),
                timestamp: revision.timestamp().map(str::to_owned),
                kind: revision_kind_name(revision.kind()).to_owned(),
                story: Some(story_snapshot(revision.story())),
            }),
        )
    }

    fn accept_all(&mut self, py: Python<'_>) -> PyResult<usize> {
        self.counted_mutation(py, rdocx::Document::accept_all)
    }

    fn reject_all(&mut self, py: Python<'_>) -> PyResult<usize> {
        self.counted_mutation(py, rdocx::Document::reject_all)
    }

    fn accept_revisions_by_author(&mut self, py: Python<'_>, author: &str) -> PyResult<usize> {
        self.counted_mutation(py, |document| document.accept_revisions_by_author(author))
    }

    fn reject_revisions_by_author(&mut self, py: Python<'_>, author: &str) -> PyResult<usize> {
        self.counted_mutation(py, |document| document.reject_revisions_by_author(author))
    }

    #[pyo3(signature = (*, start, end))]
    fn accept_revisions_in_date_range(
        &mut self,
        py: Python<'_>,
        start: &str,
        end: &str,
    ) -> PyResult<usize> {
        self.counted_mutation(py, |document| {
            document.accept_revisions_in_date_range(start, end)
        })
    }

    #[pyo3(signature = (*, start, end))]
    fn reject_revisions_in_date_range(
        &mut self,
        py: Python<'_>,
        start: &str,
        end: &str,
    ) -> PyResult<usize> {
        self.counted_mutation(py, |document| {
            document.reject_revisions_in_date_range(start, end)
        })
    }

    fn accept_revision_id(&mut self, py: Python<'_>, id: i32) -> PyResult<usize> {
        self.counted_mutation(py, |document| document.accept_revision_id(id))
    }

    fn reject_revision_id(&mut self, py: Python<'_>, id: i32) -> PyResult<usize> {
        self.counted_mutation(py, |document| document.reject_revision_id(id))
    }

    /// Replace literal text only inside a checked paragraph, table or control item.
    #[pyo3(signature = (item, old, new, *, expect = None))]
    fn replace_text_at(
        &mut self,
        py: Python<'_>,
        item: &PyStoryItem,
        old: &str,
        new: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        let location = self.native_location(py, item)?;
        self.scoped_replacement(py, |document| {
            document.try_replace_text_at(&location, old, new, expect)
        })
    }

    #[pyo3(signature = (placeholder, replacement, *, expect = None))]
    fn try_replace_text(
        &mut self,
        py: Python<'_>,
        placeholder: &str,
        replacement: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        if expect.is_none() {
            return self.counted_mutation(py, |document| {
                document.try_replace_text(placeholder, replacement)
            });
        }
        let counts =
            self.expected_replacements(py, &[(placeholder, replacement, expect)], false)?;
        Ok(counts[0])
    }

    fn replace_all<'py>(
        &mut self,
        py: Python<'py>,
        pairs: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let mut owned = Vec::new();
        for (index, pair) in pairs.try_iter()?.enumerate() {
            let pair = pair?;
            let parsed = match pair.extract::<(String, String)>() {
                Ok((placeholder, replacement)) => (placeholder, replacement, None),
                Err(_) => pair
                    .extract::<(String, String, Option<usize>)>()
                    .map_err(|_| {
                        PyTypeError::new_err(format!(
                            "pair {index} must be an (old, new) or (old, new, expected) tuple, \
                         where expected is a nonnegative int or None"
                        ))
                    })?,
            };
            owned.push(parsed);
        }
        let pairs = owned
            .iter()
            .map(|(placeholder, replacement, expected)| {
                (placeholder.as_str(), replacement.as_str(), *expected)
            })
            .collect::<Vec<_>>();
        let counts = self.expected_replacements(py, &pairs, true)?;
        PyTuple::new(py, counts)
    }

    fn replace_all_regex(
        &mut self,
        py: Python<'_>,
        patterns: Vec<(String, String)>,
    ) -> PyResult<usize> {
        self.counted_mutation(py, |document| document.replace_all_regex(&patterns))
    }

    #[pyo3(signature = (
        *,
        now = None,
        file_name = None,
        file_path = None,
        merge_fields = None,
        included_text = None,
        merge_record_number = None,
        merge_sequence_number = None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn update_fields(
        &mut self,
        py: Python<'_>,
        now: Option<&Bound<'_, PyAny>>,
        file_name: Option<String>,
        file_path: Option<String>,
        merge_fields: Option<BTreeMap<String, String>>,
        included_text: Option<BTreeMap<String, String>>,
        merge_record_number: Option<u32>,
        merge_sequence_number: Option<u32>,
    ) -> PyResult<usize> {
        // Read field by field: the abi3 build has no datetime accessors, and
        // the wall-clock values are used as given.
        let now = now
            .map(|now| {
                Ok::<_, PyErr>(rdocx::FieldDateTime {
                    year: now.getattr("year")?.extract()?,
                    month: now.getattr("month")?.extract()?,
                    day: now.getattr("day")?.extract()?,
                    hour: now.getattr("hour")?.extract()?,
                    minute: now.getattr("minute")?.extract()?,
                    second: now.getattr("second")?.extract()?,
                })
            })
            .transpose()?;
        let context = rdocx::FieldEvaluationContext {
            now,
            file_name,
            file_path,
            merge_fields: merge_fields.unwrap_or_default(),
            included_text: included_text.unwrap_or_default(),
            merge_record_number,
            merge_sequence_number,
        };
        self.counted_mutation(py, |document| document.update_fields(&context))
    }

    fn update_page_fields(&mut self, py: Python<'_>) -> PyResult<usize> {
        let updated = py
            .detach(|| self.inner.update_page_fields())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if updated != 0 {
            self.revisions.bump();
        }
        Ok(updated)
    }

    fn update_layout_backed_fields(
        &mut self,
        py: Python<'_>,
    ) -> PyResult<PyLayoutBackedFieldUpdateReport> {
        let report = py
            .detach(|| self.inner.update_layout_backed_fields())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if report.updated_count() != 0 {
            self.revisions.bump();
        }
        Ok(PyLayoutBackedFieldUpdateReport {
            page_fields: report.page_fields,
            num_pages_fields: report.num_pages_fields,
            page_reference_fields: report.page_reference_fields,
            section_fields: report.section_fields,
            section_pages_fields: report.section_pages_fields,
            diagnostics: report.diagnostics,
        })
    }

    #[pyo3(signature = (index, max_level = 3))]
    fn insert_toc(&mut self, py: Python<'_>, index: usize, max_level: u32) -> PyResult<()> {
        if index > self.inner.content_count() {
            return Err(PyIndexError::new_err("content index out of range"));
        }
        if !(1..=9).contains(&max_level) {
            return Err(PyValueError::new_err("max_level must be between 1 and 9"));
        }
        self.inner
            .insert_toc(index, max_level)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    #[getter]
    fn paragraphs(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyParagraphCollection>> {
        Py::new(py, PyParagraphCollection::new(slf))
    }

    #[getter]
    fn tables(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyTableCollection>> {
        Py::new(py, PyTableCollection::new(slf))
    }

    fn add_paragraph(slf: Py<Self>, py: Python<'_>, text: &str) -> PyResult<Py<PyParagraph>> {
        let (index, path) = {
            let mut document = slf.borrow_mut(py);
            let index = document.inner.paragraph_count();
            document.inner.add_paragraph(text);
            document.revisions.bump();
            let path = document
                .revisions
                .capture(smallvec![PathSeg::Body(0), PathSeg::Para(index)]);
            (index, path)
        };
        debug_assert!(matches!(path.segs.last(), Some(PathSeg::Para(i)) if *i == index));
        Py::new(py, PyParagraph::new(slf, path))
    }

    #[pyo3(signature = (rows, cols))]
    fn add_table(slf: Py<Self>, py: Python<'_>, rows: usize, cols: usize) -> PyResult<Py<PyTable>> {
        let (index, path) = {
            let mut document = slf.borrow_mut(py);
            let index = document.inner.table_count();
            document.inner.add_table(rows, cols);
            document.revisions.bump();
            let path = document.revisions.capture(smallvec![PathSeg::Body(index)]);
            (index, path)
        };
        debug_assert!(matches!(path.segs.last(), Some(PathSeg::Body(i)) if *i == index));
        Py::new(py, PyTable::new(slf, path))
    }

    fn remove_content(&mut self, py: Python<'_>, index: usize) -> PyResult<bool> {
        let removed = py
            .detach(|| self.inner.try_remove_content(index))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if removed {
            self.revisions.bump();
        }
        Ok(removed)
    }

    fn find_content_index(
        slf: Py<Self>,
        py: Python<'_>,
        content: &Bound<'_, PyAny>,
    ) -> PyResult<usize> {
        if let Ok(text) = content.extract::<String>() {
            return slf
                .borrow(py)
                .inner
                .find_content_index(&text)
                .ok_or_else(|| PyValueError::new_err("text was not found in body content"));
        }
        Self::direct_content_index(&slf, py, content, "content")
    }

    fn find_content_indices<'py>(
        &self,
        py: Python<'py>,
        text: &str,
    ) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.inner.find_content_indices(text))
    }

    fn insert_paragraph(
        slf: Py<Self>,
        py: Python<'_>,
        index: usize,
        text: &str,
    ) -> PyResult<Py<PyParagraph>> {
        let path = {
            let mut document = slf.borrow_mut(py);
            if index > document.inner.content_count() {
                return Err(PyIndexError::new_err("content index out of range"));
            }
            document.inner.insert_paragraph(index, text);
            let paragraph = document
                .inner
                .paragraph_index_of_content(index)
                .expect("an inserted body paragraph has a paragraph index");
            document.revisions.bump();
            document
                .revisions
                .capture(smallvec![PathSeg::Body(0), PathSeg::Para(paragraph)])
        };
        Py::new(py, PyParagraph::new(slf, path))
    }

    #[pyo3(signature = (index, rows, cols))]
    fn insert_table(
        slf: Py<Self>,
        py: Python<'_>,
        index: usize,
        rows: usize,
        cols: usize,
    ) -> PyResult<Py<PyTable>> {
        let path = {
            let mut document = slf.borrow_mut(py);
            if index > document.inner.content_count() {
                return Err(PyIndexError::new_err("content index out of range"));
            }
            document.inner.insert_table(index, rows, cols);
            // Table handles count every table in document order, including
            // those inside block content controls, so find the new one.
            let table = (0..document.inner.table_count())
                .find(|table| document.inner.content_index_of_table(*table) == Some(index))
                .expect("an inserted body table has a table index");
            document.revisions.bump();
            document.revisions.capture(smallvec![PathSeg::Body(table)])
        };
        Py::new(py, PyTable::new(slf, path))
    }

    fn pop_content(
        slf: Py<Self>,
        py: Python<'_>,
        index: &Bound<'_, PyAny>,
    ) -> PyResult<PyContentFragment> {
        let location = if let Ok(item) = index.cast::<PyStoryItem>() {
            slf.borrow(py).native_location(py, &item.borrow())?
        } else {
            let index = body_index_argument(index, "index must be an int or a StoryItem")?;
            let location = slf.borrow(py).body_location(py, index)?;
            if index == slf.borrow(py).inner.content_count() {
                return Err(PyIndexError::new_err("content index out of range"));
            }
            location
        };
        let fragment = slf
            .borrow_mut(py)
            .inner
            .remove_content_at(&location)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        slf.borrow_mut(py).revisions.bump();
        Ok(PyContentFragment { inner: fragment })
    }

    fn insert_content(
        slf: Py<Self>,
        py: Python<'_>,
        destination: &Bound<'_, PyAny>,
        fragment: PyRef<'_, PyContentFragment>,
    ) -> PyResult<()> {
        // A story item names the boundary before it, and a story its end.
        let location = if let Ok(item) = destination.cast::<PyStoryItem>() {
            slf.borrow(py).native_location(py, &item.borrow())?
        } else if let Ok(story) = destination.cast::<PyStory>() {
            rdocx::ContentLocation::end(slf.borrow(py).native_story(py, &story.borrow())?)
        } else {
            let destination = body_index_argument(
                destination,
                "destination must be an int, a StoryItem or a Story",
            )?;
            slf.borrow(py).body_location(py, destination)?
        };
        let fragment = fragment.inner.clone();
        slf.borrow_mut(py)
            .inner
            .insert_content(&location, fragment)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        slf.borrow_mut(py).revisions.bump();
        Ok(())
    }

    fn clone_content(
        slf: Py<Self>,
        py: Python<'_>,
        source: &Bound<'_, PyAny>,
        destination: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let source_index = Self::direct_content_index(&slf, py, source, "source")?;
        let destination = destination
            .extract::<usize>()
            .map_err(|_| PyTypeError::new_err("destination must be a direct body index integer"))?;
        let (source, destination) = {
            let document = slf.borrow(py);
            let mut locations = document.body_locations(py, &[source_index, destination])?;
            let destination = locations.pop().expect("two requested body locations");
            let source = locations.pop().expect("two requested body locations");
            (source, destination)
        };
        slf.borrow_mut(py)
            .inner
            .clone_content(&source, &destination)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        slf.borrow_mut(py).revisions.bump();
        Ok(())
    }

    fn move_content(
        slf: Py<Self>,
        py: Python<'_>,
        source: &Bound<'_, PyAny>,
        destination: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let source_index = Self::direct_content_index(&slf, py, source, "source")?;
        let destination = destination
            .extract::<usize>()
            .map_err(|_| PyTypeError::new_err("destination must be a direct body index integer"))?;
        let (source, destination) = {
            let document = slf.borrow(py);
            let mut locations = document.body_locations(py, &[source_index, destination])?;
            let destination = locations.pop().expect("two requested body locations");
            let source = locations.pop().expect("two requested body locations");
            (source, destination)
        };
        slf.borrow_mut(py)
            .inner
            .move_content(&source, &destination)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        slf.borrow_mut(py).revisions.bump();
        Ok(())
    }

    // ---- Settings, notes and page decoration ----

    /// The document settings, python-docx's `document.settings`.
    #[getter]
    fn settings(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySettings>> {
        Py::new(py, PySettings::new(slf))
    }

    /// Add a footnote whose reference ends a body paragraph, given as the
    /// paragraph or its last run, and return the note ID.
    fn add_footnote(
        slf: Py<Self>,
        py: Python<'_>,
        target: &Bound<'_, PyAny>,
        text: &str,
    ) -> PyResult<i32> {
        Self::add_note(slf, py, target, text, rdocx::NoteFamily::Footnote)
    }

    /// Add an endnote whose reference ends a body paragraph, given as the
    /// paragraph or its last run, and return the note ID.
    fn add_endnote(
        slf: Py<Self>,
        py: Python<'_>,
        target: &Bound<'_, PyAny>,
        text: &str,
    ) -> PyResult<i32> {
        Self::add_note(slf, py, target, text, rdocx::NoteFamily::Endnote)
    }

    /// Remove a footnote and every reference to it.
    fn remove_footnote(&mut self, py: Python<'_>, id: i32) -> PyResult<()> {
        py.detach(|| self.inner.remove_footnote(id))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    /// Remove an endnote and every reference to it.
    fn remove_endnote(&mut self, py: Python<'_>, id: i32) -> PyResult<()> {
        py.detach(|| self.inner.remove_endnote(id))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    /// Set how footnotes or endnotes are numbered and placed, for the whole
    /// document or, with `section`, for one section.
    #[pyo3(signature = (kind, *, number_format = "decimal", start = 1, restart = "continuous", placement = None, section = None))]
    #[allow(clippy::too_many_arguments)]
    fn set_note_numbering(
        &mut self,
        py: Python<'_>,
        kind: &str,
        number_format: &str,
        start: u32,
        restart: &str,
        placement: Option<&str>,
        section: Option<usize>,
    ) -> PyResult<()> {
        let family = crate::story::note_family(kind)?;
        let policy = crate::story::note_policy(family, number_format, start, restart, placement)?;
        match section {
            None => self.inner.set_note_policy(family, policy),
            Some(index) => self
                .inner
                .section_mut(index)
                .ok_or_else(|| PyIndexError::new_err("section index out of range"))?
                .set_note_policy(family, policy),
        }
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// The page colour as an `RGBColor`, or `None`.
    #[getter]
    fn page_color(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.inner
            .page_color()
            .map(|hex| {
                py.import("rdocx")?
                    .getattr("RGBColor")?
                    .call_method1("from_string", (hex,))
                    .map(Bound::unbind)
            })
            .transpose()
    }

    /// Set the page colour from an `RGBColor` or six hexadecimal digits, or
    /// remove it with `None`.
    #[setter]
    fn set_page_color(&mut self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let color = value
            .map(|value| crate::story::hex_color(value, false))
            .transpose()?;
        self.inner
            .set_page_color(color.as_deref())
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Put a diagonal text watermark in every header variant in use.
    fn set_text_watermark(&mut self, py: Python<'_>, text: &str) -> PyResult<()> {
        py.detach(|| self.inner.set_text_watermark(text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    /// Put a picture watermark of `width` by `height` in every header variant
    /// in use.
    fn set_image_watermark(
        &mut self,
        py: Python<'_>,
        data: &[u8],
        filename: &str,
        width: i64,
        height: i64,
    ) -> PyResult<()> {
        py.detach(|| {
            self.inner.set_image_watermark(
                data,
                filename,
                rdocx::Length::emu(width),
                rdocx::Length::emu(height),
            )
        })
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    /// Frame every page of one section with a border on all four sides, or
    /// remove the frame with `style=None`.
    #[pyo3(
        signature = (section, style = Some("single"), *, width = 6_350, color = None, space = 304_800, offset_from = "page"),
        text_signature = "($self, section, style='single', *, width=6350, color=None, space=304800, offset_from='page')"
    )]
    #[allow(clippy::too_many_arguments)]
    fn set_page_borders(
        &mut self,
        section: usize,
        style: Option<&str>,
        width: i64,
        color: Option<&Bound<'_, PyAny>>,
        space: i64,
        offset_from: &str,
    ) -> PyResult<()> {
        let borders = style
            .map(|style| -> PyResult<_> {
                let color = color
                    .map(|color| crate::story::hex_color(color, true))
                    .transpose()?;
                let edge = crate::story::page_border_edge(style, width, color, space)?;
                let offset_from = rdocx_oxml::document::ST_PageBorderOffset::from_str(offset_from)
                    .ok_or_else(|| PyValueError::new_err("offset_from must be page or text"))?;
                Ok(rdocx_oxml::document::CT_PageBorders {
                    offset_from: Some(offset_from),
                    top: Some(edge.clone()),
                    left: Some(edge.clone()),
                    bottom: Some(edge.clone()),
                    right: Some(edge),
                    ..Default::default()
                })
            })
            .transpose()?;
        let mut target = self
            .inner
            .section_mut(section)
            .ok_or_else(|| PyIndexError::new_err("section index out of range"))?;
        match borders {
            Some(borders) => target.set_page_borders(borders),
            None => target.properties_mut().page_borders = None,
        }
        Ok(())
    }
}

impl PyDocument {
    fn add_note(
        slf: Py<Self>,
        py: Python<'_>,
        target: &Bound<'_, PyAny>,
        text: &str,
        family: rdocx::NoteFamily,
    ) -> PyResult<i32> {
        let paragraph = crate::story::note_reference_paragraph(py, &slf, target)?;
        let mut document = slf.borrow_mut(py);
        let location = document
            .inner
            .paragraph_story_location(paragraph)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
        let inner = &mut document.inner;
        let id = py
            .detach(|| match family {
                rdocx::NoteFamily::Footnote => inner.create_footnote(&location, text),
                rdocx::NoteFamily::Endnote => inner.create_endnote(&location, text),
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        // The reference run is appended after the paragraph's runs, so no
        // run or content index moves and live handles stay valid.
        Ok(id)
    }
}

fn parse_render_options(revision_view: &str) -> PyResult<rdocx::RenderOptions> {
    let revision_view = match revision_view {
        "accepted" => rdocx::RevisionView::Accepted,
        "tracked" => rdocx::RevisionView::Tracked,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown revision view {other:?}, expected accepted or tracked"
            )));
        }
    };
    Ok(rdocx::RenderOptions { revision_view })
}

fn parse_raster_format(
    format: &str,
    quality: u8,
    transparent: bool,
) -> rdocx::Result<rdocx::RasterFormat> {
    match format {
        "png" => Ok(rdocx::RasterFormat::Png {
            transparent_background: transparent,
        }),
        "jpg" | "jpeg" => Ok(rdocx::RasterFormat::Jpeg { quality }),
        "tif" | "tiff" => Ok(rdocx::RasterFormat::Tiff),
        other => Err(rdocx::Error::Other(format!(
            "unknown raster format {other:?}, expected png, jpeg, or tiff"
        ))),
    }
}

fn parse_comparison_granularity(name: &str) -> rdocx::Result<rdocx::ComparisonGranularity> {
    match name {
        "run" => Ok(rdocx::ComparisonGranularity::Run),
        "word" => Ok(rdocx::ComparisonGranularity::Word),
        "character" => Ok(rdocx::ComparisonGranularity::Character),
        other => Err(rdocx::Error::Other(format!(
            "unknown comparison granularity {other:?}, expected run, word, or character"
        ))),
    }
}

/// Story names follow `Story.kind`, so `body` selects the main story.
fn parse_comparison_story(name: &str) -> rdocx::Result<rdocx::ComparisonStoryKind> {
    match name {
        "body" => Ok(rdocx::ComparisonStoryKind::Main),
        "header" => Ok(rdocx::ComparisonStoryKind::Header),
        "footer" => Ok(rdocx::ComparisonStoryKind::Footer),
        "comment" => Ok(rdocx::ComparisonStoryKind::Comment),
        "text_box" => Ok(rdocx::ComparisonStoryKind::TextBox),
        "footnote" => Ok(rdocx::ComparisonStoryKind::Footnote),
        "endnote" => Ok(rdocx::ComparisonStoryKind::Endnote),
        other => Err(rdocx::Error::Other(format!(
            "unknown comparison story {other:?}, expected body, header, footer, comment, \
             text_box, footnote, or endnote"
        ))),
    }
}
