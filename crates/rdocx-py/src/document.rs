use std::collections::{BTreeMap, HashMap};
use std::hash::{DefaultHasher, Hash, Hasher};
use std::path::PathBuf;

use oxml_py_support::{PathSeg, RevisionCounter, StaleElementError};
use pyo3::exceptions::{
    PyFileNotFoundError, PyIndexError, PyKeyError, PyNotADirectoryError, PyOverflowError,
    PyTypeError, PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBool, PyBytes, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple};
use smallvec::smallvec;

use rdocx_oxml::{ST_PageOrientation, ST_SectionType};

use crate::paragraph::{
    ParagraphLocation, PyParagraph, PyParagraphCollection, defined_style, style_id_of_type,
};
use crate::rdocx_to_pyerr;
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

#[pyclass(name = "Section", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PySection {
    pub ordinal: usize,
    pub is_final: bool,
    pub orientation: Option<String>,
    pub page_width: Option<i64>,
    pub page_height: Option<i64>,
    pub margin_top: Option<i64>,
    pub margin_right: Option<i64>,
    pub margin_bottom: Option<i64>,
    pub margin_left: Option<i64>,
    pub gutter: Option<i64>,
    pub column_count: Option<u32>,
    pub column_spacing: Option<i64>,
    pub page_number_start: Option<u32>,
    pub header_distance: Option<i64>,
    pub footer_distance: Option<i64>,
    pub different_first_page: Option<bool>,
    pub break_type: Option<String>,
}

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
        }
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
}

impl PyDocument {
    fn from_document(inner: rdocx::Document) -> Self {
        Self {
            inner,
            revisions: RevisionCounter::new(),
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
        let ParagraphLocation::Body(paragraph_index) = paragraph.validate(py)? else {
            return Err(PyValueError::new_err(
                "StoryRunPosition does not accept a table cell paragraph handle",
            ));
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
                crate::paragraph::ParagraphLocation::Cell { .. } => {
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

fn section_snapshot(section: rdocx::SectionRef<'_>) -> PySection {
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

    // ---- Document-level views, templates, assembly, properties, content
    // controls and validation. The views share their code with the CLI, so
    // `text()`, `to_markdown()`, `to_html()` and `validate()` give what
    // `rdocx text`, `rdocx convert --to md|html` and `rdocx validate` give.

    /// The plain text of every story, as `rdocx text` prints it.
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let export = py.detach(|| self.inner.text_with_stories());
        story_export_text(py, export)
    }

    /// The Markdown that `rdocx convert --to md` writes.
    fn to_markdown(&self, py: Python<'_>) -> PyResult<String> {
        let export = py.detach(|| self.inner.to_markdown_with_stories());
        story_export_text(py, export)
    }

    /// The HTML document that `rdocx convert --to html` writes.
    fn to_html(&self, py: Python<'_>) -> PyResult<String> {
        let export = py.detach(|| self.inner.to_html_with_stories());
        story_export_text(py, export)
    }

    fn to_odt<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let result = py
            .detach(|| self.inner.to_odt_bytes())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let diagnostics = result
            .diagnostics
            .iter()
            .map(|diagnostic| format!("{}: {}", diagnostic.path, diagnostic.message));
        warn_conversion_losses(py, "to_odt", diagnostics)?;
        Ok(PyBytes::new(py, &result.bytes))
    }

    fn to_rtf<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let result = py
            .detach(|| self.inner.to_rtf_bytes())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        // An export diagnostic names its location as the destination.
        let diagnostics =
            result
                .diagnostics
                .iter()
                .map(|diagnostic| match &diagnostic.destination {
                    Some(location) => format!("{location}: {}", diagnostic.message),
                    None => diagnostic.message.clone(),
                });
        warn_conversion_losses(py, "to_rtf", diagnostics)?;
        Ok(PyBytes::new(py, &result.bytes))
    }

    fn to_epub<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let result = py
            .detach(|| self.inner.to_epub_bytes())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let diagnostics = result
            .diagnostics
            .iter()
            .map(|diagnostic| format!("{}: {}", diagnostic.path, diagnostic.message));
        warn_conversion_losses(py, "to_epub", diagnostics)?;
        Ok(PyBytes::new(py, &result.bytes))
    }

    fn word_count(&self) -> usize {
        self.inner.word_count()
    }

    #[pyo3(signature = (*, include_spaces = true))]
    fn character_count(&self, include_spaces: bool) -> usize {
        self.inner.character_count(include_spaces)
    }

    fn page_count(&self, py: Python<'_>) -> PyResult<usize> {
        py.detach(|| {
            self.inner
                .layout_deterministic()
                .map(|layout| layout.layout.pages.len())
        })
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Validate the document as it would be saved now.
    fn validate(&mut self, py: Python<'_>) -> PyResult<PyValidationReport> {
        py.detach(|| {
            let bytes = self.inner.to_bytes()?;
            rdocx::Document::validate_bytes(&bytes)
        })
        .map(PyValidationReport::from)
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Validate a file as `rdocx validate` does, even one that does not open.
    #[staticmethod]
    fn validate_file(path: PathBuf, py: Python<'_>) -> PyResult<PyValidationReport> {
        py.detach(|| rdocx::Document::validate_file(path))
            .map(PyValidationReport::from)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn render_template(&mut self, py: Python<'_>, data: &Bound<'_, PyAny>) -> PyResult<usize> {
        if !is_mapping(data)? {
            return Err(PyTypeError::new_err(format!(
                "render_template data must be a dict of tag paths to values, got {}",
                type_name(data)
            )));
        }
        let data = template_value(data, "data", &mut Vec::new())?;
        let count = py
            .detach(|| self.inner.render_template(&data))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        // A loop or a condition can change the content without one scalar tag.
        self.revisions.bump();
        Ok(count)
    }

    #[pyo3(signature = (other, at = None, *, conflict = "reuse_equivalent"))]
    fn insert_document(
        slf: Py<Self>,
        py: Python<'_>,
        other: Py<PyDocument>,
        at: Option<&Bound<'_, PyAny>>,
        conflict: &str,
    ) -> PyResult<()> {
        let policy = fragment_conflict_policy(conflict)?;
        let fragment = {
            let other = other.borrow(py);
            if other.inner.content_count() == 0 {
                return Ok(());
            }
            let start = other.body_location(py, 0)?;
            let end = rdocx::ContentLocation::end(other.body_story(py)?);
            rdocx::DocumentFragment::from_range(&other.inner, &start, &end, false)
                .map_err(|error| rdocx_to_pyerr(py, error))?
        };
        Self::import_document_fragment(&slf, py, &fragment, at, policy)
    }

    #[pyo3(signature = (start, end = None))]
    fn copy_fragment(
        &self,
        py: Python<'_>,
        start: &Bound<'_, PyAny>,
        end: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyDocumentFragment> {
        let start = if let Ok(item) = start.cast::<PyStoryItem>() {
            self.native_location(py, &item.borrow())?
        } else {
            let index = start
                .extract::<isize>()
                .map_err(|_| PyTypeError::new_err("start must be an int or a StoryItem"))?;
            let index = crate::normalize_index(index, self.inner.content_count(), "content")?;
            self.body_location(py, index)?
        };
        let end = match end {
            None => rdocx::ContentLocation::end(start.story().clone()),
            Some(end) => self.content_destination(py, end, "end")?,
        };
        rdocx::DocumentFragment::from_range(&self.inner, &start, &end, false)
            .map(|inner| PyDocumentFragment { inner })
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[pyo3(signature = (fragment, at = None, *, conflict = "reuse_equivalent"))]
    fn import_fragment(
        slf: Py<Self>,
        py: Python<'_>,
        fragment: PyRef<'_, PyDocumentFragment>,
        at: Option<&Bound<'_, PyAny>>,
        conflict: &str,
    ) -> PyResult<()> {
        let policy = fragment_conflict_policy(conflict)?;
        Self::import_document_fragment(&slf, py, &fragment.inner, at, policy)
    }

    #[getter]
    fn custom_properties(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyCustomProperties>> {
        Py::new(py, PyCustomProperties { document: slf })
    }

    #[getter]
    fn app_properties(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyAppProperties>> {
        Py::new(py, PyAppProperties { document: slf })
    }

    #[getter]
    fn settings(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySettings>> {
        Py::new(py, PySettings { document: slf })
    }

    #[getter]
    fn content_controls<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.inner
                .content_controls()
                .into_iter()
                .map(PyContentControl::from),
        )
    }

    #[pyo3(signature = (value, *, tag = None, alias = None))]
    fn set_content_control_value(
        &mut self,
        py: Python<'_>,
        value: &str,
        tag: Option<&str>,
        alias: Option<&str>,
    ) -> PyResult<usize> {
        let (key, name, matches) = match (tag, alias) {
            (Some(tag), None) => ("tag", tag, self.inner.content_controls_by_tag(tag)),
            (None, Some(alias)) => ("alias", alias, self.inner.content_controls_by_alias(alias)),
            _ => {
                return Err(PyTypeError::new_err(
                    "set_content_control_value takes exactly one of tag= or alias=",
                ));
            }
        };
        if matches.is_empty() {
            let mut known = self
                .inner
                .content_controls()
                .iter()
                .filter_map(|control| match key {
                    "tag" => control.tag().map(str::to_owned),
                    _ => control.alias().map(str::to_owned),
                })
                .collect::<Vec<_>>();
            known.sort();
            known.dedup();
            return Err(PyKeyError::new_err(format!(
                "no content control with {key} {name:?} in the document body; the {key}s are {known:?}"
            )));
        }
        if let Some(control) = matches
            .iter()
            .find(|control| !content_control_takes_text(control.control_type()))
        {
            return Err(PyValueError::new_err(format!(
                "content control with {key} {name:?} is a {} control, which holds no text value; \
                 only rich_text, plain_text, combo_box, dropdown_list and date controls take one",
                content_control_type_name(control.control_type())
            )));
        }
        for control in &matches {
            control
                .check_value(value)
                .map_err(|error| PyValueError::new_err(error.to_string()))?;
        }
        let count = match key {
            "tag" => self.inner.set_content_control_value_by_tag(name, value),
            _ => self.inner.set_content_control_value_by_alias(name, value),
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(count)
    }

    // ---- End of the document-level block.

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
    fn sections<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let sections = self
            .inner
            .sections()
            .map(section_snapshot)
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
        &mut self,
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
        let mut section = self
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
        Ok(self
            .inner
            .sections()
            .nth(index)
            .map(section_snapshot)
            .expect("the updated section exists"))
    }

    fn insert_section(&mut self, py: Python<'_>, index: usize) -> PyResult<()> {
        if index > self.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        self.inner
            .insert_section(index)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.revisions.bump();
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

// ---- Document-level views, templates, assembly, properties, content
// controls and validation: the records and helpers of the methods above.

impl PyDocument {
    /// The boundary that an `at=` or `end=` argument names: before a body
    /// index or a story item, at the end of a story, or at the end of the
    /// body when it is `None`.
    fn content_destination(
        &self,
        py: Python<'_>,
        value: &Bound<'_, PyAny>,
        name: &str,
    ) -> PyResult<rdocx::ContentLocation> {
        if let Ok(item) = value.cast::<PyStoryItem>() {
            self.native_location(py, &item.borrow())
        } else if let Ok(story) = value.cast::<PyStory>() {
            Ok(rdocx::ContentLocation::end(
                self.native_story(py, &story.borrow())?,
            ))
        } else {
            let message = format!("{name} must be an int, a StoryItem or a Story");
            let index = value
                .extract::<isize>()
                .map_err(|_| PyTypeError::new_err(message))?;
            // A boundary runs from 0 to the item count, and a negative index
            // counts from the end, so -1 is the boundary before the last item.
            let count = self.inner.content_count() as isize;
            let boundary = if index < 0 { count + index } else { index };
            if !(0..=count).contains(&boundary) {
                return Err(PyIndexError::new_err(format!(
                    "{name} index {index} is out of range for {count} body items"
                )));
            }
            self.body_location(py, boundary as usize)
        }
    }

    fn import_document_fragment(
        slf: &Py<Self>,
        py: Python<'_>,
        fragment: &rdocx::DocumentFragment,
        at: Option<&Bound<'_, PyAny>>,
        policy: rdocx::FragmentConflictPolicy,
    ) -> PyResult<()> {
        let destination = {
            let document = slf.borrow(py);
            match at {
                Some(at) => document.content_destination(py, at, "at")?,
                None => rdocx::ContentLocation::end(document.body_story(py)?),
            }
        };
        let mut document = slf.borrow_mut(py);
        let inner = &mut document.inner;
        py.detach(|| inner.import_fragment(&destination, fragment, policy))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.revisions.bump();
        Ok(())
    }
}

/// Return the text of a whole-document view, warning when the stories other
/// than the body had to be left out.
fn story_export_text(py: Python<'_>, export: rdocx::StoryTextExport) -> PyResult<String> {
    if let Some(reason) = &export.omitted_stories {
        warn_conversion(py, &format!("other stories left out: {reason}"))?;
    }
    Ok(export.text)
}

/// Warn once about the content an export could not represent.
fn warn_conversion_losses(
    py: Python<'_>,
    method: &str,
    diagnostics: impl Iterator<Item = String>,
) -> PyResult<()> {
    const SHOWN: usize = 5;
    let diagnostics = diagnostics.collect::<Vec<_>>();
    if diagnostics.is_empty() {
        return Ok(());
    }
    let mut message = format!(
        "{method} could not represent {} item(s): {}",
        diagnostics.len(),
        diagnostics[..diagnostics.len().min(SHOWN)].join("; ")
    );
    if diagnostics.len() > SHOWN {
        message.push_str(&format!("; and {} more", diagnostics.len() - SHOWN));
    }
    warn_conversion(py, &message)
}

/// Emit an `rdocx.ConversionWarning`, or a `UserWarning` without the package.
fn warn_conversion(py: Python<'_>, message: &str) -> PyResult<()> {
    let category = match py
        .import("rdocx")
        .and_then(|module| module.getattr("ConversionWarning"))
    {
        Ok(category) => category,
        Err(_) => py.get_type::<pyo3::exceptions::PyUserWarning>().into_any(),
    };
    let message = std::ffi::CString::new(message.replace('\0', " "))
        .map_err(|error| PyValueError::new_err(error.to_string()))?;
    PyErr::warn(py, &category, &message, 1)
}

fn fragment_conflict_policy(name: &str) -> PyResult<rdocx::FragmentConflictPolicy> {
    match name {
        "reuse_equivalent" => Ok(rdocx::FragmentConflictPolicy::reuse_equivalent()),
        "rename" => Ok(rdocx::FragmentConflictPolicy::rename_all()),
        other => Err(PyValueError::new_err(format!(
            "unknown conflict policy {other:?}, expected \"reuse_equivalent\" (reuse identical \
             styles, lists and media, rename the others) or \"rename\" (rename every one)"
        ))),
    }
}

fn type_name(value: &Bound<'_, PyAny>) -> String {
    value
        .get_type()
        .name()
        .map(|name| name.to_string())
        .unwrap_or_else(|_| "object".to_owned())
}

fn is_mapping(value: &Bound<'_, PyAny>) -> PyResult<bool> {
    if value.is_instance_of::<PyDict>() {
        return Ok(true);
    }
    let mapping = value.py().import("collections.abc")?.getattr("Mapping")?;
    value.is_instance(&mapping)
}

/// The deepest nesting of dicts and lists template data may use.
const MAX_TEMPLATE_DATA_DEPTH: usize = 256;

/// Convert template data to JSON. `path` names the value in error messages,
/// and `open` holds the containers being converted, to refuse a cycle and
/// bound the depth.
fn template_value(
    value: &Bound<'_, PyAny>,
    path: &str,
    open: &mut Vec<usize>,
) -> PyResult<serde_json::Value> {
    use serde_json::Value;

    if value.is_none() {
        return Ok(Value::Null);
    }
    if let Ok(flag) = value.cast::<PyBool>() {
        return Ok(Value::Bool(flag.is_true()));
    }
    if value.is_instance_of::<PyInt>() {
        if let Ok(number) = value.extract::<i64>() {
            return Ok(Value::from(number));
        }
        if let Ok(number) = value.extract::<u64>() {
            return Ok(Value::from(number));
        }
        return Err(PyValueError::new_err(format!(
            "template data {path} is an integer outside the 64-bit range; pass it as a str"
        )));
    }
    if let Ok(number) = value.cast::<PyFloat>() {
        return serde_json::Number::from_f64(number.value())
            .map(Value::Number)
            .ok_or_else(|| {
                PyValueError::new_err(format!(
                    "template data {path} is {}, which has no text form; pass a str instead",
                    number.value()
                ))
            });
    }
    if let Ok(text) = value.cast::<PyString>() {
        return Ok(Value::String(text.to_cow()?.into_owned()));
    }
    let container =
        is_mapping(value)? || value.is_instance_of::<PyList>() || value.is_instance_of::<PyTuple>();
    if container {
        let identity = value.as_ptr() as usize;
        if open.contains(&identity) {
            return Err(PyValueError::new_err(format!(
                "template data {path} contains itself; template data must be a tree"
            )));
        }
        if open.len() == MAX_TEMPLATE_DATA_DEPTH {
            return Err(PyValueError::new_err(format!(
                "template data {path} nests deeper than {MAX_TEMPLATE_DATA_DEPTH} dicts and lists"
            )));
        }
        open.push(identity);
        let converted = template_container(value, path, open);
        open.pop();
        return converted;
    }
    let datetime = value.py().import("datetime")?;
    let hint = if value.is_instance(&datetime.getattr("date")?)?
        || value.is_instance(&datetime.getattr("time")?)?
    {
        "format it first, for example with value.strftime(\"%d %B %Y\") or value.isoformat()"
    } else {
        "convert it to a str, int, float, bool, None, dict or list first"
    };
    Err(PyTypeError::new_err(format!(
        "template data {path} is a {}, which a template cannot render; {hint}",
        type_name(value)
    )))
}

/// Convert one dict or list of template data.
fn template_container(
    value: &Bound<'_, PyAny>,
    path: &str,
    open: &mut Vec<usize>,
) -> PyResult<serde_json::Value> {
    use serde_json::Value;

    if is_mapping(value)? {
        let mut object = serde_json::Map::new();
        for item in value.call_method0("items")?.try_iter()? {
            let (key, item) = item?.extract::<(Bound<'_, PyAny>, Bound<'_, PyAny>)>()?;
            let Ok(key) = key.extract::<String>() else {
                return Err(PyTypeError::new_err(format!(
                    "template data {path} has a {} key; tag paths need str keys",
                    type_name(&key)
                )));
            };
            let value = template_value(&item, &format!("{path}[{key:?}]"), open)?;
            object.insert(key, value);
        }
        return Ok(Value::Object(object));
    }
    let mut array = Vec::new();
    for (index, item) in value.try_iter()?.enumerate() {
        array.push(template_value(&item?, &format!("{path}[{index}]"), open)?);
    }
    Ok(Value::Array(array))
}

/// The findings of `Document.validate()` and `rdocx validate`.
#[pyclass(name = "ValidationReport", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyValidationReport {
    errors: Vec<String>,
    warnings: Vec<String>,
}

impl From<rdocx::ValidationReport> for PyValidationReport {
    fn from(report: rdocx::ValidationReport) -> Self {
        Self {
            errors: report.errors,
            warnings: report.warnings,
        }
    }
}

#[pymethods]
impl PyValidationReport {
    #[getter]
    fn errors<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.errors)
    }

    #[getter]
    fn warnings<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.warnings)
    }

    #[getter]
    fn ok(&self) -> bool {
        self.errors.is_empty()
    }

    fn __repr__(&self) -> String {
        format!(
            "ValidationReport(errors={:?}, warnings={:?})",
            self.errors, self.warnings
        )
    }
}

/// A range of story items copied out of a document, with its styles,
/// lists, pictures and other dependencies, for `Document.import_fragment`.
#[pyclass(name = "DocumentFragment", frozen, skip_from_py_object)]
pub struct PyDocumentFragment {
    inner: rdocx::DocumentFragment,
}

/// One content control of the document body.
#[pyclass(name = "ContentControl", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyContentControl {
    #[pyo3(get)]
    tag: Option<String>,
    #[pyo3(get)]
    alias: Option<String>,
    #[pyo3(get)]
    id: Option<i32>,
    #[pyo3(get, name = "type")]
    control_type: &'static str,
    #[pyo3(get)]
    text: String,
}

impl From<rdocx::ContentControlRef<'_>> for PyContentControl {
    fn from(control: rdocx::ContentControlRef<'_>) -> Self {
        Self {
            tag: control.tag().map(str::to_owned),
            alias: control.alias().map(str::to_owned),
            id: control.id(),
            control_type: content_control_type_name(control.control_type()),
            text: control.text(),
        }
    }
}

#[pymethods]
impl PyContentControl {
    fn __repr__(&self) -> String {
        let optional = |value: &Option<String>| match value {
            Some(value) => format!("{value:?}"),
            None => "None".to_owned(),
        };
        format!(
            "ContentControl(tag={}, alias={}, type={:?}, text={:?})",
            optional(&self.tag),
            optional(&self.alias),
            self.control_type,
            self.text
        )
    }
}

/// A control without a type element is a rich text control.
fn content_control_type_name(
    control_type: Option<rdocx_oxml::content_control::SdtType>,
) -> &'static str {
    use rdocx_oxml::content_control::SdtType;
    match control_type {
        None | Some(SdtType::RichText) => "rich_text",
        Some(SdtType::PlainText) => "plain_text",
        Some(SdtType::Picture) => "picture",
        Some(SdtType::CheckBox) => "checkbox",
        Some(SdtType::ComboBox) => "combo_box",
        Some(SdtType::DropDownList) => "dropdown_list",
        Some(SdtType::Date) => "date",
        Some(SdtType::DocumentPartList) => "document_part_list",
        Some(SdtType::DocumentPartObject) => "document_part_object",
        Some(SdtType::Group) => "group",
        Some(SdtType::RepeatingSection) => "repeating_section",
        Some(SdtType::RepeatingSectionItem) => "repeating_section_item",
        Some(SdtType::Citation) => "citation",
        Some(SdtType::Equation) => "equation",
        Some(SdtType::Bibliography) => "bibliography",
    }
}

fn content_control_takes_text(control_type: Option<rdocx_oxml::content_control::SdtType>) -> bool {
    matches!(
        content_control_type_name(control_type),
        "rich_text" | "plain_text" | "combo_box" | "dropdown_list" | "date"
    )
}

/// The document settings (`word/settings.xml`) under python-docx's name.
#[pyclass(name = "Settings")]
pub struct PySettings {
    document: Py<PyDocument>,
}

#[pymethods]
impl PySettings {
    /// Whether Word records edits as tracked changes.
    #[getter]
    fn track_revisions(&self, py: Python<'_>) -> bool {
        self.document
            .borrow(py)
            .inner
            .track_revisions()
            .unwrap_or(false)
    }

    #[setter]
    fn set_track_revisions(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let Some(value) = value else {
            return Err(PyTypeError::new_err(
                "track_revisions takes True or False, not None",
            ));
        };
        self.document
            .borrow_mut(py)
            .inner
            .set_track_revisions(value)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }
}

/// The custom document properties (`docProps/custom.xml`) as a mapping of
/// names to typed values.
#[pyclass(name = "CustomProperties", mapping)]
pub struct PyCustomProperties {
    document: Py<PyDocument>,
}

/// The format id Word gives every user-defined custom property.
const CUSTOM_PROPERTY_FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

impl PyCustomProperties {
    fn names(&self, py: Python<'_>) -> Vec<String> {
        self.document
            .borrow(py)
            .inner
            .custom_properties()
            .iter()
            .filter_map(|property| property.name.clone())
            .collect()
    }
}

fn custom_value_to_python(
    py: Python<'_>,
    value: &rdocx::CustomPropertyValue,
) -> PyResult<Py<PyAny>> {
    use rdocx::CustomPropertyValue as Value;
    Ok(match value {
        Value::Lpstr(text) | Value::Lpwstr(text) => text.into_pyobject(py)?.into_any().unbind(),
        Value::I4(number) => number.into_pyobject(py)?.into_any().unbind(),
        Value::R8(number) => number.into_pyobject(py)?.into_any().unbind(),
        Value::Bool(flag) => flag.into_pyobject(py)?.to_owned().into_any().unbind(),
        Value::FileTime(stamp) => match w3cdtf_to_datetime(py, stamp)? {
            Some(datetime) => datetime,
            None => stamp.into_pyobject(py)?.into_any().unbind(),
        },
        Value::Empty => py.None(),
        Value::Raw(xml) => PyBytes::new(py, xml).into_any().unbind(),
    })
}

fn custom_value_from_python(
    name: &str,
    value: &Bound<'_, PyAny>,
) -> PyResult<rdocx::CustomPropertyValue> {
    use rdocx::CustomPropertyValue as Value;
    if let Ok(flag) = value.cast::<PyBool>() {
        return Ok(Value::Bool(flag.is_true()));
    }
    if value.is_instance_of::<PyInt>() {
        return value.extract::<i32>().map(Value::I4).map_err(|_| {
            PyOverflowError::new_err(format!(
                "custom property {name:?} stores integers as 32 bits (vt:i4); \
                 store a larger number as a float or a str"
            ))
        });
    }
    if let Ok(number) = value.cast::<PyFloat>() {
        let number = number.value();
        if !number.is_finite() {
            return Err(PyValueError::new_err(format!(
                "custom property {name:?} cannot store {number}; store it as a str"
            )));
        }
        return Ok(Value::R8(number));
    }
    if let Ok(text) = value.cast::<PyString>() {
        return Ok(Value::Lpwstr(text.to_cow()?.into_owned()));
    }
    let datetime = value.py().import("datetime")?;
    if value.is_instance(&datetime.getattr("datetime")?)? {
        return Ok(Value::FileTime(datetime_to_w3cdtf(value)?));
    }
    if value.is_none() {
        return Err(PyTypeError::new_err(format!(
            "custom property {name:?} cannot be None; remove it with del doc.custom_properties[{name:?}]"
        )));
    }
    let hint = if value.is_instance(&datetime.getattr("date")?)? {
        "; pass a datetime.datetime, such as datetime.datetime.combine(value, datetime.time())"
    } else {
        ""
    };
    Err(PyTypeError::new_err(format!(
        "custom property {name:?} takes a str, int, float, bool or datetime, not {}{hint}",
        type_name(value)
    )))
}

#[pymethods]
impl PyCustomProperties {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.names(py).len()
    }

    fn __contains__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> bool {
        key.extract::<String>().is_ok_and(|name| {
            self.document
                .borrow(py)
                .inner
                .custom_property(&name)
                .is_some()
        })
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        PyList::new(py, self.names(py))?
            .into_any()
            .try_iter()
            .map(Bound::into_any)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let missing = || PyKeyError::new_err(key.clone().unbind());
        let name = key.extract::<String>().map_err(|_| missing())?;
        let document = self.document.borrow(py);
        let property = document.inner.custom_property(&name).ok_or_else(missing)?;
        custom_value_to_python(py, &property.value)
    }

    #[pyo3(signature = (key, default = None))]
    fn get(
        &self,
        py: Python<'_>,
        key: &Bound<'_, PyAny>,
        default: Option<Py<PyAny>>,
    ) -> PyResult<Py<PyAny>> {
        let document = self.document.borrow(py);
        let property = key
            .extract::<String>()
            .ok()
            .and_then(|name| document.inner.custom_property(&name).cloned());
        match property {
            Some(property) => custom_value_to_python(py, &property.value),
            None => Ok(default.unwrap_or_else(|| py.None())),
        }
    }

    fn __setitem__(
        &self,
        py: Python<'_>,
        key: &Bound<'_, PyAny>,
        value: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let name = key.extract::<String>().map_err(|_| {
            PyTypeError::new_err(format!(
                "custom property names are str, not {}",
                type_name(key)
            ))
        })?;
        let name = name.as_str();
        if name.is_empty() {
            return Err(PyValueError::new_err(
                "a custom property needs a non-empty name",
            ));
        }
        let value = custom_value_from_python(name, value)?;
        let mut document = self.document.borrow_mut(py);
        let property = match document.inner.custom_property(name) {
            Some(existing) => rdocx::CustomProperty {
                value,
                ..existing.clone()
            },
            None => rdocx::CustomProperty {
                fmtid: CUSTOM_PROPERTY_FMTID.to_owned(),
                pid: document
                    .inner
                    .custom_properties()
                    .iter()
                    .map(|property| property.pid)
                    .max()
                    .unwrap_or(1)
                    .max(1)
                    + 1,
                name: Some(name.to_owned()),
                value,
            },
        };
        document
            .inner
            .set_custom_property(property)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn __delitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<()> {
        let missing = || PyKeyError::new_err(key.clone().unbind());
        let name = key.extract::<String>().map_err(|_| missing())?;
        let removed = self
            .document
            .borrow_mut(py)
            .inner
            .remove_custom_property(&name)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        match removed {
            Some(_) => Ok(()),
            None => Err(missing()),
        }
    }

    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let document = self.document.borrow(py);
        let mut values = Vec::new();
        for property in document.inner.custom_properties() {
            if property.name.is_some() {
                values.push(custom_value_to_python(py, &property.value)?);
            }
        }
        PyList::new(py, values)
    }

    fn keys<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        PyList::new(py, self.names(py))
    }

    fn items<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let document = self.document.borrow(py);
        let mut items = Vec::new();
        for property in document.inner.custom_properties() {
            if let Some(name) = &property.name {
                items.push((name.clone(), custom_value_to_python(py, &property.value)?));
            }
        }
        PyList::new(py, items)
    }
}

/// One text field of the native application-properties model.
type AppTextField = fn(&mut rdocx::AppProperties) -> &mut Option<String>;

/// The application properties (`docProps/app.xml`). Word writes the counts
/// when it saves, so they are read-only here and may be stale; the
/// `word_count()`, `character_count()` and `page_count()` methods of the
/// document compute them.
#[pyclass(name = "AppProperties")]
pub struct PyAppProperties {
    document: Py<PyDocument>,
}

impl PyAppProperties {
    fn read<T>(&self, py: Python<'_>, field: fn(&rdocx::AppProperties) -> Option<T>) -> Option<T> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(field)
    }

    fn text(&self, py: Python<'_>, field: AppTextField) -> String {
        let document = self.document.borrow(py);
        let mut properties = document
            .inner
            .application_properties()
            .cloned()
            .unwrap_or_default();
        field(&mut properties).clone().unwrap_or_default()
    }

    fn set_text(&self, py: Python<'_>, field: AppTextField, value: Option<String>) -> PyResult<()> {
        let mut document = self.document.borrow_mut(py);
        let mut properties = document
            .inner
            .application_properties()
            .cloned()
            .unwrap_or_default();
        *field(&mut properties) = value.filter(|value| !value.is_empty());
        document
            .inner
            .set_application_properties(properties)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyAppProperties {
    #[getter]
    fn company(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.company)
    }

    #[setter]
    fn set_company(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.company, value)
    }

    #[getter]
    fn manager(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.manager)
    }

    #[setter]
    fn set_manager(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.manager, value)
    }

    #[getter]
    fn template(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.template)
    }

    #[setter]
    fn set_template(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.template, value)
    }

    #[getter]
    fn application(&self, py: Python<'_>) -> String {
        self.text(py, |properties| &mut properties.application)
    }

    #[setter]
    fn set_application(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |properties| &mut properties.application, value)
    }

    #[getter]
    fn pages(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |properties| properties.pages)
    }

    #[getter]
    fn words(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |properties| properties.words)
    }

    #[getter]
    fn characters(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |properties| properties.characters)
    }

    #[getter]
    fn characters_with_spaces(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |properties| properties.characters_with_spaces)
    }

    #[getter]
    fn lines(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |properties| properties.lines)
    }

    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |properties| properties.paragraphs)
    }
}
