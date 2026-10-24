#[cfg(test)]
use std::cell::Cell;
use std::collections::{BTreeMap, HashMap};
use std::hash::{DefaultHasher, Hash, Hasher};
use std::path::PathBuf;

#[cfg(test)]
thread_local! {
    static FX260_WORD_VISITS: Cell<usize> = const { Cell::new(0) };
}

pub(crate) fn note_fx260_word(count: usize) {
    #[cfg(test)]
    FX260_WORD_VISITS.with(|cell| cell.set(cell.get().saturating_add(count)));
    #[cfg(not(test))]
    let _ = count;
}

#[cfg(test)]
fn fx260_word_visits() -> usize {
    FX260_WORD_VISITS.with(|cell| cell.get())
}

#[cfg(test)]
fn reset_fx260_word_visits() {
    FX260_WORD_VISITS.with(|cell| cell.set(0));
}

// Debug builds count the whole-body work that direct reads and body splices
// avoid, outside the guards that check them against it, so that a test of
// the extension module proves each call took its direct route.
#[cfg(debug_assertions)]
static WORD_INVENTORY_REBUILDS: std::sync::atomic::AtomicUsize =
    std::sync::atomic::AtomicUsize::new(0);
#[cfg(debug_assertions)]
static WORD_WHOLE_SOURCE_READS: std::sync::atomic::AtomicUsize =
    std::sync::atomic::AtomicUsize::new(0);
#[cfg(debug_assertions)]
static WORD_PIECE_READS: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
#[cfg(debug_assertions)]
thread_local! {
    static WORD_GUARDING: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

/// Count one owner inventory rebuild from the whole document.
fn note_word_inventory_rebuild() {
    #[cfg(debug_assertions)]
    if !WORD_GUARDING.with(std::cell::Cell::get) {
        WORD_INVENTORY_REBUILDS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    }
}

/// Count one body block or story part whose owners are read again.
fn note_word_piece_read() {
    #[cfg(debug_assertions)]
    if !WORD_GUARDING.with(std::cell::Cell::get) {
        WORD_PIECE_READS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    }
}

/// Count one element read from its whole story source.
pub(crate) fn note_word_whole_source_read() {
    #[cfg(debug_assertions)]
    if !WORD_GUARDING.with(std::cell::Cell::get) {
        WORD_WHOLE_SOURCE_READS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    }
}

/// Run a debug guard's own whole-body work without counting it.
#[cfg(debug_assertions)]
fn guarded<T>(work: impl FnOnce() -> T) -> T {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) {
            WORD_GUARDING.with(|cell| cell.set(self.0));
        }
    }
    let _restore = Restore(WORD_GUARDING.with(|cell| cell.replace(true)));
    work()
}

/// The whole-body work counted so far: owner inventory rebuilds and
/// whole-source element reads. Debug builds only.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_word_whole_body_work")]
pub(crate) fn word_whole_body_work() -> (usize, usize) {
    (
        WORD_INVENTORY_REBUILDS.load(std::sync::atomic::Ordering::Relaxed),
        WORD_WHOLE_SOURCE_READS.load(std::sync::atomic::Ordering::Relaxed),
    )
}

/// The body blocks and story parts whose owners were read again, rather
/// than reused from the inventory by block and part. Debug builds only.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_word_piece_work")]
pub(crate) fn word_piece_work() -> usize {
    WORD_PIECE_READS.load(std::sync::atomic::Ordering::Relaxed)
}

/// The main part work staging and comparing saves did on this thread since
/// the previous call, in bytes of the main part parsed again, scanned for
/// nested namespace owners, canonicalized whole as a retained source and
/// written as a ZIP, then in body children serialized and read alone as kept
/// blocks. Debug builds only.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_staging_work")]
pub(crate) fn staging_work() -> [usize; 6] {
    rdocx::take_staging_work()
}

/// The work reopening documents did since the previous call, as packages
/// written as a ZIP, main document parts parsed and bytes of identifier
/// parts scanned. Debug builds only.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_reopen_work")]
pub(crate) fn reopen_work() -> [usize; 3] {
    rdocx::take_reopen_work()
}

/// The main-part body children that story lookups scanned again, rather
/// than reused from the scans the document keeps. Debug builds only.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_main_body_block_scans")]
pub(crate) fn main_body_block_scans() -> usize {
    rdocx::Document::main_body_block_scans()
}

use oxml_py_support::{
    ContentPath, PathSeg, PreparedRevision, RevisionCounter, RevisionScope, RevisionState,
    StaleElementError, revision_to_pyerr,
};
use pyo3::exceptions::{PyIndexError, PyKeyError, PyOverflowError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{
    PyAny, PyBool, PyBytes, PyDict, PyFloat, PyInt, PyList, PyMapping, PyString, PyTuple,
};
use smallvec::smallvec;

use rdocx_oxml::{ST_PageOrientation, ST_SectionType};

use crate::paragraph::{
    ParagraphLocation, PyParagraph, PyParagraphCollection, defined_style, style_id_of_type,
};
use crate::rdocx_to_pyerr;
use crate::table::{PyTable, PyTableCollection};

const CHART_TYPES: [(i64, rdocx::ChartKind); 7] = [
    (1, rdocx::ChartKind::Area),
    (51, rdocx::ChartKind::Bar),
    (-4120, rdocx::ChartKind::Doughnut),
    (65, rdocx::ChartKind::Line),
    (5, rdocx::ChartKind::Pie),
    (81, rdocx::ChartKind::Radar),
    (74, rdocx::ChartKind::Scatter),
];

fn checked_chart_inputs(
    chart_type: &Bound<'_, PyAny>,
    data: &Bound<'_, PyAny>,
) -> PyResult<(rdocx::ChartKind, rdocx::ChartData)> {
    if chart_type.is_instance_of::<pyo3::types::PyBool>() {
        return Err(PyTypeError::new_err(
            "chart type must be an XL_CHART_TYPE integer",
        ));
    }
    let value = chart_type.extract::<i64>()?;
    let kind = CHART_TYPES
        .iter()
        .find_map(|(v, k)| (*v == value).then_some(*k))
        .ok_or_else(|| PyValueError::new_err(format!("unsupported XL_CHART_TYPE value {value}")))?;
    let (categories, series, number_format) =
        data.call_method0("_data")?
            .extract::<(Vec<String>, Vec<(String, Vec<f64>)>, Option<String>)>()?;
    Ok((
        kind,
        rdocx::ChartData {
            categories,
            series,
            number_format,
            ..Default::default()
        },
    ))
}

#[pyclass(name = "PictureOptions", frozen, skip_from_py_object)]
#[derive(Clone)]
pub struct PyPictureOptions {
    pub(crate) inner: rdocx::PictureOptions,
}

#[pymethods]
impl PyPictureOptions {
    #[new]
    #[pyo3(signature = (*, name=None, description=None, title=None, decorative=false, crop=None, placement="inline", horizontal_relative_from="page", vertical_relative_from="page", horizontal_offset=None, vertical_offset=None, horizontal_alignment=None, vertical_alignment=None, distances=None, relative_height=0))]
    #[allow(clippy::too_many_arguments)] // Explicit operands match the native picture vocabulary.
    fn new(
        name: Option<String>,
        description: Option<String>,
        title: Option<String>,
        decorative: bool,
        crop: Option<&Bound<'_, PyAny>>,
        placement: &str,
        horizontal_relative_from: &str,
        vertical_relative_from: &str,
        horizontal_offset: Option<&Bound<'_, PyAny>>,
        vertical_offset: Option<&Bound<'_, PyAny>>,
        horizontal_alignment: Option<&str>,
        vertical_alignment: Option<&str>,
        distances: Option<&Bound<'_, PyAny>>,
        relative_height: u32,
    ) -> PyResult<Self> {
        let horizontal =
            checked_section_length(horizontal_offset, "picture horizontal offset")?.unwrap_or(0);
        let vertical =
            checked_section_length(vertical_offset, "picture vertical offset")?.unwrap_or(0);
        let crop = crop
            .map(|crop| {
                let values = crop.cast::<PyTuple>().map_err(|_| {
                    PyTypeError::new_err("picture crop must be a four-number tuple")
                })?;
                if values.len() != 4 {
                    return Err(PyValueError::new_err("picture crop must have four values"));
                }
                let mut out = [0; 4];
                for (index, value) in values.iter().enumerate() {
                    if value.is_instance_of::<PyBool>() {
                        return Err(PyTypeError::new_err(
                            "picture crop values must be numbers, not bool",
                        ));
                    }
                    let value = value.extract::<f64>()?;
                    if !value.is_finite() || !(0.0..1.0).contains(&value) {
                        return Err(PyValueError::new_err(
                            "picture crop values must be finite fractions from zero up to one",
                        ));
                    }
                    out[index] = (value * 100_000.0).trunc() as i32;
                }
                if out[0] + out[2] >= 100_000 || out[1] + out[3] >= 100_000 {
                    return Err(PyValueError::new_err(
                        "picture crop must leave a visible rectangle",
                    ));
                }
                Ok(rdocx::PictureCrop {
                    left: out[0],
                    top: out[1],
                    right: out[2],
                    bottom: out[3],
                })
            })
            .transpose()?;
        let mut distance = [0; 4];
        if let Some(distances) = distances {
            let values = distances.cast::<PyTuple>().map_err(|_| {
                PyTypeError::new_err("picture distances must be a four-length tuple")
            })?;
            if values.len() != 4 {
                return Err(PyValueError::new_err(
                    "picture distances must have four lengths",
                ));
            }
            for (index, value) in values.iter().enumerate() {
                distance[index] =
                    checked_section_length(Some(&value), "picture wrap distance")?.unwrap();
                if distance[index] < 0 {
                    return Err(PyValueError::new_err(
                        "picture wrap distances cannot be negative",
                    ));
                }
            }
        }
        let h = match horizontal_relative_from {
            "page" => rdocx::DrawingHorizontalRelativeFrom::Page,
            "margin" => rdocx::DrawingHorizontalRelativeFrom::Margin,
            "column" => rdocx::DrawingHorizontalRelativeFrom::Column,
            "character" => rdocx::DrawingHorizontalRelativeFrom::Character,
            "left_margin" => rdocx::DrawingHorizontalRelativeFrom::LeftMargin,
            "right_margin" => rdocx::DrawingHorizontalRelativeFrom::RightMargin,
            "inside_margin" => rdocx::DrawingHorizontalRelativeFrom::InsideMargin,
            "outside_margin" => rdocx::DrawingHorizontalRelativeFrom::OutsideMargin,
            _ => {
                return Err(PyValueError::new_err(
                    "unsupported horizontal picture frame",
                ));
            }
        };
        let v = match vertical_relative_from {
            "page" => rdocx::DrawingVerticalRelativeFrom::Page,
            "margin" => rdocx::DrawingVerticalRelativeFrom::Margin,
            "paragraph" => rdocx::DrawingVerticalRelativeFrom::Paragraph,
            "line" => rdocx::DrawingVerticalRelativeFrom::Line,
            "top_margin" => rdocx::DrawingVerticalRelativeFrom::TopMargin,
            "bottom_margin" => rdocx::DrawingVerticalRelativeFrom::BottomMargin,
            "inside_margin" => rdocx::DrawingVerticalRelativeFrom::InsideMargin,
            "outside_margin" => rdocx::DrawingVerticalRelativeFrom::OutsideMargin,
            _ => return Err(PyValueError::new_err("unsupported vertical picture frame")),
        };
        let ha = horizontal_alignment
            .map(|value| match value {
                "left" => Ok(rdocx::DrawingHorizontalAlignment::Left),
                "center" => Ok(rdocx::DrawingHorizontalAlignment::Center),
                "right" => Ok(rdocx::DrawingHorizontalAlignment::Right),
                "inside" => Ok(rdocx::DrawingHorizontalAlignment::Inside),
                "outside" => Ok(rdocx::DrawingHorizontalAlignment::Outside),
                _ => Err(PyValueError::new_err(
                    "unsupported horizontal picture alignment",
                )),
            })
            .transpose()?;
        let va = vertical_alignment
            .map(|value| match value {
                "top" => Ok(rdocx::DrawingVerticalAlignment::Top),
                "center" => Ok(rdocx::DrawingVerticalAlignment::Center),
                "bottom" => Ok(rdocx::DrawingVerticalAlignment::Bottom),
                "inside" => Ok(rdocx::DrawingVerticalAlignment::Inside),
                "outside" => Ok(rdocx::DrawingVerticalAlignment::Outside),
                _ => Err(PyValueError::new_err(
                    "unsupported vertical picture alignment",
                )),
            })
            .transpose()?;
        if (ha.is_some() && horizontal != 0) || (va.is_some() && vertical != 0) {
            return Err(PyValueError::new_err(
                "picture alignment and offset conflict",
            ));
        }
        let anchor = if placement == "inline" {
            if horizontal != 0
                || vertical != 0
                || ha.is_some()
                || va.is_some()
                || distance != [0; 4]
                || relative_height != 0
                || horizontal_relative_from != "page"
                || vertical_relative_from != "page"
            {
                return Err(PyValueError::new_err(
                    "inline pictures cannot have anchor placement operands",
                ));
            }
            None
        } else {
            let wrap = match placement {
                "square" => rdocx::DrawingWrap::Square,
                "tight" => rdocx::DrawingWrap::Tight,
                "top_and_bottom" => rdocx::DrawingWrap::TopAndBottom,
                "behind" | "in_front" => rdocx::DrawingWrap::None,
                _ => return Err(PyValueError::new_err("unsupported picture placement")),
            };
            Some(rdocx::PictureAnchor {
                horizontal_relative_from: h,
                horizontal_offset: rdocx::Length::emu(horizontal),
                horizontal_alignment: ha,
                vertical_relative_from: v,
                vertical_offset: rdocx::Length::emu(vertical),
                vertical_alignment: va,
                wrap,
                distance_top: rdocx::Length::emu(distance[0]),
                distance_right: rdocx::Length::emu(distance[1]),
                distance_bottom: rdocx::Length::emu(distance[2]),
                distance_left: rdocx::Length::emu(distance[3]),
                relative_height,
                behind_text: placement == "behind",
            })
        };
        Ok(Self {
            inner: rdocx::PictureOptions {
                width: rdocx::Length::emu(1),
                height: rdocx::Length::emu(1),
                crop,
                anchor,
                name,
                description,
                title,
                decorative,
            },
        })
    }
}

fn binary_stream_bytes(value: &Bound<'_, PyAny>, seek: bool) -> PyResult<Vec<u8>> {
    if seek {
        value.call_method1("seek", (0,)).map_err(|error| {
            if error.is_instance_of::<pyo3::exceptions::PyAttributeError>(value.py()) {
                PyTypeError::new_err(
                    "Document requires a seekable binary stream; use Document.from_bytes(stream.read())",
                )
            } else {
                error
            }
        })?;
    }
    let bytes = value.call_method0("read").map_err(|error| {
        if error.is_instance_of::<pyo3::exceptions::PyAttributeError>(value.py()) {
            PyTypeError::new_err(
                "use a binary stream, filesystem path or Document.from_bytes(bytes)",
            )
        } else {
            error
        }
    })?;
    bytes
        .cast::<PyBytes>()
        .map(|bytes| bytes.as_bytes().to_vec())
        .map_err(|_| {
            PyTypeError::new_err(
                "stream.read() must return binary bytes; use Document.from_bytes(bytes)",
            )
        })
}

fn open_document_input(py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<rdocx::Document> {
    if let Ok(path) = value.extract::<PathBuf>() {
        return rdocx::Document::open(path).map_err(|error| rdocx_to_pyerr(py, error));
    }
    rdocx::Document::from_bytes(&binary_stream_bytes(value, true)?)
        .map_err(|error| rdocx_to_pyerr(py, error))
}

fn checked_heading_level(value: &Bound<'_, PyAny>) -> PyResult<i64> {
    if value.is_instance_of::<PyBool>() {
        return Err(PyTypeError::new_err(
            "heading level must be an integer from 0 through 9, not bool",
        ));
    }
    value.extract()
}

pub(crate) fn picture_input(
    image: &Bound<'_, PyAny>,
    filename: Option<&str>,
) -> PyResult<(Vec<u8>, String)> {
    if let Ok(bytes) = image.cast::<PyBytes>() {
        return Ok((
            bytes.as_bytes().to_vec(),
            filename.unwrap_or("picture").to_owned(),
        ));
    }
    let path = match image.extract::<PathBuf>() {
        Ok(path) => path,
        Err(_) => {
            return Ok((
                binary_stream_bytes(image, false)?,
                filename.unwrap_or("picture").to_owned(),
            ));
        }
    };
    let bytes = std::fs::read(&path)
        .map_err(|error| pyo3::exceptions::PyOSError::new_err(error.to_string()))?;
    Ok((
        bytes,
        filename.map(str::to_owned).unwrap_or_else(|| {
            path.file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .into_owned()
        }),
    ))
}

/// A frozen physical picture, bound to one document and captured revision.
#[pyclass(name = "Picture", frozen, skip_from_py_object)]
pub struct PyPicture {
    pub(crate) inner: rdocx::PictureOccurrence,
    pub(crate) binding: oxml_py_support::ContentPath,
}

#[pymethods]
impl PyPicture {
    #[getter]
    fn width(&self) -> i64 {
        self.inner.width
    }
    #[getter]
    fn height(&self) -> i64 {
        self.inner.height
    }
    #[getter]
    fn relationship_id(&self) -> &str {
        &self.inner.relationship_id
    }
    #[getter]
    fn media_part_name(&self) -> Option<&str> {
        self.inner.media_part_name.as_deref()
    }
    #[getter]
    fn external_url(&self) -> Option<&str> {
        self.inner.external_url.as_deref()
    }
    #[getter]
    fn content_type(&self) -> &str {
        &self.inner.content_type
    }
    #[getter]
    fn description(&self) -> Option<&str> {
        self.inner.description.as_deref()
    }
    #[getter]
    fn title(&self) -> Option<&str> {
        self.inner.title.as_deref()
    }
    #[getter]
    fn name(&self) -> Option<&str> {
        self.inner.name.as_deref()
    }
    #[getter]
    fn decorative(&self) -> bool {
        self.inner.decorative
    }
    #[getter]
    fn anchored(&self) -> bool {
        self.inner.anchored
    }
    #[getter]
    fn bytes<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.inner.bytes)
    }
    #[getter]
    fn story(&self) -> PyStory {
        let mut story = story_snapshot(self.inner.location.story());
        story.binding = Some(self.binding.clone());
        story
    }
    #[getter]
    fn location(&self) -> PyContentLocation {
        let mut story = story_snapshot(self.inner.location.story());
        story.binding = Some(self.binding.clone());
        PyContentLocation {
            story,
            kind: "paragraph".into(),
            index_path: self.inner.location.index_path().to_vec(),
            is_end: false,
        }
    }
    #[getter]
    fn occurrence(&self) -> usize {
        self.inner.ordinal
    }
}

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
    pub resolved: Option<bool>,
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
        resolved: Option<bool>,
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

#[pyclass(name = "FieldEvaluation", frozen, get_all, skip_from_py_object)]
pub struct PyFieldEvaluation {
    field_index: usize,
    instruction: String,
    cached_result: String,
    outcome: Py<PyDict>,
}

#[pyclass(name = "FieldsUpdateReport", frozen, get_all, skip_from_py_object)]
pub struct PyFieldsUpdateReport {
    ordinary_updated: usize,
    layout: Py<PyLayoutBackedFieldUpdateReport>,
    evaluations: Py<PyTuple>,
    diagnostics: Py<PyTuple>,
    total_updated: usize,
    saved_cache_scope: String,
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

#[pyclass(name = "Story", frozen, eq, skip_from_py_object)]
#[derive(Clone)]
pub struct PyStory {
    pub kind: String,
    pub part_name: String,
    pub owner_index: usize,
    binding: Option<oxml_py_support::ContentPath>,
}

#[pymethods]
impl PyStory {
    #[getter]
    fn kind(&self) -> &str {
        &self.kind
    }
    #[getter]
    fn part_name(&self) -> &str {
        &self.part_name
    }
    #[getter]
    fn owner_index(&self) -> usize {
        self.owner_index
    }
    #[new]
    #[pyo3(signature = (*, kind, part_name, owner_index))]
    fn new(kind: String, part_name: String, owner_index: usize) -> Self {
        Self {
            kind,
            part_name,
            owner_index,
            binding: None,
        }
    }
}

#[pyclass(name = "StoryItem", frozen, eq, skip_from_py_object)]
#[derive(Clone)]
pub struct PyStoryItem {
    story: PyStory,
    kind: String,
    index_path: Vec<usize>,
    direct_body_index: Option<usize>,
    text: Option<String>,
    xml: Vec<u8>,
    revision: u64,
    binding: Option<oxml_py_support::ContentPath>,
    /// The document that listed this item. A caller-built item has none.
    document: Option<std::sync::Arc<Py<PyDocument>>>,
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
            binding: None,
            document: None,
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

    /// A paragraph handle for this paragraph item, or `None` for another kind
    /// or an item that was not read from a document.
    #[getter]
    fn paragraph(&self, py: Python<'_>) -> PyResult<Option<Py<PyParagraph>>> {
        let (Some(document), "paragraph", [index, ..]) = (
            &self.document,
            self.kind.as_str(),
            self.index_path.as_slice(),
        ) else {
            return Ok(None);
        };
        let document = document.clone_ref(py);
        let (location, path) = {
            let owner = document.borrow(py);
            let location = owner.native_location(py, self)?;
            let path = owner.revisions.capture(smallvec![PathSeg::Para(*index)]);
            (location, path)
        };
        Py::new(
            py,
            PyParagraph::for_story(document, path, location).bind_current(py, None)?,
        )
        .map(Some)
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
    owner: Option<(ContentPath, RevisionState)>,
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
            owner: None,
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

#[pyclass(name = "SectionHandle")]
pub struct PySectionHandle {
    document: Py<PyDocument>,
    index: usize,
    path: oxml_py_support::ContentPath,
    owner: RevisionState,
}

impl PySectionHandle {
    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
                "section",
                "Re-fetch it with document.section(index).",
            )
            .map_err(|error| crate::stale_to_pyerr(py, error))?;
        if !document.owners.iter().any(|owner| {
            owner.owner.same_state(&self.owner)
                && matches!(owner.location,WordOwnerLocation::Section(index) if index==self.index)
        }) {
            return Err(PyValueError::new_err(
                "section physical owner is no longer current",
            ));
        }
        if self.index >= document.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        Ok(())
    }

    fn snapshot(&self, py: Python<'_>) -> PyResult<PySection> {
        self.validate(py)?;
        Ok(section_snapshot(
            self.document
                .borrow(py)
                .inner
                .section(self.index)
                .expect("validated section"),
        ))
    }

    fn geometry(&mut self, py: Python<'_>, name: &str, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.validate(py)?;
        let value = checked_section_length(Some(&value), name)?
            .map(|emu| crate::length_object(py, rdocx::Length::emu(emu)))
            .transpose()?;
        let mut document = self.document.borrow_mut(py);
        document.update_section(
            py,
            self.index,
            None,
            if name == "page_width" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "page_height" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "margin_top" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "margin_right" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "margin_bottom" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "margin_left" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "gutter" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            None,
            None,
            None,
            if name == "header_distance" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            if name == "footer_distance" {
                value.as_ref().map(|value| value.bind(py).clone())
            } else {
                None
            },
            None,
            None,
        )?;
        Ok(())
    }

    fn furniture(
        &self,
        py: Python<'_>,
        kind: rdocx::HeaderFooterKind,
        variant: rdocx::HdrFtrType,
    ) -> PyResult<PyHeaderFooter> {
        self.validate(py)?;
        Ok(PyHeaderFooter {
            document: self.document.clone_ref(py),
            index: self.index,
            kind,
            variant,
            path: self
                .document
                .borrow(py)
                .capture_owner(
                    &WordOwnerLocation::Section(self.index),
                    self.path.segs.clone(),
                    WordView::Content,
                    None,
                )?
                .0,
        })
    }
}

#[pymethods]
impl PySectionHandle {
    #[getter]
    fn element<'py>(slf: &Bound<'py, Self>) -> Bound<'py, Self> {
        slf.clone()
    }

    #[getter]
    fn xml(&self, py: Python<'_>) -> PyResult<String> {
        self.validate(py)?;
        let bytes = self
            .document
            .borrow(py)
            .inner
            .section_properties_xml(self.index)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        String::from_utf8(bytes).map_err(|error| PyValueError::new_err(error.to_string()))
    }

    fn replace_xml(&self, py: Python<'_>, xml: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .replace_section_properties_xml(self.index, xml)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.publish_owner_candidate(
            py,
            candidate,
            "Section.replace_xml",
            WordPublication::ElementXml(WordOwnerLocation::Section(self.index)),
        )
    }

    #[getter]
    fn page_width(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.page_width)
    }
    #[setter]
    fn set_page_width(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "page_width", value)
    }
    #[getter]
    fn page_height(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.page_height)
    }
    #[setter]
    fn set_page_height(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "page_height", value)
    }
    #[getter]
    fn margin_top(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_top)
    }
    #[setter]
    fn set_margin_top(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_top", value)
    }
    #[getter]
    fn margin_right(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_right)
    }
    #[setter]
    fn set_margin_right(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_right", value)
    }
    #[getter]
    fn margin_bottom(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_bottom)
    }
    #[setter]
    fn set_margin_bottom(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_bottom", value)
    }
    #[getter]
    fn margin_left(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_left)
    }
    #[setter]
    fn set_margin_left(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_left", value)
    }
    #[getter]
    fn top_margin(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_top)
    }
    #[setter]
    fn set_top_margin(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_top", value)
    }
    #[getter]
    fn right_margin(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_right)
    }
    #[setter]
    fn set_right_margin(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_right", value)
    }
    #[getter]
    fn bottom_margin(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_bottom)
    }
    #[setter]
    fn set_bottom_margin(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_bottom", value)
    }
    #[getter]
    fn left_margin(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.margin_left)
    }
    #[setter]
    fn set_left_margin(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "margin_left", value)
    }
    #[getter]
    fn gutter(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.gutter)
    }
    #[setter]
    fn set_gutter(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "gutter", value)
    }
    #[getter]
    fn header_distance(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.header_distance)
    }
    #[setter]
    fn set_header_distance(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "header_distance", value)
    }
    #[getter]
    fn footer_distance(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        Ok(self.snapshot(py)?.footer_distance)
    }
    #[setter]
    fn set_footer_distance(&mut self, py: Python<'_>, value: Bound<'_, PyAny>) -> PyResult<()> {
        self.geometry(py, "footer_distance", value)
    }
    #[getter]
    fn orientation(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.orientation)
    }
    #[setter]
    fn set_orientation(&mut self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.validate(py)?;
        self.document.borrow_mut(py).update_section(
            py,
            self.index,
            Some(value),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        )?;
        Ok(())
    }
    #[getter]
    fn different_first_page_header_footer(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.snapshot(py)?.different_first_page.unwrap_or(false))
    }
    #[setter]
    fn set_different_first_page_header_footer(
        &mut self,
        py: Python<'_>,
        value: bool,
    ) -> PyResult<()> {
        self.validate(py)?;
        self.document.borrow_mut(py).update_section(
            py,
            self.index,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            Some(value),
            None,
        )?;
        Ok(())
    }
    #[pyo3(signature = (family, *, format="decimal", start=None, restart="continuous", placement=None))]
    fn set_note_policy(
        &mut self,
        py: Python<'_>,
        family: &str,
        format: &str,
        start: Option<&Bound<'_, PyAny>>,
        restart: &str,
        placement: Option<&str>,
    ) -> PyResult<()> {
        self.validate(py)?;
        let (family, policy) = checked_note_policy(family, format, start, restart, placement)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .section_mut(self.index)
            .ok_or_else(|| PyIndexError::new_err("section index out of range"))?
            .set_note_policy(family, policy)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        candidate
            .to_bytes()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "SectionHandle.set_note_policy",
            crate::document::WordPublication::Conservative,
        )?;
        let (next, next_owner) = document.capture_prepared_owner(
            &prepared,
            &WordOwnerLocation::Section(self.index),
            smallvec![],
            WordView::Owner,
            None,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        self.owner = next_owner;
        Ok(())
    }

    #[pyo3(signature = (edge, *, style="single", width=8, space=0, color="000000"))]
    fn set_page_border(
        &mut self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        width: u32,
        space: u32,
        color: &str,
    ) -> PyResult<()> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .section_mut(self.index)
            .ok_or_else(|| PyIndexError::new_err("section index out of range"))?
            .set_page_border(edge, style, width, space, color)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        candidate
            .to_bytes()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "SectionHandle.set_page_border",
            crate::document::WordPublication::Conservative,
        )?;
        let (next, next_owner) = document.capture_prepared_owner(
            &prepared,
            &WordOwnerLocation::Section(self.index),
            smallvec![],
            WordView::Owner,
            None,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        self.owner = next_owner;
        Ok(())
    }

    /// Author actual DrawingML text, with styled real text in both representations.
    #[allow(clippy::too_many_arguments)] // Keep explicit checked public watermark operands.
    #[pyo3(signature = (text, *, width, height, font_size, variant="default", rotation=315.0, color="D9D9D9", font_family="Calibri", opacity=0.5))]
    fn set_text_watermark(
        &mut self,
        py: Python<'_>,
        text: &str,
        width: &Bound<'_, PyAny>,
        height: &Bound<'_, PyAny>,
        font_size: &Bound<'_, PyAny>,
        variant: &str,
        rotation: f64,
        color: &str,
        font_family: &str,
        opacity: f64,
    ) -> PyResult<()> {
        self.validate(py)?;
        let variant = match variant {
            "default" => rdocx::HdrFtrType::Default,
            "first" => rdocx::HdrFtrType::First,
            "even" => rdocx::HdrFtrType::Even,
            _ => {
                return Err(PyValueError::new_err(
                    "watermark variant must be default, first or even",
                ));
            }
        };
        let width = checked_section_length(Some(width), "watermark width")?.unwrap();
        let height = checked_section_length(Some(height), "watermark height")?.unwrap();
        let font_size = checked_section_length(Some(font_size), "watermark font size")?.unwrap();
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .set_drawingml_text_watermark_for(
                self.index,
                variant,
                text,
                rdocx::TextWatermarkOptions {
                    width: rdocx::Length::emu(width),
                    height: rdocx::Length::emu(height),
                    rotation_degrees: rotation,
                    color: color.to_owned(),
                    font_family: Some(font_family.to_owned()),
                    opacity,
                },
                rdocx::Length::emu(font_size),
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "SectionHandle.set_text_watermark",
            crate::document::WordPublication::Conservative,
        )?;
        let (next, next_owner) = document.capture_prepared_owner(
            &prepared,
            &WordOwnerLocation::Section(self.index),
            smallvec![],
            WordView::Owner,
            None,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        self.owner = next_owner;
        Ok(())
    }

    #[allow(clippy::too_many_arguments)] // Explicit checked watermark operands.
    #[pyo3(signature = (data, filename, *, width, height, variant="default", rotation=0.0, opacity=0.5))]
    fn set_image_watermark(
        &mut self,
        py: Python<'_>,
        data: &[u8],
        filename: &str,
        width: &Bound<'_, PyAny>,
        height: &Bound<'_, PyAny>,
        variant: &str,
        rotation: f64,
        opacity: f64,
    ) -> PyResult<()> {
        self.validate(py)?;
        let variant = match variant {
            "default" => rdocx::HdrFtrType::Default,
            "first" => rdocx::HdrFtrType::First,
            "even" => rdocx::HdrFtrType::Even,
            _ => {
                return Err(PyValueError::new_err(
                    "watermark variant must be default, first or even",
                ));
            }
        };
        let width = checked_section_length(Some(width), "watermark width")?.unwrap();
        let height = checked_section_length(Some(height), "watermark height")?.unwrap();
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .set_drawingml_image_watermark_for(
                self.index,
                variant,
                data,
                filename,
                rdocx::ImageWatermarkOptions {
                    width: rdocx::Length::emu(width),
                    height: rdocx::Length::emu(height),
                    rotation_degrees: rotation,
                    opacity,
                },
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "SectionHandle.set_image_watermark",
            crate::document::WordPublication::Conservative,
        )?;
        let (next, next_owner) = document.capture_prepared_owner(
            &prepared,
            &WordOwnerLocation::Section(self.index),
            smallvec![],
            WordView::Owner,
            None,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        self.owner = next_owner;
        Ok(())
    }

    #[pyo3(signature = (*, variant="default"))]
    fn remove_watermark(&mut self, py: Python<'_>, variant: &str) -> PyResult<()> {
        self.validate(py)?;
        let variant = match variant {
            "default" => rdocx::HdrFtrType::Default,
            "first" => rdocx::HdrFtrType::First,
            "even" => rdocx::HdrFtrType::Even,
            _ => {
                return Err(PyValueError::new_err(
                    "watermark variant must be default, first or even",
                ));
            }
        };
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .remove_drawingml_watermark_for(self.index, variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "SectionHandle.remove_watermark",
            crate::document::WordPublication::Conservative,
        )?;
        let (next, next_owner) = document.capture_prepared_owner(
            &prepared,
            &WordOwnerLocation::Section(self.index),
            smallvec![],
            WordView::Owner,
            None,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        self.owner = next_owner;
        Ok(())
    }

    #[getter]
    fn header(&self, py: Python<'_>) -> PyResult<PyHeaderFooter> {
        self.furniture(
            py,
            rdocx::HeaderFooterKind::Header,
            rdocx::HdrFtrType::Default,
        )
    }
    #[getter]
    fn first_page_header(&self, py: Python<'_>) -> PyResult<PyHeaderFooter> {
        self.furniture(
            py,
            rdocx::HeaderFooterKind::Header,
            rdocx::HdrFtrType::First,
        )
    }
    #[getter]
    fn even_page_header(&self, py: Python<'_>) -> PyResult<PyHeaderFooter> {
        self.furniture(py, rdocx::HeaderFooterKind::Header, rdocx::HdrFtrType::Even)
    }
    #[getter]
    fn footer(&self, py: Python<'_>) -> PyResult<PyHeaderFooter> {
        self.furniture(
            py,
            rdocx::HeaderFooterKind::Footer,
            rdocx::HdrFtrType::Default,
        )
    }
    #[getter]
    fn first_page_footer(&self, py: Python<'_>) -> PyResult<PyHeaderFooter> {
        self.furniture(
            py,
            rdocx::HeaderFooterKind::Footer,
            rdocx::HdrFtrType::First,
        )
    }
    #[getter]
    fn even_page_footer(&self, py: Python<'_>) -> PyResult<PyHeaderFooter> {
        self.furniture(py, rdocx::HeaderFooterKind::Footer, rdocx::HdrFtrType::Even)
    }
}

fn checked_note_policy(
    family: &str,
    format: &str,
    start: Option<&Bound<'_, PyAny>>,
    restart: &str,
    placement: Option<&str>,
) -> PyResult<(rdocx::NoteFamily, rdocx::NotePolicy)> {
    let family = match family {
        "footnote" => rdocx::NoteFamily::Footnote,
        "endnote" => rdocx::NoteFamily::Endnote,
        _ => {
            return Err(PyValueError::new_err(
                "note family must be footnote or endnote",
            ));
        }
    };
    let format = match format {
        "decimal" => rdocx::NoteNumberFormat::Decimal,
        "upperRoman" => rdocx::NoteNumberFormat::UpperRoman,
        "lowerRoman" => rdocx::NoteNumberFormat::LowerRoman,
        "upperLetter" => rdocx::NoteNumberFormat::UpperLetter,
        "lowerLetter" => rdocx::NoteNumberFormat::LowerLetter,
        _ => return Err(PyValueError::new_err("unsupported note number format")),
    };
    let start = match start {
        None => 1,
        Some(value) if value.is_instance_of::<PyBool>() => {
            return Err(PyTypeError::new_err(
                "note numbering start must be an integer, not bool",
            ));
        }
        Some(value) => value.extract::<u32>()?,
    };
    let restart = match restart {
        "continuous" => rdocx::NoteRestart::Continuous,
        "eachSect" => rdocx::NoteRestart::EachSection,
        "eachPage" => rdocx::NoteRestart::EachPage,
        _ => return Err(PyValueError::new_err("unsupported note restart")),
    };
    let placement = match placement {
        None if family == rdocx::NoteFamily::Footnote => rdocx::NotePlacement::PageBottom,
        None => rdocx::NotePlacement::DocumentEnd,
        Some("pageBottom") => rdocx::NotePlacement::PageBottom,
        Some("beneathText") => rdocx::NotePlacement::BeneathText,
        Some("sectEnd") => rdocx::NotePlacement::SectionEnd,
        Some("docEnd") => rdocx::NotePlacement::DocumentEnd,
        _ => return Err(PyValueError::new_err("unsupported note placement")),
    };
    Ok((
        family,
        rdocx::NotePolicy {
            format,
            start,
            restart,
            placement,
        },
    ))
}

fn related_table_fragment(rows: usize, cols: usize) -> PyResult<rdocx::ContentFragment> {
    use rdocx_oxml::table::{
        CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc,
    };
    if rows == 0 || cols == 0 {
        return Err(PyValueError::new_err("table dimensions must be positive"));
    }
    let columns = i32::try_from(cols)
        .map_err(|_| PyOverflowError::new_err("table column count is too large"))?;
    let width = 9360 / columns;
    if width == 0 {
        return Err(PyValueError::new_err(
            "table columns must be at least one twip wide",
        ));
    }
    let mut table = CT_Tbl::new();
    table.properties = Some(CT_TblPr {
        width: Some(CT_TblWidth::dxa(width * columns)),
        ..Default::default()
    });
    table.grid = Some(CT_TblGrid {
        columns: (0..cols)
            .map(|_| CT_TblGridCol {
                width: rdocx_oxml::units::Twips(width),
            })
            .collect(),
        ..Default::default()
    });
    for _ in 0..rows {
        let mut row = CT_Row::new();
        row.cells.extend((0..cols).map(|_| CT_Tc::new()));
        table.rows.push(row);
    }
    rdocx::ContentFragment::table(table).map_err(|error| PyValueError::new_err(error.to_string()))
}

#[pyclass(name = "HeaderFooter")]
pub struct PyHeaderFooter {
    document: Py<PyDocument>,
    index: usize,
    kind: rdocx::HeaderFooterKind,
    variant: rdocx::HdrFtrType,
    path: oxml_py_support::ContentPath,
}

impl PyHeaderFooter {
    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        self.path
            .validate_scopes(
                &self.document.borrow(py).revisions,
                "header/footer",
                "Re-fetch it from document.section(index).",
            )
            .map_err(|error| crate::stale_to_pyerr(py, error))?;
        Ok(())
    }
}

#[pymethods]
impl PyHeaderFooter {
    #[getter]
    fn is_linked_to_previous(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        Ok(self
            .document
            .borrow(py)
            .inner
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .is_none_or(|story| story.is_inherited()))
    }
    #[setter]
    fn set_is_linked_to_previous(&mut self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.validate(py)?;
        if self.is_linked_to_previous(py)? == value {
            return Ok(());
        }
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        if value {
            candidate.inherit_section_story(self.index, self.kind, self.variant)
        } else if candidate
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .is_none()
        {
            candidate
                .create_section_story(self.index, self.kind, self.variant)
                .map(|_| ())
        } else {
            candidate
                .unlink_section_story(self.index, self.kind, self.variant)
                .map(|_| ())
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "HeaderFooter.is_linked_to_previous",
            crate::document::WordPublication::Conservative,
        )?;
        let next = document
            .capture_prepared_owner(
                &prepared,
                &WordOwnerLocation::Section(self.index),
                smallvec![],
                WordView::Content,
                None,
            )?
            .0;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        Ok(())
    }
    #[getter]
    fn paragraphs<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        self.validate(py)?;
        let document = self.document.borrow(py);
        let mut result = Vec::new();
        if let Some(story) = document
            .inner
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            for item in document
                .inner
                .story_items(story.story())
                .map_err(|error| rdocx_to_pyerr(py, error))?
            {
                if item.kind() == rdocx::StoryItemKind::Paragraph {
                    let location = item.location().clone();
                    let path = document
                        .revisions
                        .capture(smallvec![PathSeg::Para(location.index_path()[0])]);
                    result.push(Py::new(
                        py,
                        PyParagraph::for_story(self.document.clone_ref(py), path, location)
                            .bind_current(py, Some(document.section_child_scope(self.index)?))?,
                    )?);
                }
            }
        }
        PyTuple::new(py, result)
    }
    #[pyo3(signature = (text = "", style = None))]
    fn add_paragraph(
        &mut self,
        py: Python<'_>,
        text: &str,
        style: Option<&str>,
    ) -> PyResult<Py<PyParagraph>> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let style = match style {
            Some(value) => Some(style_id_of_type(
                &candidate,
                value,
                rdocx::StyleType::Paragraph,
            )?),
            None => {
                let wanted = if self.kind == rdocx::HeaderFooterKind::Header {
                    "Header"
                } else {
                    "Footer"
                };
                candidate
                    .styles()
                    .into_iter()
                    .find(|style| {
                        style.style_type() == rdocx::StyleType::Paragraph
                            && (style.style_id() == wanted
                                || style
                                    .name()
                                    .is_some_and(|name| name.eq_ignore_ascii_case(wanted)))
                    })
                    .map(|style| style.style_id().to_owned())
            }
        };
        let story = match candidate
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            Some(story) => story.story().clone(),
            None => candidate
                .create_section_story(self.index, self.kind, self.variant)
                .map_err(|error| rdocx_to_pyerr(py, error))?,
        };
        let mut paragraph = rdocx_oxml::text::CT_P::new();
        if !text.is_empty() {
            paragraph.add_run(text);
        }
        if let Some(style) = style {
            paragraph.properties = Some(rdocx::CT_PPr {
                style_id: Some(style),
                ..Default::default()
            });
        }
        let index = candidate
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .len();
        candidate
            .insert_content(
                &rdocx::ContentLocation::end(story.clone()),
                rdocx::ContentFragment::paragraph(paragraph)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let current = candidate
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyValueError::new_err("header/footer disappeared during insertion"))?;
        let location = rdocx::ContentLocation::new(
            current.story().clone(),
            rdocx::StoryItemKind::Paragraph,
            vec![index],
        );
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "HeaderFooter.add_paragraph",
            crate::document::WordPublication::HeaderFooterAppend(self.index),
        )?;
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![PathSeg::Para(index)])
            .map_err(revision_to_pyerr)?;
        let next = document
            .capture_prepared_owner(
                &prepared,
                &WordOwnerLocation::Section(self.index),
                smallvec![],
                WordView::Content,
                None,
            )?
            .0;
        let result = Py::new(
            py,
            PyParagraph::for_story(self.document.clone_ref(py), path, location).bind_prepared(
                py,
                &document,
                &prepared,
                &candidate,
                Some(document.section_child_scope(self.index)?),
            )?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        Ok(result)
    }
    #[getter]
    fn tables<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        self.validate(py)?;
        let document = self.document.borrow(py);
        let mut tables = Vec::new();
        if let Some(story) = document
            .inner
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            for item in document
                .inner
                .story_items(story.story())
                .map_err(|error| rdocx_to_pyerr(py, error))?
            {
                if item.kind() == rdocx::StoryItemKind::Table {
                    let location = item.location().clone();
                    let path = document
                        .revisions
                        .capture(smallvec![PathSeg::Body(location.index_path()[0])]);
                    tables.push(Py::new(
                        py,
                        PyTable::for_story(self.document.clone_ref(py), path, location)
                            .bind_current(py, Some(document.section_child_scope(self.index)?))?,
                    )?);
                }
            }
        }
        PyTuple::new(py, tables)
    }
    fn add_table(&mut self, py: Python<'_>, rows: usize, cols: usize) -> PyResult<Py<PyTable>> {
        self.validate(py)?;
        let fragment = related_table_fragment(rows, cols)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let story = match candidate
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            Some(story) => story.story().clone(),
            None => candidate
                .create_section_story(self.index, self.kind, self.variant)
                .map_err(|error| rdocx_to_pyerr(py, error))?,
        };
        let index = candidate
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .len();
        candidate
            .insert_content(&rdocx::ContentLocation::end(story), fragment)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let story = candidate
            .section_story(self.index, self.kind, self.variant)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyValueError::new_err("header/footer disappeared during insertion"))?;
        let location = rdocx::ContentLocation::new(
            story.story().clone(),
            rdocx::StoryItemKind::Table,
            vec![index],
        );
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "HeaderFooter.add_table",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![PathSeg::Body(index)])
            .map_err(revision_to_pyerr)?;
        let next = document
            .capture_prepared_owner(
                &prepared,
                &WordOwnerLocation::Section(self.index),
                smallvec![],
                WordView::Content,
                None,
            )?
            .0;
        let result = Py::new(
            py,
            PyTable::for_story(self.document.clone_ref(py), path, location).bind_prepared(
                py,
                &document,
                &prepared,
                &candidate,
                Some(document.section_child_scope(self.index)?),
            )?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        Ok(result)
    }
}

#[pyclass(name = "FootnoteEndnote")]
pub struct PyNote {
    document: Py<PyDocument>,
    id: i32,
    endnote: bool,
    path: oxml_py_support::ContentPath,
}
impl PyNote {
    fn story(&self, py: Python<'_>) -> PyResult<rdocx::StoryId> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
                "note",
                "Re-fetch the note after structural mutation.",
            )
            .map_err(|error| crate::stale_to_pyerr(py, error))?;
        let story = if self.endnote {
            document.inner.endnote_story(self.id)
        } else {
            document.inner.footnote_story(self.id)
        };
        story
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyValueError::new_err("note owner no longer exists"))
    }
}
#[pymethods]
impl PyNote {
    #[getter]
    fn id(&self) -> i32 {
        self.id
    }
    #[getter]
    fn paragraphs<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let story = self.story(py)?;
        let document = self.document.borrow(py);
        let mut paragraphs = Vec::new();
        for item in document
            .inner
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            if item.kind() == rdocx::StoryItemKind::Paragraph {
                let location = item.location().clone();
                let path = document
                    .revisions
                    .capture(smallvec![PathSeg::Para(location.index_path()[0])]);
                paragraphs.push(Py::new(
                    py,
                    PyParagraph::for_story(self.document.clone_ref(py), path, location)
                        .bind_current(py, None)?,
                )?);
            }
        }
        PyTuple::new(py, paragraphs)
    }
    fn delete(&self, py: Python<'_>) -> PyResult<()> {
        self.story(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        if self.endnote {
            candidate.remove_endnote(self.id)
        } else {
            candidate.remove_footnote(self.id)
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Note.delete",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }
    #[getter]
    fn tables<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let story = self.story(py)?;
        let document = self.document.borrow(py);
        let mut tables = Vec::new();
        for item in document
            .inner
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            if item.kind() == rdocx::StoryItemKind::Table {
                let location = item.location().clone();
                let path = document
                    .revisions
                    .capture(smallvec![PathSeg::Body(location.index_path()[0])]);
                tables.push(Py::new(
                    py,
                    PyTable::for_story(self.document.clone_ref(py), path, location)
                        .bind_current(py, None)?,
                )?);
            }
        }
        PyTuple::new(py, tables)
    }
    #[pyo3(signature=(text=""))]
    fn add_paragraph(&mut self, py: Python<'_>, text: &str) -> PyResult<Py<PyParagraph>> {
        let story = self.story(py)?;
        let mut paragraph = rdocx_oxml::text::CT_P::new();
        if !text.is_empty() {
            paragraph.add_run(text);
        }
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let index = candidate
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .len();
        candidate
            .insert_content(
                &rdocx::ContentLocation::end(story),
                rdocx::ContentFragment::paragraph(paragraph)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let story = if self.endnote {
            candidate.endnote_story(self.id)
        } else {
            candidate.footnote_story(self.id)
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?
        .ok_or_else(|| PyValueError::new_err("note disappeared during insertion"))?;
        let location =
            rdocx::ContentLocation::new(story, rdocx::StoryItemKind::Paragraph, vec![index]);
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Note.add_paragraph",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![PathSeg::Para(index)])
            .map_err(revision_to_pyerr)?;
        let next = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyParagraph::for_story(self.document.clone_ref(py), path, location)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        Ok(result)
    }
    fn add_table(&mut self, py: Python<'_>, rows: usize, cols: usize) -> PyResult<Py<PyTable>> {
        let story = self.story(py)?;
        let fragment = related_table_fragment(rows, cols)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let index = candidate
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .len();
        candidate
            .insert_content(&rdocx::ContentLocation::end(story), fragment)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let story = if self.endnote {
            candidate.endnote_story(self.id)
        } else {
            candidate.footnote_story(self.id)
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?
        .ok_or_else(|| PyValueError::new_err("note disappeared during insertion"))?;
        let location = rdocx::ContentLocation::new(story, rdocx::StoryItemKind::Table, vec![index]);
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Note.add_table",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![PathSeg::Body(index)])
            .map_err(revision_to_pyerr)?;
        let next = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyTable::for_story(self.document.clone_ref(py), path, location)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        self.path = next;
        Ok(result)
    }
}

#[pyclass(name = "Settings")]
pub struct PySettings {
    document: Py<PyDocument>,
}
#[pymethods]
impl PySettings {
    #[getter]
    fn track_revisions(&self, py: Python<'_>) -> Option<bool> {
        self.document.borrow(py).inner.track_revisions()
    }
    #[setter]
    fn set_track_revisions(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let mut document = self.document.borrow_mut(py);
        if document.inner.track_revisions() == value {
            return Ok(());
        }
        match value {
            Some(value) => document.inner.set_track_revisions(value),
            None => document.inner.remove_track_revisions().map(|_| ()),
        }
        .map_err(|error| rdocx_to_pyerr(py, error))
    }
    #[getter]
    fn odd_and_even_pages_header_footer(&self, py: Python<'_>) -> bool {
        self.document.borrow(py).inner.even_and_odd_headers()
    }
    #[setter]
    fn set_odd_and_even_pages_header_footer(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut document = self.document.borrow_mut(py);
        if document.inner.even_and_odd_headers() == value {
            return Ok(());
        }
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .set_even_and_odd_headers(value)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Settings.odd_and_even_pages_header_footer",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
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

impl PySection {
    fn live_section_alternative(&self, property: &str) -> PyResult<()> {
        Err(pyo3::exceptions::PyNotImplementedError::new_err(format!(
            "Section is a snapshot; use document.section({}).{}",
            self.ordinal, property,
        )))
    }
}

#[pymethods]
impl PySection {
    #[getter]
    fn left_margin(&self) -> PyResult<()> {
        self.live_section_alternative("left_margin")
    }
    #[setter]
    fn set_left_margin(&self, _value: &Bound<'_, PyAny>) -> PyResult<()> {
        self.live_section_alternative("left_margin")
    }
    #[getter]
    fn header(&self) -> PyResult<()> {
        self.live_section_alternative("header")
    }
    #[getter]
    fn footer(&self) -> PyResult<()> {
        self.live_section_alternative("footer")
    }
    #[getter]
    fn first_page_header(&self) -> PyResult<()> {
        self.live_section_alternative("first_page_header")
    }
    #[getter]
    fn different_first_page_header_footer(&self) -> PyResult<()> {
        self.live_section_alternative("different_first_page_header_footer")
    }
    #[setter]
    fn set_different_first_page_header_footer(&self, _value: &Bound<'_, PyAny>) -> PyResult<()> {
        self.live_section_alternative("different_first_page_header_footer")
    }

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

/// Owned, read-only snapshot of one qualified direct-body chart occurrence.
#[pyclass(name = "Chart", frozen)]
pub struct PyChart {
    chart_type: Option<i64>,
    data: rdocx::ChartData,
}

#[pymethods]
impl PyChart {
    #[getter]
    fn chart_type(&self) -> Option<i64> {
        self.chart_type
    }
    #[getter]
    fn categories<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.data.categories)
    }
    #[getter]
    fn series<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let class = py.import("rdocx")?.getattr("CategorySeriesData")?;
        let values = self
            .data
            .series
            .iter()
            .map(|(name, values)| {
                class.call1((name, values.clone(), self.data.number_format.clone()))
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, values)
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
    pub font: Option<String>,
    pub text: Option<String>,
    pub start: Option<u32>,
    pub left_indent: Option<i64>,
    pub hanging_indent: Option<i64>,
}

#[pymethods]
impl PyListLevel {
    #[staticmethod]
    #[pyo3(signature = (*, checked = false))]
    fn checklist(checked: bool) -> Self {
        Self {
            format: "bullet".to_owned(),
            font: None,
            text: Some(if checked { "☑" } else { "☐" }.to_owned()),
            start: None,
            left_indent: None,
            hanging_indent: None,
        }
    }

    #[new]
    #[pyo3(signature = (*, format = "decimal", text = None, start = None, left_indent = None, hanging_indent = None, font = None))]
    fn new(
        format: &str,
        text: Option<String>,
        start: Option<u32>,
        left_indent: Option<i64>,
        hanging_indent: Option<i64>,
        font: Option<String>,
    ) -> PyResult<Self> {
        if let rdocx::ListNumberFormat::Other(_) = rdocx::ListNumberFormat::from_name(format) {
            return Err(PyValueError::new_err(format!(
                "'{format}' is not a Word numbering format"
            )));
        }
        if hanging_indent.is_some_and(|value| value < 0) {
            return Err(PyValueError::new_err("hanging_indent cannot be negative"));
        }
        if let Some(name) = &font {
            if name.trim().is_empty() {
                return Err(PyValueError::new_err("font must be nonempty"));
            }
            if name.chars().any(|c| {
                (c < ' ' && !matches!(c, '\t' | '\n' | '\r'))
                    || matches!(c, '\u{fffe}' | '\u{ffff}')
            }) {
                return Err(PyValueError::new_err("font contains a non-XML character"));
            }
        }
        Ok(Self {
            format: format.to_owned(),
            font,
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
        if let Some(name) = &self.font {
            level = level.marker_properties(rdocx_oxml::properties::CT_RPr {
                font_ascii: Some(name.clone()),
                font_hansi: Some(name.clone()),
                ..Default::default()
            });
        }
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
    alignment: Option<i32>,
    line_spacing: Option<Py<PyAny>>,
    underline: Option<Py<PyAny>>,
    keep_with_next: Option<bool>,
    keep_together: Option<bool>,
    page_break_before: Option<bool>,
    borders: Option<Py<PyAny>>,
    shading: Option<Py<PyAny>>,
    tab_stops: Option<Py<PyAny>>,
}

impl StyleFormatting {
    fn run_properties(
        &self,
        py: Python<'_>,
        existing: Option<&rdocx::CT_RPr>,
    ) -> PyResult<Option<rdocx::CT_RPr>> {
        let supplied = self.font_name.is_some()
            || self.font_size.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.color.is_some()
            || self.underline.is_some();
        if !supplied {
            return Ok(None);
        }
        let mut properties = existing.cloned().unwrap_or_default();
        if let Some(name) = &self.font_name {
            properties.font_ascii = Some(name.clone());
            properties.font_hansi = Some(name.clone());
            properties.font_east_asia = Some(name.clone());
            properties.font_cs = Some(name.clone());
        }
        if let Some(emu) = self.font_size {
            let half_points = emu / 6350;
            let size = u32::try_from(half_points).map_err(|_| {
                PyValueError::new_err("style font_size must fit nonnegative half-points")
            })?;
            if emu < 0 {
                return Err(PyValueError::new_err("style font_size cannot be negative"));
            }
            properties.sz = Some(rdocx::HalfPoint(size));
            properties.sz_cs = Some(rdocx::HalfPoint(size));
        }
        if let Some(value) = self.bold {
            properties.bold = Some(value);
            properties.bold_cs = Some(value);
        }
        if let Some(value) = self.italic {
            properties.italic = Some(value);
            properties.italic_cs = Some(value);
        }
        if let Some((r, g, b)) = self.color {
            properties.color = Some(format!("{r:02X}{g:02X}{b:02X}"));
        }
        if let Some(value) = &self.underline {
            let value = value.bind(py);
            let code = if value.is_instance_of::<PyBool>() {
                i32::from(value.extract::<bool>()?)
            } else {
                value.extract::<i32>()?
            };
            use rdocx_oxml::shared::ST_Underline;
            properties.underline = Some(match code {
                0 => ST_Underline::None,
                1 => ST_Underline::Single,
                2 => ST_Underline::Words,
                3 => ST_Underline::Double,
                4 => ST_Underline::Dotted,
                6 => ST_Underline::Thick,
                7 => ST_Underline::Dash,
                9 => ST_Underline::DotDash,
                10 => ST_Underline::DotDotDash,
                11 => ST_Underline::Wave,
                _ => return Err(PyValueError::new_err("unsupported underline style")),
            });
        }
        Ok(Some(properties))
    }

    fn paragraph_properties(
        &self,
        py: Python<'_>,
        existing: Option<&rdocx::CT_PPr>,
    ) -> PyResult<Option<rdocx::CT_PPr>> {
        let supplied = self.space_before.is_some()
            || self.space_after.is_some()
            || self.left_indent.is_some()
            || self.right_indent.is_some()
            || self.first_line_indent.is_some()
            || self.alignment.is_some()
            || self.line_spacing.is_some()
            || self.keep_with_next.is_some()
            || self.keep_together.is_some()
            || self.page_break_before.is_some()
            || self.borders.is_some()
            || self.shading.is_some()
            || self.tab_stops.is_some();
        if !supplied {
            return Ok(None);
        }
        let mut properties = existing.cloned().unwrap_or_default();
        let twips = |emu: i64| -> PyResult<rdocx::Twips> {
            i32::try_from(emu / 635)
                .map(rdocx::Twips)
                .map_err(|_| PyValueError::new_err("style length exceeds signed twip range"))
        };
        if let Some(value) = self.space_before {
            properties.space_before = Some(twips(value)?);
        }
        if let Some(value) = self.space_after {
            properties.space_after = Some(twips(value)?);
        }
        if let Some(value) = self.left_indent {
            properties.ind_left = Some(twips(value)?);
        }
        if let Some(value) = self.right_indent {
            properties.ind_right = Some(twips(value)?);
        }
        if let Some(value) = self.first_line_indent {
            let value = twips(value)?.0;
            properties.ind_first_line = (value >= 0).then_some(rdocx::Twips(value));
            properties.ind_hanging = if value < 0 {
                Some(rdocx::Twips(value.checked_abs().ok_or_else(|| {
                    PyValueError::new_err("style hanging indent exceeds signed twip range")
                })?))
            } else {
                None
            };
        }
        if let Some(value) = self.alignment {
            properties.jc = Some(match crate::formatting::alignment_from_int(value)? {
                rdocx::Alignment::Left => rdocx_oxml::ST_Jc::Left,
                rdocx::Alignment::Center => rdocx_oxml::ST_Jc::Center,
                rdocx::Alignment::Right => rdocx_oxml::ST_Jc::Right,
                rdocx::Alignment::Justify => rdocx_oxml::ST_Jc::Both,
            });
        }
        if let Some(value) = &self.line_spacing {
            let value = value.bind(py);
            if value.is_instance_of::<PyFloat>() {
                let multiple = value.extract::<f64>()?;
                let twip = multiple * 240.0;
                if !multiple.is_finite() || twip < i32::MIN as f64 || twip > i32::MAX as f64 {
                    return Err(PyValueError::new_err(
                        "style line_spacing must be finite and fit signed twips",
                    ));
                }
                properties.line_spacing = Some(rdocx::Twips(twip as i32));
                properties.line_rule = Some("auto".to_owned());
            } else {
                properties.line_spacing = Some(twips(value.extract()?)?);
                properties.line_rule = Some("exact".to_owned());
            }
        }
        if let Some(value) = self.keep_with_next {
            properties.keep_next = Some(value);
        }
        if let Some(value) = self.keep_together {
            properties.keep_lines = Some(value);
        }
        if let Some(value) = self.page_break_before {
            properties.page_break_before = Some(value);
        }
        if let Some(value) = &self.shading {
            let fill = crate::formatting::paragraph_color_hex(value.bind(py), "style shading")?;
            let shading = properties.shading.get_or_insert_with(|| {
                Box::new(rdocx_oxml::properties::CT_Shd {
                    val: "clear".to_owned(),
                    color: Some("auto".to_owned()),
                    ..Default::default()
                })
            });
            shading.fill = Some(fill);
            shading.theme_fill = None;
            shading.theme_fill_tint = None;
            shading.theme_fill_shade = None;
        }
        if let Some(value) = &self.borders {
            let mapping = value.bind(py).cast::<PyMapping>()?;
            let items = mapping
                .items()?
                .extract::<Vec<(String, (String, u32, Py<PyAny>))>>()?;
            if items.is_empty() {
                properties.borders = None;
            } else {
                for (edge, (style, size, color)) in items {
                    let edge = crate::formatting::paragraph_border_edge(&edge)?;
                    let style = crate::table::border_style_from_name(&style)?;
                    let color = crate::formatting::paragraph_color_hex(
                        color.bind(py),
                        "style borders color",
                    )?;
                    let borders = properties
                        .borders
                        .get_or_insert_with(|| Box::new(rdocx_oxml::CT_PBdr::default()));
                    let slot = match edge {
                        rdocx::ParagraphBorderEdge::Top => &mut borders.top,
                        rdocx::ParagraphBorderEdge::Bottom => &mut borders.bottom,
                        rdocx::ParagraphBorderEdge::Left => &mut borders.left,
                        rdocx::ParagraphBorderEdge::Right => &mut borders.right,
                        rdocx::ParagraphBorderEdge::Between => &mut borders.between,
                        rdocx::ParagraphBorderEdge::Bar => &mut borders.bar,
                    };
                    use rdocx_oxml::ST_Border;
                    let style = match style {
                        rdocx::BorderStyle::None => ST_Border::None,
                        rdocx::BorderStyle::Single => ST_Border::Single,
                        rdocx::BorderStyle::Thick => ST_Border::Thick,
                        rdocx::BorderStyle::Double => ST_Border::Double,
                        rdocx::BorderStyle::Dotted => ST_Border::Dotted,
                        rdocx::BorderStyle::Dashed => ST_Border::Dashed,
                        rdocx::BorderStyle::DotDash => ST_Border::DotDash,
                        rdocx::BorderStyle::Wave => ST_Border::Wave,
                    };
                    let border = slot.get_or_insert_with(|| rdocx_oxml::CT_BorderEdge::new(style));
                    border.val = style;
                    border.sz = Some(size);
                    border.color = Some(color);
                }
            }
        }
        if let Some(value) = &self.tab_stops {
            let entries = value.bind(py).extract::<Vec<(i64, i32, Option<i32>)>>()?;
            let mut tabs = Vec::with_capacity(entries.len());
            for (position, alignment, leader) in entries {
                use rdocx_oxml::{ST_TabJc, ST_TabLeader};
                let val = match crate::formatting::tab_alignment_from_int(alignment)? {
                    rdocx::TabAlignment::Left => ST_TabJc::Left,
                    rdocx::TabAlignment::Center => ST_TabJc::Center,
                    rdocx::TabAlignment::Right => ST_TabJc::Right,
                    rdocx::TabAlignment::Decimal => ST_TabJc::Decimal,
                    rdocx::TabAlignment::Bar => ST_TabJc::Bar,
                    rdocx::TabAlignment::Clear => ST_TabJc::Clear,
                    rdocx::TabAlignment::Num => ST_TabJc::Num,
                    rdocx::TabAlignment::List => ST_TabJc::List,
                    rdocx::TabAlignment::Start => ST_TabJc::Start,
                    rdocx::TabAlignment::End => ST_TabJc::End,
                };
                let leader =
                    crate::formatting::tab_leader_from_int(leader)?.map(|value| match value {
                        rdocx::TabLeader::None => ST_TabLeader::None,
                        rdocx::TabLeader::Dot => ST_TabLeader::Dot,
                        rdocx::TabLeader::Hyphen => ST_TabLeader::Hyphen,
                        rdocx::TabLeader::Underscore => ST_TabLeader::Underscore,
                        rdocx::TabLeader::Heavy => ST_TabLeader::Heavy,
                        rdocx::TabLeader::MiddleDot => ST_TabLeader::MiddleDot,
                    });
                let mut tab = rdocx_oxml::CT_TabStop::new(val, twips(position)?);
                tab.leader = leader;
                tabs.push(tab);
            }
            tabs.sort_by_key(|tab| tab.pos.0);
            properties.tabs = (!tabs.is_empty()).then_some(rdocx_oxml::CT_Tabs { tabs });
            properties
                .refresh_tabs_source(None)
                .map_err(|error| PyValueError::new_err(error.to_string()))?;
        }
        Ok(Some(properties))
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

#[pyclass(name = "CustomFileTime", frozen, get_all, skip_from_py_object)]
#[derive(Clone)]
pub struct PyCustomFileTime {
    value: String,
}

fn valid_custom_filetime(value: &str) -> bool {
    let b = value.as_bytes();
    if b.len() < 20
        || !b.is_ascii()
        || b[4] != b'-'
        || b[7] != b'-'
        || b[10] != b'T'
        || b[13] != b':'
        || b[16] != b':'
    {
        return false;
    }
    let number = |start: usize, end: usize| -> Option<u32> {
        b[start..end]
            .iter()
            .all(u8::is_ascii_digit)
            .then(|| value[start..end].parse().ok())
            .flatten()
    };
    let (Some(year), Some(month), Some(day), Some(hour), Some(minute), Some(second)) = (
        number(0, 4),
        number(5, 7),
        number(8, 10),
        number(11, 13),
        number(14, 16),
        number(17, 19),
    ) else {
        return false;
    };
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if leap {
                29
            } else {
                28
            }
        }
        _ => return false,
    };
    if year == 0 || day == 0 || day > days || hour > 23 || minute > 59 || second > 59 {
        return false;
    }
    let mut at = 19;
    if b.get(at) == Some(&b'.') {
        at += 1;
        let start = at;
        while b.get(at).is_some_and(u8::is_ascii_digit) {
            at += 1;
        }
        if at == start {
            return false;
        }
    }
    if b.get(at) == Some(&b'Z') {
        return at + 1 == b.len();
    }
    if !matches!(b.get(at), Some(b'+') | Some(b'-')) || at + 6 != b.len() || b[at + 3] != b':' {
        return false;
    }
    number(at + 1, at + 3).is_some_and(|n| n <= 23)
        && number(at + 4, at + 6).is_some_and(|n| n <= 59)
}

#[pymethods]
impl PyCustomFileTime {
    #[new]
    fn new(value: String) -> PyResult<Self> {
        if !valid_custom_filetime(&value) {
            return Err(PyValueError::new_err(
                "CustomFileTime requires a valid Gregorian RFC3339 date with explicit timezone and no leap second",
            ));
        }
        Ok(Self { value })
    }
}

#[pyclass(name = "CustomProperty", frozen, skip_from_py_object)]
#[derive(Clone)]
pub struct PyCustomProperty {
    property: rdocx::CustomProperty,
}

#[pymethods]
impl PyCustomProperty {
    #[getter]
    fn name(&self) -> Option<String> {
        self.property.name.clone()
    }
    #[getter]
    fn pid(&self) -> i32 {
        self.property.pid
    }
    #[getter]
    fn fmtid(&self) -> String {
        self.property.fmtid.clone()
    }
    #[getter]
    fn r#type(&self) -> &'static str {
        use rdocx::CustomPropertyValue as V;
        match &self.property.value {
            V::Lpstr(_) => "lpstr",
            V::Lpwstr(_) => "lpwstr",
            V::I4(_) => "i4",
            V::R8(_) => "r8",
            V::Bool(_) => "bool",
            V::FileTime(_) => "filetime",
            V::Empty => "empty",
            V::Raw(_) => "raw",
        }
    }
    #[getter]
    fn value(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        use rdocx::CustomPropertyValue as V;
        Ok(match &self.property.value {
            V::Lpstr(v) | V::Lpwstr(v) | V::FileTime(v) => v.into_pyobject(py)?.into_any().unbind(),
            V::I4(v) => v.into_pyobject(py)?.into_any().unbind(),
            V::R8(v) => v.into_pyobject(py)?.into_any().unbind(),
            V::Bool(v) => v.into_pyobject(py)?.to_owned().into_any().unbind(),
            V::Empty => py.None(),
            V::Raw(v) => PyBytes::new(py, v).into_any().unbind(),
        })
    }
}

#[pyclass(name = "DefaultParagraphFormat", frozen, get_all, skip_from_py_object)]
pub struct PyDefaultParagraphFormat {
    pub space_before: Option<Py<PyAny>>,
    pub space_after: Option<Py<PyAny>>,
    pub line_spacing: Option<Py<PyAny>>,
}

fn length_from_twips(py: Python<'_>, twips: Option<rdocx::Twips>) -> PyResult<Option<Py<PyAny>>> {
    twips
        .map(|twips| crate::length_object(py, rdocx::Length::twips(twips.0)))
        .transpose()
}

fn line_spacing_value(py: Python<'_>, properties: &rdocx::CT_PPr) -> PyResult<Option<Py<PyAny>>> {
    let Some(line) = properties.line_spacing else {
        return Ok(None);
    };
    match properties.line_rule.as_deref() {
        Some("exact") | Some("atLeast") => {
            crate::length_object(py, rdocx::Length::twips(line.0)).map(Some)
        }
        None | Some("auto") => Ok(Some(
            (line.0 as f64 / 240.0)
                .into_pyobject(py)?
                .to_owned()
                .unbind()
                .into_any(),
        )),
        _ => Ok(None),
    }
}

fn default_paragraph_format_snapshot(
    py: Python<'_>,
    properties: Option<&rdocx::CT_PPr>,
) -> PyResult<PyDefaultParagraphFormat> {
    let Some(properties) = properties else {
        return Ok(PyDefaultParagraphFormat {
            space_before: None,
            space_after: None,
            line_spacing: None,
        });
    };
    Ok(PyDefaultParagraphFormat {
        space_before: length_from_twips(py, properties.space_before)?,
        space_after: length_from_twips(py, properties.space_after)?,
        line_spacing: line_spacing_value(py, properties)?,
    })
}

fn length_keyword(
    py: Python<'_>,
    value: &Bound<'_, PyAny>,
    name: &str,
) -> PyResult<Option<rdocx::Twips>> {
    if value.is_none() {
        return Ok(None);
    }
    let message = format!("{name} requires a Length or None");
    let length_type = py.import("rdocx")?.getattr("Length")?;
    if value.is_instance_of::<PyBool>()
        || value.is_instance_of::<PyFloat>()
        || !value.is_instance(&length_type)?
    {
        return Err(PyTypeError::new_err(message));
    }
    let emu = value
        .extract::<i64>()
        .map_err(|_| PyValueError::new_err(format!("{name} must be a representable Length")))?;
    let twips = i32::try_from(emu / 635)
        .map_err(|_| PyValueError::new_err(format!("{name} must fit signed twips")))?;
    Ok(Some(rdocx::Twips(twips)))
}

#[pyclass(name = "CustomProperties")]
pub struct PyCustomProperties {
    document: Py<PyDocument>,
}

#[pymethods]
impl PyCustomProperties {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.document.borrow(py).inner.custom_properties().len()
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let properties = self
            .document
            .borrow(py)
            .inner
            .custom_properties()
            .iter()
            .cloned()
            .map(|property| PyCustomProperty { property })
            .collect::<Vec<_>>();
        PyTuple::new(py, properties)?
            .call_method0("__iter__")
            .map(Bound::unbind)
    }

    fn __getitem__(&self, py: Python<'_>, name: &str) -> PyResult<PyCustomProperty> {
        self.get(py, name)
            .ok_or_else(|| PyKeyError::new_err(name.to_owned()))
    }

    fn items<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let pairs = self
            .document
            .borrow(py)
            .inner
            .custom_properties()
            .iter()
            .map(|property| {
                (
                    property.name.clone().unwrap_or_default(),
                    PyCustomProperty {
                        property: property.clone(),
                    },
                )
            })
            .collect::<Vec<_>>();
        PyTuple::new(py, pairs)
    }

    fn __setitem__(&self, py: Python<'_>, name: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        self.set(py, name, value)
    }
    fn __delitem__(&self, py: Python<'_>, name: &str) -> PyResult<()> {
        if name.is_empty() || !self.delete(py, name)? {
            Err(pyo3::exceptions::PyKeyError::new_err(name.to_owned()))
        } else {
            Ok(())
        }
    }
    fn __contains__(&self, py: Python<'_>, name: &str) -> bool {
        !name.is_empty() && self.get(py, name).is_some()
    }
    fn get(&self, py: Python<'_>, name: &str) -> Option<PyCustomProperty> {
        self.document
            .borrow(py)
            .inner
            .custom_property(name)
            .cloned()
            .map(|property| PyCustomProperty { property })
    }
    fn set(&self, py: Python<'_>, name: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        use rdocx::CustomPropertyValue as V;
        if name.is_empty() {
            return Err(PyValueError::new_err(
                "custom property name must not be empty",
            ));
        }
        let authored = if value.is_instance_of::<PyBool>() {
            V::Bool(value.extract()?)
        } else if value.is_instance_of::<PyInt>() {
            V::I4(
                value
                    .extract()
                    .map_err(|_| PyOverflowError::new_err("custom integer must fit signed32"))?,
            )
        } else if value.is_instance_of::<PyFloat>() {
            let n: f64 = value.extract()?;
            if !n.is_finite() {
                return Err(PyValueError::new_err("custom float must be finite"));
            }
            V::R8(n)
        } else if value.is_instance_of::<PyString>() {
            V::Lpwstr(value.extract()?)
        } else if let Ok(v) = value.extract::<PyRef<'_, PyCustomFileTime>>() {
            V::FileTime(v.value.clone())
        } else {
            return Err(PyTypeError::new_err(
                "custom property requires str, bool, signed32 int, finite float or CustomFileTime",
            ));
        };
        let mut document = self.document.borrow_mut(py);
        let mut property = if let Some(existing) = document.inner.custom_property(name) {
            existing.clone()
        } else {
            let used: std::collections::BTreeSet<i32> = document
                .inner
                .custom_properties()
                .iter()
                .map(|p| p.pid)
                .collect();
            let pid = (2..=i32::MAX)
                .find(|pid| !used.contains(pid))
                .ok_or_else(|| PyOverflowError::new_err("no custom property PID available"))?;
            rdocx::CustomProperty {
                fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}".into(),
                pid,
                name: Some(name.into()),
                value: V::Empty,
            }
        };
        let authored = match (&property.value, authored) {
            (V::Lpstr(_), V::Lpwstr(v)) => V::Lpstr(v),
            (_, v) => v,
        };
        if property.value == authored {
            return Ok(());
        }
        property.value = authored;
        document
            .inner
            .set_custom_property(property)
            .map_err(|e| rdocx_to_pyerr(py, e))
    }
    fn delete(&self, py: Python<'_>, name: &str) -> PyResult<bool> {
        self.document
            .borrow_mut(py)
            .inner
            .remove_custom_property(name)
            .map(|p| p.is_some())
            .map_err(|e| rdocx_to_pyerr(py, e))
    }
}

// These concrete field selectors update the latest owned model, preserving its
// opaque child order and namespace metadata. Reading never creates a part.
type AppTextField = fn(&mut rdocx::AppProperties) -> &mut Option<String>;
#[pyclass(name = "AppProperties")]
pub struct PyAppProperties {
    document: Py<PyDocument>,
}
impl PyAppProperties {
    fn text(&self, py: Python<'_>, field: AppTextField) -> Option<String> {
        let mut p = self
            .document
            .borrow(py)
            .inner
            .application_properties()?
            .clone();
        field(&mut p).clone()
    }
    fn set_text(&self, py: Python<'_>, field: AppTextField, value: Option<String>) -> PyResult<()> {
        let mut document = self.document.borrow_mut(py);
        let mut p = document
            .inner
            .application_properties()
            .cloned()
            .unwrap_or_default();
        if *field(&mut p) == value {
            return Ok(());
        }
        *field(&mut p) = value;
        document
            .inner
            .set_application_properties(p)
            .map_err(|e| rdocx_to_pyerr(py, e))
    }
}
#[pymethods]
impl PyAppProperties {
    #[getter]
    fn company(&self, py: Python<'_>) -> Option<String> {
        self.text(py, |p| &mut p.company)
    }
    #[setter]
    fn set_company(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |p| &mut p.company, value)
    }
    #[getter]
    fn manager(&self, py: Python<'_>) -> Option<String> {
        self.text(py, |p| &mut p.manager)
    }
    #[setter]
    fn set_manager(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |p| &mut p.manager, value)
    }
    #[getter]
    fn template(&self, py: Python<'_>) -> Option<String> {
        self.text(py, |p| &mut p.template)
    }
    #[setter]
    fn set_template(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |p| &mut p.template, value)
    }
    #[getter]
    fn application(&self, py: Python<'_>) -> Option<String> {
        self.text(py, |p| &mut p.application)
    }
    #[setter]
    fn set_application(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |p| &mut p.application, value)
    }
    #[getter]
    fn application_version(&self, py: Python<'_>) -> Option<String> {
        self.text(py, |p| &mut p.application_version)
    }
    #[setter]
    fn set_application_version(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |p| &mut p.application_version, value)
    }
    #[getter]
    fn presentation_format(&self, py: Python<'_>) -> Option<String> {
        self.text(py, |p| &mut p.presentation_format)
    }
    #[setter]
    fn set_presentation_format(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.set_text(py, |p| &mut p.presentation_format, value)
    }
    #[getter]
    fn pages(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.pages)
    }
    #[getter]
    fn words(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.words)
    }
    #[getter]
    fn characters(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.characters)
    }
    #[getter]
    fn characters_with_spaces(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.characters_with_spaces)
    }
    #[getter]
    fn lines(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.lines)
    }
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.paragraphs)
    }
    #[getter]
    fn slides(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.slides)
    }
    #[getter]
    fn notes(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.notes)
    }
    #[getter]
    fn total_time(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.total_time)
    }
    #[getter]
    fn hidden_slides(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.hidden_slides)
    }
    #[getter]
    fn multimedia_clips(&self, py: Python<'_>) -> Option<i32> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.multimedia_clips)
    }
    #[getter]
    fn scale_crop(&self, py: Python<'_>) -> Option<bool> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.scale_crop)
    }
    #[getter]
    fn links_up_to_date(&self, py: Python<'_>) -> Option<bool> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.links_up_to_date)
    }
    #[getter]
    fn shared_document(&self, py: Python<'_>) -> Option<bool> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.shared_document)
    }
    #[getter]
    fn hyperlinks_changed(&self, py: Python<'_>) -> Option<bool> {
        self.document
            .borrow(py)
            .inner
            .application_properties()
            .and_then(|p| p.hyperlinks_changed)
    }
}

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

// Only current physical owners are retained. A committed mutation supplies the
// replacement inventory before any state is published.
#[derive(Clone)]
#[cfg_attr(any(test, debug_assertions), derive(Debug, PartialEq, Eq))]
pub(crate) enum WordOwnerLocation {
    Paragraph(rdocx::ContentLocation),
    Run(rdocx::ContentLocation, rdocx_oxml::text::AcceptedRunPath),
    Table(rdocx::ContentLocation),
    Cell {
        table: rdocx::ContentLocation,
        row: usize,
        cell: usize,
        story: rdocx::StoryId,
    },
    Section(usize),
    Hyperlink(rdocx::ContentLocation, usize),
}

#[derive(Clone, Default)]
#[cfg_attr(any(test, debug_assertions), derive(Debug, PartialEq, Eq))]
enum WordReadFacts {
    #[default]
    None,
    Paragraph(String),
    // One entry per modeled row/cell, containing the facade paragraph count.
    Table(Vec<Vec<usize>>),
}

#[cfg_attr(any(test, debug_assertions), derive(Debug, PartialEq, Eq))]
struct WordOwnerAddress {
    location: WordOwnerLocation,
    parent: Option<usize>,
    body_slot: Option<usize>,
    projection_ordinal: Option<usize>,
    read_facts: WordReadFacts,
}

pub(crate) struct WordOwner {
    location: WordOwnerLocation,
    parent: Option<usize>,
    body_slot: Option<usize>,
    projection_ordinal: Option<usize>,
    read_facts: WordReadFacts,
    owner: RevisionState,
    content: Option<RevisionState>,
    properties: Option<RevisionState>,
    grid: Option<RevisionState>,
}

#[derive(Clone, Copy)]
pub(crate) enum WordView {
    Owner,
    Content,
    Properties,
    Grid,
}

struct WordOwnerIndexes {
    by_generation: IndexMap<u64, usize>,
    by_route: IndexMap<u64, Vec<usize>>,
    direct_body_paragraphs: IndexMap<usize, usize>,
    body_paragraphs: Option<Vec<Option<rdocx::ContentLocation>>>,
    body_tables: Vec<Option<rdocx::ContentLocation>>,
    cell_paragraphs: HashMap<(usize, usize, usize), Vec<rdocx::ContentLocation>>,
    paragraphs_by_parent: IndexMap<usize, Vec<usize>>,
    tables_by_parent: IndexMap<usize, Vec<usize>>,
}

/// A multiply-rotate hasher over machine words, for owner routes and the
/// owner index maps. Their keys are process-local integers and route
/// hashes whose bucket members are still compared with `same_route`, so
/// they need no keyed hash, and a body edit indexes every owner again.
#[derive(Default)]
struct IndexHasher(u64);

/// A map of the owner indexes.
type IndexMap<K, V> = HashMap<K, V, std::hash::BuildHasherDefault<IndexHasher>>;

impl Hasher for IndexHasher {
    fn finish(&self) -> u64 {
        self.0
    }

    fn write(&mut self, bytes: &[u8]) {
        for chunk in bytes.chunks(8) {
            let mut word = [0; 8];
            word[..chunk.len()].copy_from_slice(chunk);
            self.write_u64(u64::from_le_bytes(word));
        }
    }

    fn write_u64(&mut self, value: u64) {
        self.0 = (self.0.rotate_left(5) ^ value).wrapping_mul(0x517c_c1b7_2722_0a95);
    }

    fn write_usize(&mut self, value: usize) {
        self.write_u64(value as u64);
    }
}

impl WordOwnerLocation {
    fn route_hash(&self) -> u64 {
        let mut hash = IndexHasher::default();
        std::mem::discriminant(self).hash(&mut hash);
        let mut content = |location: &rdocx::ContentLocation| {
            let story = location.story();
            story.kind().hash(&mut hash);
            // Equal routes share their part, so its length tells parts apart
            // well enough for buckets.
            story.part_name().len().hash(&mut hash);
            story.owner_index().hash(&mut hash);
            location.item_kind().hash(&mut hash);
            location.index_path().hash(&mut hash);
        };
        match self {
            Self::Paragraph(location) | Self::Table(location) => content(location),
            Self::Run(location, path) => {
                content(location);
                path.hash(&mut hash);
            }
            Self::Cell {
                table, row, cell, ..
            } => {
                content(table);
                row.hash(&mut hash);
                cell.hash(&mut hash);
            }
            Self::Section(index) => index.hash(&mut hash),
            Self::Hyperlink(location, index) => {
                content(location);
                index.hash(&mut hash);
            }
        }
        hash.finish()
    }
}

impl WordOwnerIndexes {
    fn prepare(owners: &[WordOwner], projections: WordIndexProjections) -> PyResult<Self> {
        let WordIndexProjections {
            body_paragraphs,
            cell_paragraphs,
            table_slots,
        } = projections;
        let mut indexes = Self {
            by_generation: IndexMap::default(),
            by_route: IndexMap::default(),
            direct_body_paragraphs: IndexMap::default(),
            body_paragraphs,
            body_tables: Vec::new(),
            cell_paragraphs,
            paragraphs_by_parent: IndexMap::default(),
            tables_by_parent: IndexMap::default(),
        };
        let fail =
            |_| pyo3::exceptions::PyMemoryError::new_err("Word owner index allocation failed");
        indexes
            .by_generation
            .try_reserve(owners.len())
            .map_err(fail)?;
        indexes.by_route.try_reserve(owners.len()).map_err(fail)?;
        indexes
            .direct_body_paragraphs
            .try_reserve(owners.len())
            .map_err(fail)?;
        let body_tables = indexes.index_owners(owners);
        indexes.body_tables = table_slots
            .into_iter()
            .map(|slot| slot.and_then(|slot| body_tables.get(&slot).cloned()))
            .collect();
        Ok(indexes)
    }

    /// Index `owners` again after they moved in place, keeping the document
    /// projections, which the caller moves with them.
    fn index_owners_again(&mut self, owners: &[WordOwner]) {
        self.by_generation.clear();
        self.by_generation.reserve(owners.len());
        // Route buckets keep their allocations, and those left empty go.
        self.by_route.values_mut().for_each(Vec::clear);
        self.by_route
            .reserve(owners.len().saturating_sub(self.by_route.len()));
        self.direct_body_paragraphs.clear();
        self.paragraphs_by_parent.clear();
        self.tables_by_parent.clear();
        self.index_owners(owners);
        self.by_route.retain(|_, bucket| !bucket.is_empty());
    }

    /// Index every owner, and return the first body table owner of each
    /// direct body slot.
    fn index_owners(&mut self, owners: &[WordOwner]) -> HashMap<usize, rdocx::ContentLocation> {
        let indexes = self;
        let mut body_tables = HashMap::new();
        for (index, owner) in owners.iter().enumerate() {
            note_fx260_word(1);
            indexes
                .by_generation
                .insert(owner.owner.generation(), index);
            indexes
                .by_route
                .entry(owner.location.route_hash())
                .or_default()
                .push(index);
            match &owner.location {
                WordOwnerLocation::Paragraph(location) => {
                    if location.story().kind() == rdocx::StoryKind::Body
                        && location.index_path().len() == 1
                        && let Some(slot) = owner.body_slot
                    {
                        indexes.direct_body_paragraphs.entry(slot).or_insert(index);
                    }
                    if let Some(parent) = owner.parent {
                        indexes
                            .paragraphs_by_parent
                            .entry(parent)
                            .or_default()
                            .push(index);
                    }
                }
                WordOwnerLocation::Table(location) => {
                    if location.story().kind() == rdocx::StoryKind::Body
                        && let Some(slot) = owner.body_slot
                    {
                        body_tables.entry(slot).or_insert_with(|| location.clone());
                    }
                    if let Some(parent) = owner.parent {
                        indexes
                            .tables_by_parent
                            .entry(parent)
                            .or_default()
                            .push(index);
                    }
                }
                _ => {}
            }
        }
        body_tables
    }

    fn find(&self, owners: &[WordOwner], location: &WordOwnerLocation) -> Option<usize> {
        self.by_route
            .get(&location.route_hash())?
            .iter()
            .copied()
            .find(|index| {
                note_fx260_word(1);
                owners[*index].location.same_route(location)
            })
    }
}

pub(crate) struct PreparedPlainParagraphXml {
    revision: PreparedRevision,
    states: Vec<RevisionState>,
    runs: Vec<usize>,
    newborn: Vec<WordOwner>,
    route_buckets: HashMap<u64, Vec<usize>>,
}

pub(crate) struct PreparedWordRevision {
    pub(crate) revision: PreparedRevision,
    owners: Vec<WordOwner>,
    indexes: WordOwnerIndexes,
    snapshot_scope: RevisionState,
    /// The pieces of the next inventory, unless the rebuild read it.
    pieces: Option<WordPieces>,
}

/// What a revision does with the current owner states: the states of each
/// next owner it keeps, by scope, the scopes of those it gives birth to,
/// in order, the states it retires, and whether the snapshot scope stays.
struct WordRevisionPlan {
    retained: Vec<Vec<Option<RevisionState>>>,
    births: Vec<RevisionScope>,
    retired: Vec<RevisionState>,
    preserve_snapshot: bool,
}

/// Under debug assertions, check a direct read against the whole-source
/// read: the same bytes up to the root `w14` declaration, which a direct
/// read decides from the root and the read item alone.
#[cfg(debug_assertions)]
pub(crate) fn assert_direct_read(
    direct: &[u8],
    whole: impl FnOnce() -> PyResult<Vec<u8>>,
    owner: &str,
) {
    let without_w14 = |xml: &[u8]| {
        let xml = String::from_utf8_lossy(xml).into_owned();
        let root = xml.find('>').unwrap_or(xml.len());
        let declaration = r#" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
        format!(
            "{}{}",
            xml[..root].replacen(declaration, "", 1),
            &xml[root..]
        )
    };
    let whole = guarded(whole).expect("a direct read has a whole-source read");
    assert_eq!(
        without_w14(direct),
        without_w14(&whole),
        "direct {owner} XML"
    );
}

/// Under debug assertions, check a handle's route bound from an owner
/// inventory against the route the document resolves, story fingerprints
/// aside.
#[cfg(debug_assertions)]
pub(crate) fn assert_indexed_route(
    indexed: &rdocx::ContentLocation,
    native: impl FnOnce() -> rdocx::Result<rdocx::ContentLocation>,
    owner: &str,
) {
    #[cfg(test)]
    let visits = fx260_word_visits();
    let native = guarded(native).expect("an indexed route resolves in the document");
    assert_eq!(
        unfingerprinted_location(indexed),
        unfingerprinted_location(&native),
        "indexed {owner} route"
    );
    #[cfg(test)]
    FX260_WORD_VISITS.with(|cell| cell.set(visits));
}

#[cfg(debug_assertions)]
fn unfingerprinted_location(location: &rdocx::ContentLocation) -> rdocx::ContentLocation {
    let story = location.story();
    rdocx::ContentLocation::new(
        rdocx::StoryId::new(story.kind(), story.part_name(), story.owner_index()),
        location.item_kind(),
        location.index_path().to_vec(),
    )
}

#[cfg(debug_assertions)]
fn unfingerprinted_owner_location(location: &WordOwnerLocation) -> WordOwnerLocation {
    owner_location_without(location, |_| true)
}

/// `location` with the fingerprints of the stories that `strip` names set
/// aside.
#[cfg(debug_assertions)]
fn owner_location_without(
    location: &WordOwnerLocation,
    strip: impl Fn(&rdocx::StoryId) -> bool,
) -> WordOwnerLocation {
    let content = |location: &rdocx::ContentLocation| {
        if strip(location.story()) {
            unfingerprinted_location(location)
        } else {
            location.clone()
        }
    };
    match location {
        WordOwnerLocation::Paragraph(location) => WordOwnerLocation::Paragraph(content(location)),
        WordOwnerLocation::Table(location) => WordOwnerLocation::Table(content(location)),
        WordOwnerLocation::Run(location, path) => {
            WordOwnerLocation::Run(content(location), path.clone())
        }
        WordOwnerLocation::Hyperlink(location, index) => {
            WordOwnerLocation::Hyperlink(content(location), *index)
        }
        WordOwnerLocation::Cell {
            table,
            row,
            cell,
            story,
        } => WordOwnerLocation::Cell {
            table: content(table),
            row: *row,
            cell: *cell,
            story: if strip(story) {
                rdocx::StoryId::new(story.kind(), story.part_name(), story.owner_index())
            } else {
                story.clone()
            },
        },
        WordOwnerLocation::Section(index) => WordOwnerLocation::Section(*index),
    }
}

/// Whether a story is the main body, whose fingerprint covers the whole
/// body, so that an inventory read by block and part cannot give it.
#[cfg(debug_assertions)]
fn body_story(story: &rdocx::StoryId) -> bool {
    story.kind() == rdocx::StoryKind::Body
}

/// An owner address with the body story fingerprint set aside.
#[cfg(debug_assertions)]
fn body_unfingerprinted_address(address: &WordOwnerAddress) -> WordOwnerAddress {
    WordOwnerAddress {
        location: owner_location_without(&address.location, body_story),
        parent: address.parent,
        body_slot: address.body_slot,
        projection_ordinal: address.projection_ordinal,
        read_facts: address.read_facts.clone(),
    }
}

/// Under debug assertions, prove an owner inventory and index projections
/// read by block and part equal their rebuild from the whole document,
/// fingerprints included but the body story's: a block's body story is
/// that of its block alone.
#[cfg(debug_assertions)]
fn assert_word_pieces_rebuild(
    document: &rdocx::Document,
    locations: &[WordOwnerAddress],
    projections: &WordIndexProjections,
) {
    #[cfg(test)]
    let visits = fx260_word_visits();
    let expected = guarded(|| word_owner_locations(document, None))
        .expect("an inventory from pieces has a rebuild");
    assert_eq!(
        locations
            .iter()
            .map(body_unfingerprinted_address)
            .collect::<Vec<_>>(),
        expected
            .iter()
            .map(body_unfingerprinted_address)
            .collect::<Vec<_>>(),
        "Word owner inventory from pieces"
    );
    let expected = guarded(|| WordIndexProjections::of(document));
    let paragraphs = |projections: &WordIndexProjections| {
        projections.body_paragraphs.as_ref().map(|locations| {
            locations
                .iter()
                .map(|location| location.as_ref().map(unfingerprinted_location))
                .collect::<Vec<_>>()
        })
    };
    let cells = |projections: &WordIndexProjections| {
        projections
            .cell_paragraphs
            .iter()
            .map(|(key, locations)| {
                // Cell stories keep their fingerprints.
                (*key, locations.clone())
            })
            .collect::<BTreeMap<_, _>>()
    };
    assert_eq!(
        paragraphs(projections),
        paragraphs(&expected),
        "body paragraph projection from pieces"
    );
    assert_eq!(
        cells(projections),
        cells(&expected),
        "cell paragraph projection from pieces"
    );
    assert_eq!(
        projections.table_slots, expected.table_slots,
        "body table projection from pieces"
    );
    #[cfg(test)]
    FX260_WORD_VISITS.with(|cell| cell.set(visits));
}

fn word_owner_capture(
    counter: &RevisionCounter,
    owners: &[WordOwner],
    document_scopes: (RevisionState, &RevisionState),
    index: usize,
    segments: smallvec::SmallVec<[PathSeg; 5]>,
    view: WordView,
    furniture_scope: Option<&RevisionState>,
) -> PyResult<(ContentPath, RevisionState)> {
    let (document_scope, removal_scope) = document_scopes;
    let target = &owners[index];
    // Every structural publisher retires unproved current-owner states. Actual
    // Run handles retain the conservative document epoch and its first cause.
    let mut states = vec![removal_scope.clone(), target.owner.clone()];
    if matches!(&target.location, WordOwnerLocation::Run(..)) {
        states.push(document_scope);
    }
    if let Some(scope) = furniture_scope {
        states.push(scope.clone());
    }
    let child = match view {
        WordView::Owner => None,
        WordView::Content => target.content.as_ref(),
        WordView::Properties => target.properties.as_ref(),
        WordView::Grid => target.grid.as_ref(),
    };
    if let Some(child) = child {
        states.push(child.clone());
    }
    let mut parent = target.parent;
    while let Some(index) = parent {
        let ancestor = &owners[index];
        states.push(ancestor.owner.clone());
        if let Some(content) = &ancestor.content {
            states.push(content.clone());
        }
        if let Some(grid) = &ancestor.grid {
            states.push(grid.clone());
        }
        parent = ancestor.parent;
    }
    let path = counter
        .capture_scopes(segments, &states)
        .map_err(revision_to_pyerr)?;
    Ok((path, target.owner.clone()))
}

fn same_story_owner(a: &rdocx::StoryId, b: &rdocx::StoryId) -> bool {
    a.kind() == b.kind() && a.part_name() == b.part_name() && a.owner_index() == b.owner_index()
}

fn same_content_owner(a: &rdocx::ContentLocation, b: &rdocx::ContentLocation) -> bool {
    same_story_owner(a.story(), b.story())
        && a.item_kind() == b.item_kind()
        && a.index_path() == b.index_path()
}

impl WordOwnerLocation {
    fn same_route(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Paragraph(a), Self::Paragraph(b)) | (Self::Table(a), Self::Table(b)) => {
                same_content_owner(a, b)
            }
            (Self::Run(a, ap), Self::Run(b, bp)) => same_content_owner(a, b) && ap == bp,
            (
                Self::Cell {
                    table: a,
                    row: ar,
                    cell: ac,
                    ..
                },
                Self::Cell {
                    table: b,
                    row: br,
                    cell: bc,
                    ..
                },
            ) => same_content_owner(a, b) && ar == br && ac == bc,
            (Self::Section(a), Self::Section(b)) => a == b,
            (Self::Hyperlink(a, ai), Self::Hyperlink(b, bi)) => {
                same_content_owner(a, b) && ai == bi
            }
            _ => false,
        }
    }

    fn scopes(&self) -> &'static [RevisionScope] {
        match self {
            Self::Paragraph(_) => &[
                RevisionScope::Owner,
                RevisionScope::Text,
                RevisionScope::Text,
            ],
            Self::Table(_) => &[RevisionScope::Owner, RevisionScope::Grid],
            Self::Hyperlink(..) => &[RevisionScope::Owner],
            Self::Run(..) | Self::Cell { .. } | Self::Section(_) => {
                &[RevisionScope::Owner, RevisionScope::Text]
            }
        }
    }
}

// Source keys include the complete physical route, not projected text ordinals.
type FormSourceKey = (
    rdocx::StoryKind,
    String,
    usize,
    rdocx::StoryItemKind,
    Vec<usize>,
);
type FormStoryKey = (rdocx::StoryKind, String, usize);

#[derive(Clone, PartialEq, Eq, Hash)]
enum FormOwnerRoute {
    Paragraph(FormSourceKey),
    Table(FormSourceKey),
    Cell(FormSourceKey, usize, usize, FormStoryKey),
    Section(usize),
}

#[derive(Default)]
struct FormSources {
    xml: HashMap<FormOwnerRoute, Option<Vec<u8>>>,
    cells: HashMap<FormOwnerRoute, usize>,
    cell_stories: HashMap<FormStoryKey, usize>,
}

fn form_source_key(location: &rdocx::ContentLocation) -> FormSourceKey {
    let story = location.story();
    (
        story.kind(),
        story.part_name().to_owned(),
        story.owner_index(),
        location.item_kind(),
        location.index_path().to_vec(),
    )
}

fn form_owner_route(location: &WordOwnerLocation) -> Option<FormOwnerRoute> {
    match location {
        WordOwnerLocation::Paragraph(location) => {
            Some(FormOwnerRoute::Paragraph(form_source_key(location)))
        }
        WordOwnerLocation::Table(location) => {
            Some(FormOwnerRoute::Table(form_source_key(location)))
        }
        WordOwnerLocation::Cell {
            table,
            row,
            cell,
            story,
        } => Some(FormOwnerRoute::Cell(
            form_source_key(table),
            *row,
            *cell,
            (
                story.kind(),
                story.part_name().to_owned(),
                story.owner_index(),
            ),
        )),
        WordOwnerLocation::Section(index) => Some(FormOwnerRoute::Section(*index)),
        WordOwnerLocation::Run(..) | WordOwnerLocation::Hyperlink(..) => None,
    }
}

fn word_table_read_facts(
    document: &rdocx::Document,
) -> HashMap<rdocx::ContentLocation, WordReadFacts> {
    // Optional acceleration only. The checked native reader remains the fallback
    // for source that the standalone checked table parser cannot project.
    document
        .standalone_story_item_snapshots()
        .unwrap_or_default()
        .into_iter()
        .filter(|item| item.location().item_kind() == rdocx::StoryItemKind::Table)
        .filter_map(|item| Some((item.location().clone(), word_table_facts(item.xml())?)))
        .collect()
}

/// The facade paragraph count of each modeled cell of a standalone table.
fn word_table_facts(xml: &[u8]) -> Option<WordReadFacts> {
    let table = rdocx_oxml::table::CT_Tbl::from_xml_fragment_checked(xml).ok()?;
    let rows = table
        .rows
        .iter()
        .map(|row| {
            row.cells
                .iter()
                .map(|cell| cell.paragraphs().len())
                .collect()
        })
        .collect();
    Some(WordReadFacts::Table(rows))
}

/// The four owner snapshot families of [`rdocx::Document::story_owner_snapshots`].
type WordOwnerSnapshots = (
    Vec<rdocx::StoryItemSnapshot>,
    Vec<(rdocx::ContentLocation, usize, usize, rdocx::StoryId)>,
    Vec<(rdocx::ContentLocation, Vec<u8>)>,
    Vec<(rdocx::ContentLocation, rdocx::LinkInfo)>,
);

/// A range paragraph's location, accepted text and accepted run paths.
type WordParagraphFacts = (
    rdocx::ContentLocation,
    String,
    Vec<rdocx_oxml::text::AcceptedRunPath>,
);

/// What the owner inventory reads of owner snapshots, each family in
/// snapshot order: every table, paragraph and content control item, with
/// its direct body index when it is a direct body child and a table's read
/// facts, the physical cells, the range paragraphs, and the item that owns
/// each hyperlink.
#[derive(Clone, Default)]
struct WordOwnerFacts {
    items: Vec<(rdocx::ContentLocation, Option<usize>, WordReadFacts)>,
    cells: Vec<(rdocx::ContentLocation, usize, usize, rdocx::StoryId)>,
    paragraphs: Vec<WordParagraphFacts>,
    links: Vec<rdocx::ContentLocation>,
}

/// The owner facts of `snapshots`, each table's read facts from
/// `table_facts`.
fn word_owner_facts(
    (snapshots, cells, paragraphs, links): WordOwnerSnapshots,
    table_facts: impl Fn(&rdocx::StoryItemSnapshot) -> WordReadFacts,
) -> rdocx::Result<WordOwnerFacts> {
    let items = snapshots
        .iter()
        .filter(|snapshot| {
            matches!(
                snapshot.location().item_kind(),
                rdocx::StoryItemKind::Table
                    | rdocx::StoryItemKind::Paragraph
                    | rdocx::StoryItemKind::ContentControl
            )
        })
        .map(|snapshot| {
            let read_facts = if snapshot.location().item_kind() == rdocx::StoryItemKind::Table {
                table_facts(snapshot)
            } else {
                WordReadFacts::None
            };
            let slot = snapshot
                .direct_body_index()
                .filter(|_| snapshot.is_direct_child());
            (snapshot.location().clone(), slot, read_facts)
        })
        .collect();
    let paragraphs = paragraphs
        .into_iter()
        .map(|(location, xml)| {
            let paragraph = rdocx_oxml::CT_P::from_xml_fragment(&xml)?;
            Ok((
                location,
                paragraph.accepted_text(),
                paragraph.accepted_run_paths(),
            ))
        })
        .collect::<rdocx::Result<_>>()?;
    Ok(WordOwnerFacts {
        items,
        cells,
        paragraphs,
        links: links.into_iter().map(|(location, _)| location).collect(),
    })
}

fn word_owner_locations(
    document: &rdocx::Document,
    mut form_sources: Option<&mut FormSources>,
) -> rdocx::Result<Vec<WordOwnerAddress>> {
    note_fx260_word(document.content_count().max(1));
    note_word_inventory_rebuild();
    let table_facts = word_table_read_facts(document);
    let snapshots = document.story_owner_snapshots()?;
    if let Some(sources) = form_sources.as_deref_mut() {
        record_form_sources(sources, &snapshots);
    }
    let facts = word_owner_facts(snapshots, |snapshot| {
        table_facts
            .get(snapshot.location())
            .cloned()
            .unwrap_or_default()
    })?;
    let sections = document.section_count();
    if let Some(sources) = form_sources {
        for index in 0..sections {
            sources.xml.insert(
                FormOwnerRoute::Section(index),
                Some(document.section_properties_xml(index)?),
            );
        }
    }
    Ok(word_owner_addresses(facts, sections))
}

/// Record what a form publication compares of `snapshots`: the XML of each
/// table and paragraph route, and how many cells share each cell route
/// and story.
fn record_form_sources(
    sources: &mut FormSources,
    (snapshots, cell_snapshots, paragraph_snapshots, _): &WordOwnerSnapshots,
) {
    for snapshot in snapshots {
        if snapshot.location().item_kind() == rdocx::StoryItemKind::Table {
            let key = FormOwnerRoute::Table(form_source_key(snapshot.location()));
            sources
                .xml
                .entry(key)
                .and_modify(|xml| *xml = None)
                .or_insert_with(|| Some(snapshot.xml().to_vec()));
        }
    }
    // Preserve raw multiplicity before the legacy cell-parent map collapses it.
    for (table, row, cell, story) in cell_snapshots {
        let story_key = (
            story.kind(),
            story.part_name().to_owned(),
            story.owner_index(),
        );
        *sources.cell_stories.entry(story_key.clone()).or_default() += 1;
        *sources
            .cells
            .entry(FormOwnerRoute::Cell(
                form_source_key(table),
                *row,
                *cell,
                story_key,
            ))
            .or_default() += 1;
    }
    for (location, xml) in paragraph_snapshots {
        let key = FormOwnerRoute::Paragraph(form_source_key(location));
        sources
            .xml
            .entry(key)
            .and_modify(|source| *source = None)
            .or_insert_with(|| Some(xml.clone()));
    }
}

/// The owner inventory of `facts` and of `sections` sections, in the order
/// that births and generations follow: tables, each followed by its cells,
/// then paragraphs, each followed by its runs, then sections, then
/// hyperlinks, each family in the order of `facts`.
fn word_owner_addresses(facts: WordOwnerFacts, sections: usize) -> Vec<WordOwnerAddress> {
    let WordOwnerFacts {
        items,
        cells,
        paragraphs,
        links,
    } = facts;
    let mut result = Vec::new();
    let mut table_ordinals = HashMap::new();
    let mut body_slots = HashMap::new();
    let mut cell_parents = HashMap::new();
    let mut table_cells = HashMap::<rdocx::ContentLocation, Vec<_>>::new();
    for (table, row, cell, story) in cells {
        table_cells
            .entry(table)
            .or_default()
            .push((row, cell, story));
    }
    for (location, slot, _) in &items {
        if location.story().kind() == rdocx::StoryKind::Body
            && matches!(
                location.item_kind(),
                rdocx::StoryItemKind::Paragraph | rdocx::StoryItemKind::ContentControl
            )
            && let Some(slot) = slot
        {
            body_slots.insert(location.clone(), *slot);
        }
    }
    for (location, slot, read_facts) in items {
        if location.item_kind() != rdocx::StoryItemKind::Table {
            continue;
        }
        let ordinal = table_ordinals
            .entry(location.story().clone())
            .or_insert(0usize);
        let projection_ordinal = Some(*ordinal);
        *ordinal += 1;
        let table_index = result.len();
        let cells = table_cells.remove(&location).unwrap_or_default();
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Table(location.clone()),
            parent: None,
            projection_ordinal,
            read_facts,
            body_slot: slot,
        });
        for (row, cell, story) in cells {
            result.push(WordOwnerAddress {
                location: WordOwnerLocation::Cell {
                    table: location.clone(),
                    row,
                    cell,
                    story,
                },
                parent: Some(table_index),
                body_slot: None,
                projection_ordinal: None,
                read_facts: WordReadFacts::None,
            });
        }
    }
    for (index, address) in result.iter().enumerate() {
        if let WordOwnerLocation::Cell { story, .. } = &address.location {
            cell_parents
                .entry((
                    story.kind(),
                    story.part_name().to_owned(),
                    story.owner_index(),
                ))
                .or_insert(index);
        }
    }
    // Join nested table routes through their actual physical cell owner. An
    // unresolved route stays conservative rather than acquiring an ordinal parent.
    for address in &mut result {
        if let WordOwnerLocation::Table(location) = &address.location {
            let story = location.story();
            let parent = cell_parents
                .get(&(
                    story.kind(),
                    story.part_name().to_owned(),
                    story.owner_index(),
                ))
                .copied();
            address.parent = parent;
        }
    }
    let mut paragraph_ordinals = HashMap::new();
    for (location, text, runs) in paragraphs {
        let story = location.story();
        let parent = cell_parents
            .get(&(
                story.kind(),
                story.part_name().to_owned(),
                story.owner_index(),
            ))
            .copied();
        let ordinal = paragraph_ordinals
            .entry(location.story().clone())
            .or_insert(0usize);
        let projection_ordinal = Some(*ordinal);
        *ordinal += 1;
        let index = result.len();
        // Reuse the containing item's fact from this exact immutable source
        // inventory, matching the native paragraph snapshot's containing route.
        let body_slot = if location.story().kind() == rdocx::StoryKind::Body {
            let containing = match location.index_path() {
                [_] => Some(location.clone()),
                [control, _] => Some(rdocx::ContentLocation::new(
                    location.story().clone(),
                    rdocx::StoryItemKind::ContentControl,
                    vec![*control],
                )),
                _ => None,
            };
            containing.and_then(|containing| body_slots.get(&containing).copied())
        } else {
            None
        };
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Paragraph(location.clone()),
            parent,
            body_slot,
            projection_ordinal,
            read_facts: WordReadFacts::Paragraph(text),
        });
        for path in runs {
            result.push(WordOwnerAddress {
                location: WordOwnerLocation::Run(location.clone(), path),
                parent: Some(index),
                body_slot: None,
                projection_ordinal: None,
                read_facts: WordReadFacts::None,
            });
        }
    }
    for index in 0..sections {
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Section(index),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        });
    }
    let mut local_indexes = HashMap::<rdocx::ContentLocation, usize>::new();
    let mut story_indexes = HashMap::<rdocx::StoryId, usize>::new();
    for location in links {
        let local = local_indexes.entry(location.clone()).or_default();
        let position = story_indexes.entry(location.story().clone()).or_default();
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Hyperlink(location.clone(), *local),
            parent: None,
            body_slot: None,
            projection_ordinal: Some(*position),
            read_facts: WordReadFacts::None,
        });
        *local += 1;
        *position += 1;
    }
    result
}

/// What the owner indexes read of the document besides the owners: the
/// story route of each body paragraph, the routes of the paragraphs of each
/// body table cell, and the direct body slot of each body table.
struct WordIndexProjections {
    body_paragraphs: Option<Vec<Option<rdocx::ContentLocation>>>,
    cell_paragraphs: CellParagraphRoutes,
    table_slots: Vec<Option<usize>>,
}

/// The routes of the paragraphs of each body table cell, by body table,
/// row and cell.
type CellParagraphRoutes = HashMap<(usize, usize, usize), Vec<rdocx::ContentLocation>>;

impl WordIndexProjections {
    fn of(document: &rdocx::Document) -> Self {
        Self {
            body_paragraphs: document.paragraph_story_locations().ok(),
            cell_paragraphs: document
                .cell_paragraph_story_locations()
                .unwrap_or_default(),
            table_slots: (0..document.table_count())
                .map(|index| document.content_index_of_table(index))
                .collect(),
        }
    }
}

/// The owners of one body block, read from a main part whose body holds
/// only that block: body routes count from the block's first story item
/// and the body story is that part's, table cell and text box stories
/// count from its first, and tables from its first.
struct WordBlockOwners {
    facts: WordOwnerFacts,
    /// The hyperlink relationships that its links name.
    relationships: Vec<String>,
    body_items: usize,
    cell_stories: usize,
    text_boxes: usize,
    /// The body story path of each of its paragraphs that
    /// `Document::paragraphs` counts.
    paragraphs: Vec<Option<Vec<usize>>>,
    tables: usize,
    /// Its part of `Document::cell_paragraph_story_locations`, and whether
    /// that routed every cell of the block, or `None` where it fails.
    cell_paragraphs: Option<(CellParagraphRoutes, bool)>,
    /// Whether the block holds a field, whose value compares equal to one
    /// that owns otherwise.
    holds_field: bool,
}

/// The owners of `block` as the only body child of `document`, or `None`
/// when they cannot take their place among other blocks by offsets: the
/// block does not read as one direct body item.
fn word_block_owners(
    document: &rdocx::Document,
    block: &rdocx_oxml::BodyContent,
) -> rdocx::Result<Option<WordBlockOwners>> {
    let rdocx::BodyItemOwners {
        snapshots,
        cells,
        paragraphs,
        links,
        cell_stories,
        text_boxes,
        paragraph_routes,
        tables,
        cell_paragraphs,
        holds_field,
    } = document.body_item_owners(block)?;
    let body = |snapshot: &&rdocx::StoryItemSnapshot| {
        snapshot.location().story().kind() == rdocx::StoryKind::Body
    };
    let mut direct = snapshots
        .iter()
        .filter(body)
        .filter(|snapshot| snapshot.is_direct_child())
        .map(rdocx::StoryItemSnapshot::direct_body_index);
    if direct.next() != Some(Some(0)) || direct.next().is_some() {
        return Ok(None);
    }
    let body_items = snapshots.iter().filter(body).count();
    let relationships = link_relationships(&links);
    let facts = word_owner_facts((snapshots, cells, paragraphs, links), |snapshot| {
        word_table_facts(snapshot.xml()).unwrap_or_default()
    })?;
    Ok(Some(WordBlockOwners {
        facts,
        relationships,
        body_items,
        cell_stories,
        text_boxes,
        paragraphs: paragraph_routes,
        tables,
        cell_paragraphs,
        holds_field,
    }))
}

/// The hyperlink relationships that `links` name, once each.
fn link_relationships(links: &[(rdocx::ContentLocation, rdocx::LinkInfo)]) -> Vec<String> {
    let mut relationships = links
        .iter()
        .filter_map(|(_, link)| link.rel_id.clone())
        .collect::<Vec<_>>();
    relationships.sort_unstable();
    relationships.dedup();
    relationships
}

/// The owners of a plain paragraph read as the only body child, from its
/// accepted text and run paths: one paragraph story item, and the body
/// story `body`, which a placed route replaces.
fn plain_paragraph_owners(
    body: &rdocx::StoryId,
    text: &str,
    runs: &[rdocx_oxml::text::AcceptedRunPath],
) -> WordBlockOwners {
    let location =
        rdocx::ContentLocation::new(body.clone(), rdocx::StoryItemKind::Paragraph, vec![0]);
    WordBlockOwners {
        facts: WordOwnerFacts {
            items: vec![(location.clone(), Some(0), WordReadFacts::None)],
            cells: Vec::new(),
            paragraphs: vec![(location, text.to_owned(), runs.to_vec())],
            links: Vec::new(),
        },
        relationships: Vec::new(),
        body_items: 1,
        cell_stories: 0,
        text_boxes: 0,
        paragraphs: vec![Some(vec![0])],
        tables: 0,
        cell_paragraphs: Some((HashMap::new(), true)),
        // A plain paragraph holds text runs only.
        holds_field: false,
    }
}

/// The owners of a detached table with `cells` cells, no nested table and
/// no link, read as the only body child, from its owner facts. Each cell's
/// paragraphs route when they are as many as its read facts count, as
/// `Document::cell_paragraph_story_locations` routes them.
fn table_owners(facts: WordOwnerFacts, cells: usize) -> Option<WordBlockOwners> {
    let Some((_, _, WordReadFacts::Table(rows))) = facts.items.first() else {
        return None;
    };
    let mut stories = HashMap::new();
    let mut routes = CellParagraphRoutes::new();
    for (_, row, cell, story) in &facts.cells {
        stories.insert(story.owner_index(), (*row, *cell));
        routes.insert((0, *row, *cell), Vec::new());
    }
    for (location, _, _) in &facts.paragraphs {
        if location.story().kind() == rdocx::StoryKind::TableCell
            && let Some((row, cell)) = stories.get(&location.story().owner_index())
        {
            routes
                .entry((0, *row, *cell))
                .or_default()
                .push(location.clone());
        }
    }
    routes.retain(|(_, row, cell), locations| {
        rows.get(*row).and_then(|row| row.get(*cell)) == Some(&locations.len())
    });
    Some(WordBlockOwners {
        relationships: Vec::new(),
        body_items: 1,
        cell_stories: cells,
        text_boxes: 0,
        paragraphs: Vec::new(),
        tables: 1,
        cell_paragraphs: Some((routes, true)),
        // A detached table is a new table, which holds no field.
        holds_field: false,
        facts,
    })
}

/// A route of a block's owners placed among the main part's blocks: body
/// story item paths move by `items` into `body`, and table cell and text
/// box stories by `stories`, the counts of each before the block.
fn placed_route(
    location: &rdocx::ContentLocation,
    body: &rdocx::StoryId,
    (items, stories): (usize, (usize, usize)),
) -> rdocx::ContentLocation {
    let story = location.story();
    if story.kind() == rdocx::StoryKind::Body {
        let mut placed = rdocx::ContentLocation::new(
            body.clone(),
            location.item_kind(),
            location.index_path().to_vec(),
        );
        if let Some(head) = location.index_path().first() {
            placed.set_index_path_head(head + items);
        }
        placed
    } else {
        rdocx::ContentLocation::new(
            placed_story(story, stories),
            location.item_kind(),
            location.index_path().to_vec(),
        )
    }
}

/// A table cell or text box story of a block's owners placed after the
/// `cells` table cell and `boxes` text box stories before the block.
fn placed_story(story: &rdocx::StoryId, (cells, boxes): (usize, usize)) -> rdocx::StoryId {
    let mut placed = story.clone();
    let before = if story.kind() == rdocx::StoryKind::TextBox {
        boxes
    } else {
        cells
    };
    placed.set_owner_index(story.owner_index() + before);
    placed
}

/// One body block and its owners, which hold while the document has an
/// equal block under an equal [`rdocx::BodyItemContext`].
struct WordBlockPiece {
    block: rdocx_oxml::BodyContent,
    /// `None` when the owners cannot be read or placed by offsets.
    owners: Option<WordBlockOwners>,
}

/// One story part after the main part and its owners, which hold while
/// its bytes do.
struct WordPartPiece {
    kind: rdocx::StoryKind,
    part_name: String,
    xml: Vec<u8>,
    /// The owner facts and the hyperlink relationships their links name,
    /// or `None` when they cannot be read.
    owners: Option<(WordOwnerFacts, Vec<String>)>,
}

/// The owner inventory of a document by body block and story part, each
/// piece kept with the value it was read from. A piece is reused only
/// where the document still holds an equal value: the same bytes for a
/// part, and for a block an equal `BodyContent` under an equal
/// [`rdocx::BodyItemContext`]. That holds while equal blocks own the same,
/// which the manual `PartialEq` of `CT_TabStop` and `BookmarkMarker` must
/// keep, since each ignores a field: the first one that serializes, the
/// last a read projection. A block holding a field is read again whenever
/// it is assembled, since `Field` compares equal to a field whose source
/// form or run span differs, which the block's bytes and owners follow.
#[derive(Clone, Default)]
pub(crate) struct WordPieces {
    context: Option<rdocx::BodyItemContext>,
    blocks: Vec<std::sync::Arc<WordBlockPiece>>,
    parts: Vec<std::sync::Arc<WordPartPiece>>,
}

impl WordPieces {
    /// The pieces of `document`: those of `self` whose block or part it
    /// still holds, and the others read again. With them, the inventory
    /// and index projections that `word_owner_locations` and
    /// [`WordIndexProjections::of`] give, or `None` where the pieces
    /// cannot give them and the rebuild must.
    fn assemble(
        &self,
        document: &rdocx::Document,
    ) -> (Self, Option<(Vec<WordOwnerAddress>, WordIndexProjections)>) {
        let context = document.body_item_context();
        let blocks = document.body_children();
        // A part read as two sources admits its owners once, which a part
        // read alone does not know, so the rebuild reads that inventory,
        // and nothing is read for the pieces.
        let Ok(sources) = document.story_part_sources() else {
            return (self.clone(), None);
        };
        let mut names = std::collections::HashSet::from([context.part_name()]);
        if !sources.iter().all(|(_, name, _)| names.insert(name)) {
            return (self.clone(), None);
        }
        // Equal blocks keep their pieces: the common prefix and suffix of
        // the body. Equality is that of `BodyContent`, which reads the same
        // owners only while its manual `PartialEq` impls ignore no field
        // that changes them, so a block holding a field is read again, in
        // its place (see `WordPieces`). Another root or package namespace
        // fact reads them all.
        let old = if self.context.as_ref() == Some(&context) {
            self.blocks.as_slice()
        } else {
            &[]
        };
        let prefix = old
            .iter()
            .zip(blocks)
            .take_while(|(piece, block)| piece.block == **block)
            .count();
        let suffix = old[prefix..]
            .iter()
            .rev()
            .zip(blocks[prefix..].iter().rev())
            .take_while(|(piece, block)| piece.block == **block)
            .count();
        let mut pieces = Self {
            context: Some(context),
            blocks: Vec::with_capacity(blocks.len()),
            parts: Vec::new(),
        };
        let kept = old[..prefix]
            .iter()
            .map(Some)
            .chain(std::iter::repeat_n(None, blocks.len() - prefix - suffix))
            .chain(old[old.len() - suffix..].iter().map(Some));
        for (block, piece) in blocks.iter().zip(kept) {
            let holds_field = |piece: &&std::sync::Arc<WordBlockPiece>| {
                piece
                    .owners
                    .as_ref()
                    .is_some_and(|owners| owners.holds_field)
            };
            match piece.filter(|piece| !holds_field(piece)) {
                Some(piece) => pieces.blocks.push(piece.clone()),
                None => {
                    note_word_piece_read();
                    pieces.blocks.push(std::sync::Arc::new(WordBlockPiece {
                        block: block.clone(),
                        owners: word_block_owners(document, block).ok().flatten(),
                    }));
                }
            }
        }
        for (kind, part_name, xml) in sources {
            if let Some(piece) = self.parts.iter().find(|piece| {
                piece.kind == kind && piece.part_name == part_name && piece.xml == *xml
            }) {
                pieces.parts.push(piece.clone());
                continue;
            }
            note_word_piece_read();
            let owners = document
                .story_part_owner_snapshots(kind, &part_name, &xml)
                .and_then(|snapshots| {
                    let relationships = link_relationships(&snapshots.3);
                    let facts = word_owner_facts(snapshots, |snapshot| {
                        word_table_facts(snapshot.xml()).unwrap_or_default()
                    })?;
                    Ok((facts, relationships))
                })
                .ok();
            pieces.parts.push(std::sync::Arc::new(WordPartPiece {
                kind,
                part_name,
                xml: xml.into_owned(),
                owners,
            }));
        }
        let inventory = pieces.inventory(document);
        (pieces, inventory)
    }

    /// Keep the pieces in step with a body splice made in place in
    /// `document`: the pieces of the `removed` blocks at `slot` go, and the
    /// owners `inserted` of the block now at `slot` take their place. Pieces
    /// that did not count the body's blocks before the splice stay as they
    /// were. The next inventory compares every piece with its block by value
    /// anyway.
    fn splice(
        &mut self,
        document: &rdocx::Document,
        slot: usize,
        removed: usize,
        inserted: Option<WordBlockOwners>,
    ) {
        let blocks = document.body_children();
        let added = usize::from(inserted.is_some());
        if self.blocks.len() + added != blocks.len() + removed || slot + removed > self.blocks.len()
        {
            return;
        }
        let piece = inserted.map(|owners| {
            std::sync::Arc::new(WordBlockPiece {
                block: blocks[slot].clone(),
                owners: Some(owners),
            })
        });
        self.blocks.splice(slot..slot + removed, piece);
    }

    /// The inventory and index projections of these pieces, which are those
    /// of `document`, placed by offsets.
    fn inventory(
        &self,
        document: &rdocx::Document,
    ) -> Option<(Vec<WordOwnerAddress>, WordIndexProjections)> {
        let main = self.context.as_ref()?.part_name();
        let body = rdocx::StoryId::new(rdocx::StoryKind::Body, main, 0);
        // The links of every piece name relationships that must still be
        // external hyperlinks of its part, as the rebuild checks them.
        let mut links = None;
        // Each owner family of a story source lists the body story first,
        // then the table cell and text box stories in source order.
        let mut facts = WordOwnerFacts::default();
        let mut cells = WordOwnerFacts::default();
        let mut projections = WordIndexProjections {
            body_paragraphs: Some(Vec::new()),
            cell_paragraphs: HashMap::new(),
            table_slots: Vec::new(),
        };
        let mut cell_paragraphs = Some(HashMap::new());
        let mut routed = true;
        let (mut items, mut stories, mut tables) = (0, (0, 0), 0);
        for (slot, piece) in self.blocks.iter().enumerate() {
            let owners = piece.owners.as_ref()?;
            if !owners.relationships.is_empty() {
                let valid =
                    links.get_or_insert_with(|| document.story_hyperlink_relationships(main));
                if !owners
                    .relationships
                    .iter()
                    .all(|relationship| valid.contains(relationship.as_str()))
                {
                    return None;
                }
            }
            let place = |location| placed_route(location, &body, (items, stories));
            let in_body = |location: &rdocx::ContentLocation| {
                location.story().kind() == rdocx::StoryKind::Body
            };
            for (location, direct, read_facts) in &owners.facts.items {
                let family = if in_body(location) {
                    &mut facts.items
                } else {
                    &mut cells.items
                };
                family.push((place(location), direct.map(|_| slot), read_facts.clone()));
            }
            for (table, row, cell, story) in &owners.facts.cells {
                facts
                    .cells
                    .push((place(table), *row, *cell, placed_story(story, stories)));
            }
            for (location, text, runs) in &owners.facts.paragraphs {
                let family = if in_body(location) {
                    &mut facts.paragraphs
                } else {
                    &mut cells.paragraphs
                };
                family.push((place(location), text.clone(), runs.clone()));
            }
            for location in &owners.facts.links {
                let family = if in_body(location) {
                    &mut facts.links
                } else {
                    &mut cells.links
                };
                family.push(place(location));
            }
            if let Some(paragraphs) = &mut projections.body_paragraphs {
                paragraphs.extend(owners.paragraphs.iter().map(|path| {
                    path.as_ref().map(|path| {
                        let mut location = rdocx::ContentLocation::new(
                            body.clone(),
                            rdocx::StoryItemKind::Paragraph,
                            path.clone(),
                        );
                        location.set_index_path_head(path[0] + items);
                        location
                    })
                }));
            }
            if matches!(piece.block, rdocx_oxml::BodyContent::Table(_)) {
                projections.table_slots.push(Some(slot));
            } else {
                projections
                    .table_slots
                    .extend(std::iter::repeat_n(None, owners.tables));
            }
            // The whole inventory stops at the first cell it cannot route,
            // and fails as a whole where any block it reaches fails.
            if routed {
                match &owners.cell_paragraphs {
                    Some((routes, complete)) => {
                        if let Some(all) = &mut cell_paragraphs {
                            for ((table, row, cell), locations) in routes {
                                all.insert(
                                    (table + tables, *row, *cell),
                                    locations
                                        .iter()
                                        .map(|location| {
                                            placed_route(location, &body, (items, stories))
                                        })
                                        .collect(),
                                );
                            }
                        }
                        routed = *complete;
                    }
                    None => {
                        cell_paragraphs = None;
                        routed = false;
                    }
                }
            }
            items += owners.body_items;
            stories.0 += owners.cell_stories;
            stories.1 += owners.text_boxes;
            tables += owners.tables;
        }
        projections.cell_paragraphs = cell_paragraphs.unwrap_or_default();
        facts.items.append(&mut cells.items);
        facts.paragraphs.append(&mut cells.paragraphs);
        facts.links.append(&mut cells.links);
        for piece in &self.parts {
            let (owners, relationships) = piece.owners.as_ref()?;
            if !relationships.is_empty() {
                let valid = document.story_hyperlink_relationships(&piece.part_name);
                if !relationships
                    .iter()
                    .all(|relationship| valid.contains(relationship.as_str()))
                {
                    return None;
                }
            }
            facts.items.extend_from_slice(&owners.items);
            facts.cells.extend_from_slice(&owners.cells);
            facts.paragraphs.extend_from_slice(&owners.paragraphs);
            facts.links.extend_from_slice(&owners.links);
        }
        Some((
            word_owner_addresses(facts, document.section_count()),
            projections,
        ))
    }
}

/// One direct body item that a splice inserts, with what the owner
/// inventory reads for it.
pub(crate) enum SplicedBodyItem {
    /// A plain paragraph's accepted text and run paths.
    Paragraph(String, Vec<rdocx_oxml::text::AcceptedRunPath>),
    /// A table's namespace-complete XML and its owner snapshots in a body
    /// that holds only that table.
    Table(Vec<u8>, WordOwnerSnapshots),
}

/// Plain direct body paragraphs removed from a body slot, or one item
/// inserted at it.
pub(crate) struct BodySplice {
    pub(crate) slot: usize,
    pub(crate) removed: usize,
    pub(crate) inserted: Option<SplicedBodyItem>,
}

/// A splice of the current owner inventory for one direct body edit.
///
/// The inserted item is one body story item, a plain paragraph or a table
/// with no nested table, text box or link, and the removed items are plain
/// paragraphs. So only body routes after the slot move, by the items inserted
/// less those removed, and only the cell stories after an inserted table,
/// by its cells. Every other owner keeps its place in the inventory.
pub(crate) struct BodySplicePlan {
    slot: usize,
    removed: usize,
    // The body story item index of the splice slot.
    first: usize,
    // The inserted owners, in blocks that each precede a current inventory
    // index, with next-inventory parents.
    blocks: Vec<(usize, Vec<WordOwnerAddress>)>,
    // The current inventory range of the removed paragraphs and their runs.
    removed_owners: std::ops::Range<usize>,
    // Cell story owners from the first index on move by the second.
    cells: (usize, usize),
    inserted_paragraph: bool,
    inserted_table: bool,
    // The facade paragraph and table indexes of the splice slot.
    pub(crate) paragraph: usize,
    pub(crate) table: usize,
    // The inserted table's route and the routes of its cell paragraphs.
    table_entry: Option<(rdocx::ContentLocation, Vec<CellParagraphs>)>,
    // The owners of the inserted item alone, for the inventory by block.
    piece: Option<WordBlockOwners>,
}

/// The facade coordinates of a cell and the routes of its paragraphs.
type CellParagraphs = ((usize, usize), Vec<rdocx::ContentLocation>);

impl BodySplicePlan {
    fn added(&self) -> usize {
        self.blocks.iter().map(|(_, block)| block.len()).sum()
    }

    /// The body story item index that `item` moves to.
    fn moved(&self, item: usize) -> usize {
        if item >= self.first + self.removed {
            item + usize::from(self.inserted_paragraph || self.inserted_table) - self.removed
        } else {
            item
        }
    }

    /// The next inventory index of current owner `index`, unless removed.
    fn next_index(&self, index: usize) -> Option<usize> {
        if self.removed_owners.contains(&index) {
            return None;
        }
        let mut next = index;
        if index >= self.removed_owners.end {
            next -= self.removed_owners.len();
        }
        for (at, block) in &self.blocks {
            if index >= *at {
                next += block.len();
            }
        }
        Some(next)
    }

    /// The next inventory, in order: kept current owners and inserted ones.
    fn next_owners(&self, current: usize) -> impl Iterator<Item = NextWordOwner<'_>> + '_ {
        (0..=current).flat_map(move |index| {
            self.blocks
                .iter()
                .filter(move |(at, _)| *at == index)
                .flat_map(|(_, block)| block.iter().map(NextWordOwner::Inserted))
                .chain(
                    (index < current && !self.removed_owners.contains(&index))
                        .then_some(NextWordOwner::Kept(index)),
                )
        })
    }

    /// Move the cell stories of a kept owner's routes.
    fn move_cell_stories(&self, location: &mut WordOwnerLocation, part: &str) {
        let (first, count) = self.cells;
        if count == 0 {
            return;
        }
        let route = |location: &mut rdocx::ContentLocation| {
            let story = location.story();
            if story.kind() == rdocx::StoryKind::TableCell
                && story.part_name() == part
                && story.owner_index() >= first
            {
                let mut story = story.clone();
                story.set_owner_index(story.owner_index() + count);
                *location = rdocx::ContentLocation::new(
                    story,
                    location.item_kind(),
                    location.index_path().to_vec(),
                );
            }
        };
        match location {
            WordOwnerLocation::Paragraph(location)
            | WordOwnerLocation::Table(location)
            | WordOwnerLocation::Run(location, _)
            | WordOwnerLocation::Hyperlink(location, _) => route(location),
            WordOwnerLocation::Cell { table, story, .. } => {
                route(table);
                if story.kind() == rdocx::StoryKind::TableCell
                    && story.part_name() == part
                    && story.owner_index() >= first
                {
                    story.set_owner_index(story.owner_index() + count);
                }
            }
            WordOwnerLocation::Section(_) => {}
        }
    }
}

/// One owner of the next inventory.
enum NextWordOwner<'a> {
    Kept(usize),
    Inserted(&'a WordOwnerAddress),
}

/// A prepared body splice: its revision and the scope states of the next
/// inventory, in order.
pub(crate) struct PreparedBodySplice {
    pub(crate) revision: PreparedRevision,
    pub(crate) plan: BodySplicePlan,
    states: Vec<RevisionState>,
    // The index in `states` of the first inserted owner's scopes.
    inserted_state: usize,
    snapshot_scope: RevisionState,
}

/// The body story item index of an owner's content route.
fn word_owner_body_item(location: &WordOwnerLocation) -> Option<usize> {
    let route = match location {
        WordOwnerLocation::Paragraph(location)
        | WordOwnerLocation::Table(location)
        | WordOwnerLocation::Run(location, _)
        | WordOwnerLocation::Hyperlink(location, _) => location,
        WordOwnerLocation::Cell { table, .. } => table,
        WordOwnerLocation::Section(_) => return None,
    };
    (route.story().kind() == rdocx::StoryKind::Body)
        .then(|| route.index_path().first().copied())
        .flatten()
}

/// Move the body story item index of an owner's content route.
fn move_word_owner_body_item(location: &mut WordOwnerLocation, item: usize) {
    match location {
        WordOwnerLocation::Paragraph(location)
        | WordOwnerLocation::Table(location)
        | WordOwnerLocation::Run(location, _)
        | WordOwnerLocation::Hyperlink(location, _)
        | WordOwnerLocation::Cell {
            table: location, ..
        } => location.set_index_path_head(item),
        WordOwnerLocation::Section(_) => {}
    }
}

/// The story of an owner's content route.
fn word_owner_story(location: &WordOwnerLocation) -> Option<&rdocx::StoryId> {
    match location {
        WordOwnerLocation::Paragraph(location)
        | WordOwnerLocation::Table(location)
        | WordOwnerLocation::Run(location, _)
        | WordOwnerLocation::Hyperlink(location, _) => Some(location.story()),
        WordOwnerLocation::Cell { table, .. } => Some(table.story()),
        WordOwnerLocation::Section(_) => None,
    }
}

/// Whether `owners` has the order a rebuild gives it: tables each followed
/// by its cells, paragraphs each followed by its runs with the body story
/// first and in source order, then sections, then hyperlinks. A local edit
/// appends the runs it creates, which a later rebuild puts back in place.
fn word_inventory_is_canonical(owners: &[WordOwner]) -> bool {
    let mut family = 0;
    let mut table = None;
    let mut paragraph = None;
    let mut body_item = Some(0);
    for (index, owner) in owners.iter().enumerate() {
        let (next, parent) = match &owner.location {
            WordOwnerLocation::Table(_) => {
                table = Some(index);
                (0, owner.parent)
            }
            WordOwnerLocation::Cell { .. } => (0, table),
            WordOwnerLocation::Paragraph(location) => {
                paragraph = Some(index);
                if location.story().kind() == rdocx::StoryKind::Body {
                    let item = location.index_path().first().copied();
                    if body_item.is_none() || item < body_item {
                        return false;
                    }
                    body_item = item;
                } else {
                    body_item = None;
                }
                (1, owner.parent)
            }
            WordOwnerLocation::Run(..) => (1, paragraph),
            WordOwnerLocation::Section(_) => (2, None),
            WordOwnerLocation::Hyperlink(..) => (3, None),
        };
        let child = matches!(
            owner.location,
            WordOwnerLocation::Cell { .. } | WordOwnerLocation::Run(..)
        );
        if next < family || (child && (parent.is_none() || owner.parent != parent)) {
            return false;
        }
        family = next;
    }
    true
}

impl PyDocument {
    /// Plan `splice` from the current inventory alone. `None` keeps the
    /// rebuild for every case this does not prove, and debug builds check
    /// each committed splice against the rebuild.
    pub(crate) fn plan_body_splice(&self, splice: BodySplice) -> PyResult<Option<BodySplicePlan>> {
        let Some(body_paragraphs) = &self.indexes.body_paragraphs else {
            return Ok(None);
        };
        let content_count = self.inner.content_count();
        let Some(end_slot) = splice.slot.checked_add(splice.removed) else {
            return Ok(None);
        };
        let inserted_table = matches!(splice.inserted, Some(SplicedBodyItem::Table(..)));
        if end_slot > content_count
            || (splice.removed == 0) == splice.inserted.is_none()
            || (inserted_table && self.indexes.body_tables.iter().any(Option::is_none))
            || !(splice.slot..end_slot).all(|slot| self.inner.plain_removable_paragraph(slot))
            || !word_inventory_is_canonical(&self.owners)
        {
            return Ok(None);
        }
        note_fx260_word(1);
        // The body story, and the item index of every owned direct body slot.
        let mut story = None;
        let mut slot_items = HashMap::new();
        for owner in &self.owners {
            let (WordOwnerLocation::Paragraph(location) | WordOwnerLocation::Table(location)) =
                &owner.location
            else {
                continue;
            };
            if location.story().kind() != rdocx::StoryKind::Body {
                continue;
            }
            story.get_or_insert_with(|| location.story().clone());
            if let (Some(slot), Some(item)) = (owner.body_slot, location.index_path().first()) {
                slot_items.entry(slot).or_insert(*item);
            }
        }
        let Some(story) = story else {
            return Ok(None);
        };
        let part = story.part_name().to_owned();
        // An insertion at the end, or before a slot that owns no item (a
        // content control whose paragraphs no route reaches), starts where
        // the previous slot's items end.
        let first = if let Some(item) = slot_items.get(&splice.slot) {
            *item
        } else if splice.removed == 0 && splice.slot > 0 {
            let Some(previous) = slot_items.get(&(splice.slot - 1)) else {
                return Ok(None);
            };
            let Some(count) = self
                .inner
                .direct_body_story_item_count(splice.slot - 1)
                .ok()
                .flatten()
            else {
                return Ok(None);
            };
            previous + count
        } else {
            return Ok(None);
        };
        // Each removed slot is the direct paragraph owner at its own item.
        for offset in 0..splice.removed {
            let Some(owner) = self
                .indexes
                .direct_body_paragraphs
                .get(&(splice.slot + offset))
                .and_then(|index| self.owners.get(*index))
            else {
                return Ok(None);
            };
            match &owner.location {
                WordOwnerLocation::Paragraph(location)
                    if location.index_path() == [first + offset] => {}
                _ => return Ok(None),
            }
        }
        let Some(cells_before) = (if inserted_table {
            self.cell_stories_before(first, &part)
        } else {
            Some(0)
        }) else {
            return Ok(None);
        };
        let removed_items = first..first + splice.removed;
        let mut removed_owners = None::<std::ops::Range<usize>>;
        let mut paragraph_at = None;
        let mut table_at = None;
        let mut cells_at = None;
        let mut paragraph_ordinal = 0;
        let mut table_ordinal = 0;
        for (index, owner) in self.owners.iter().enumerate() {
            note_fx260_word(1);
            let item = word_owner_body_item(&owner.location);
            let story = word_owner_story(&owner.location);
            let body = story.is_some_and(|story| story.kind() == rdocx::StoryKind::Body);
            let (paragraph, table) = match &owner.location {
                WordOwnerLocation::Paragraph(_) => (true, false),
                WordOwnerLocation::Table(_) => (false, true),
                _ => (false, false),
            };
            let later = matches!(
                owner.location,
                WordOwnerLocation::Section(_) | WordOwnerLocation::Hyperlink(..)
            );
            if paragraph_at.is_none() && (later || (paragraph && !(body && item < Some(first)))) {
                paragraph_at = Some(index);
            }
            if table_at.is_none()
                && (later
                    || paragraph
                    || matches!(owner.location, WordOwnerLocation::Run(..))
                    || (table && !(body && item < Some(first))))
            {
                table_at = Some(index);
            }
            let later_cell_story = story.is_some_and(|story| {
                story.part_name() != part
                    || (story.kind() == rdocx::StoryKind::TableCell
                        && story.owner_index() >= cells_before)
            });
            if cells_at.is_none() && (later || (paragraph && later_cell_story)) {
                cells_at = Some(index);
            }
            if item.is_some_and(|item| removed_items.contains(&item)) {
                // Only the removed paragraphs and their runs, contiguously.
                let contiguous = removed_owners
                    .as_ref()
                    .is_none_or(|range| range.end == index);
                if !contiguous
                    || !matches!(
                        owner.location,
                        WordOwnerLocation::Paragraph(_) | WordOwnerLocation::Run(..)
                    )
                {
                    return Ok(None);
                }
                removed_owners =
                    Some(removed_owners.map_or(index..index + 1, |range| range.start..index + 1));
                continue;
            }
            if owner.parent.is_some_and(|parent| {
                removed_owners
                    .as_ref()
                    .is_some_and(|range| range.contains(&parent))
            }) {
                return Ok(None);
            }
            if body && item < Some(first) {
                paragraph_ordinal += usize::from(paragraph);
                table_ordinal += usize::from(table);
            }
        }
        if splice.removed > 0 && removed_owners.is_none() {
            return Ok(None);
        }
        let current = self.owners.len();
        // Facade paragraphs and body tables keep source order, so the splice
        // is one contiguous run of their entries. A paragraph that a nested
        // content control holds has no route of its own, and stays in place.
        let head = |location: &Option<rdocx::ContentLocation>| {
            location
                .as_ref()
                .and_then(|location| location.index_path().first().copied())
        };
        let paragraph = self.inner.paragraph_count_before_content(splice.slot);
        let Some(after) = body_paragraphs.get(paragraph..) else {
            return Ok(None);
        };
        for (offset, location) in after.iter().enumerate() {
            let fits = if offset < splice.removed {
                location.as_ref().map(rdocx::ContentLocation::index_path) == Some(&[first + offset])
            } else {
                location.is_none() || head(location) >= Some(removed_items.end)
            };
            if !fits {
                return Ok(None);
            }
        }
        if body_paragraphs[..paragraph]
            .iter()
            .any(|location| location.is_some() && head(location) >= Some(first))
        {
            return Ok(None);
        }
        let table = self
            .indexes
            .body_tables
            .partition_point(|location| head(location) < Some(first));
        if self.indexes.body_tables.iter().any(|location| {
            location.as_ref().is_some_and(|location| {
                location.story().kind() == rdocx::StoryKind::Body
                    && removed_items.contains(&location.index_path()[0])
            })
        }) {
            return Ok(None);
        }
        let mut plan = BodySplicePlan {
            slot: splice.slot,
            removed: splice.removed,
            first,
            blocks: Vec::new(),
            removed_owners: removed_owners.unwrap_or(0..0),
            cells: (cells_before, 0),
            inserted_paragraph: false,
            inserted_table: false,
            paragraph,
            table,
            table_entry: None,
            piece: None,
        };
        match splice.inserted {
            None => {}
            Some(SplicedBodyItem::Paragraph(text, runs)) => {
                plan.piece = Some(plain_paragraph_owners(&story, &text, &runs));
                let at = paragraph_at.unwrap_or(current);
                let location = rdocx::ContentLocation::new(
                    story,
                    rdocx::StoryItemKind::Paragraph,
                    vec![first],
                );
                let mut block = vec![WordOwnerAddress {
                    location: WordOwnerLocation::Paragraph(location.clone()),
                    parent: None,
                    body_slot: Some(splice.slot),
                    projection_ordinal: Some(paragraph_ordinal),
                    read_facts: WordReadFacts::Paragraph(text),
                }];
                block.extend(runs.into_iter().map(|path| WordOwnerAddress {
                    location: WordOwnerLocation::Run(location.clone(), path),
                    parent: Some(at),
                    body_slot: None,
                    projection_ordinal: None,
                    read_facts: WordReadFacts::None,
                }));
                plan.blocks.push((at, block));
                plan.inserted_paragraph = true;
            }
            Some(SplicedBodyItem::Table(xml, snapshots)) => {
                let table_at = table_at.unwrap_or(current);
                let cells_at = cells_at.unwrap_or(current);
                let Some((blocks, cells, entry, piece)) = spliced_table_blocks(
                    xml,
                    snapshots,
                    &story,
                    (splice.slot, first, table_ordinal),
                    (table_at, cells_at, cells_before),
                ) else {
                    return Ok(None);
                };
                plan.blocks = blocks;
                plan.cells = (cells_before, cells);
                plan.inserted_table = true;
                plan.table_entry = Some(entry);
                plan.piece = piece;
            }
        }
        Ok(Some(plan))
    }

    /// The number of cell stories of the main part before body story item
    /// `first`, when the inventory holds every cell story of that part and
    /// can place each one: no text box story, and every cell under a body
    /// table, so that cell story owner indexes count exactly those cells.
    fn cell_stories_before(&self, first: usize, part: &str) -> Option<usize> {
        let mut cells = Vec::new();
        let mut referenced = std::collections::HashSet::new();
        for owner in &self.owners {
            note_fx260_word(1);
            if let Some(story) = word_owner_story(&owner.location)
                && story.part_name() == part
            {
                match story.kind() {
                    rdocx::StoryKind::TextBox => return None,
                    rdocx::StoryKind::TableCell => {
                        referenced.insert(story.owner_index());
                    }
                    _ => {}
                }
            }
            let WordOwnerLocation::Cell { story, .. } = &owner.location else {
                continue;
            };
            if story.part_name() != part {
                continue;
            }
            // The body table that holds this cell, through nested tables.
            let mut table = owner.parent;
            let root = loop {
                let holder = self.owners.get(table?)?;
                match &holder.location {
                    WordOwnerLocation::Table(location)
                        if location.story().kind() == rdocx::StoryKind::Body =>
                    {
                        break location.index_path().first().copied()?;
                    }
                    WordOwnerLocation::Table(_) => {
                        table = self.owners.get(holder.parent?)?.parent;
                    }
                    _ => return None,
                }
            };
            cells.push((story.owner_index(), root));
        }
        cells.sort_unstable();
        let complete = cells
            .iter()
            .enumerate()
            .all(|(index, (owner_index, _))| index == *owner_index)
            && referenced
                .iter()
                .all(|owner_index| *owner_index < cells.len());
        let before = cells.iter().filter(|(_, root)| *root < first).count();
        (complete
            && cells
                .iter()
                .all(|(owner_index, root)| (*root < first) == (*owner_index < before)))
        .then_some(before)
    }

    /// Prepare a conservative or append publication of `plan`, with the
    /// retirements and births, in inventory order, that the rebuild would
    /// give, before the native edit.
    pub(crate) fn prepare_body_splice(
        &mut self,
        plan: BodySplicePlan,
        operation: &'static str,
        publication: &WordPublication,
    ) -> PyResult<PreparedBodySplice> {
        // An append keeps every owner. A conservative publication retires
        // every owner and gives the next inventory new scopes.
        let keep = match publication {
            WordPublication::AppendBody => true,
            WordPublication::Conservative => false,
            _ => return Err(PyValueError::new_err("body splice publication")),
        };
        let fail =
            |_| pyo3::exceptions::PyMemoryError::new_err("Word owner splice allocation failed");
        let current = self.owners.len();
        let scopes = |next: &NextWordOwner<'_>| match next {
            NextWordOwner::Kept(index) => self.owners[*index].location.scopes().len(),
            NextWordOwner::Inserted(address) => address.location.scopes().len(),
        };
        let kept_scopes = self
            .owners
            .iter()
            .map(|owner| owner.location.scopes().len())
            .sum::<usize>();
        let inserted_scopes = plan
            .blocks
            .iter()
            .flat_map(|(_, block)| block)
            .map(|address| address.location.scopes().len())
            .sum::<usize>();
        let mut births = Vec::new();
        births
            .try_reserve(inserted_scopes + if keep { 0 } else { kept_scopes } + 1)
            .map_err(fail)?;
        let mut retired = Vec::new();
        retired
            .try_reserve(if keep { 1 } else { kept_scopes + 1 })
            .map_err(fail)?;
        if !keep {
            retired.extend(
                self.owners
                    .iter()
                    .flat_map(WordOwner::scope_states)
                    .cloned(),
            );
        }
        for next in plan.next_owners(current) {
            note_fx260_word(1);
            match next {
                NextWordOwner::Kept(_) if keep => {}
                NextWordOwner::Kept(index) => {
                    births.extend_from_slice(self.owners[index].location.scopes());
                }
                NextWordOwner::Inserted(address) => {
                    births.extend_from_slice(address.location.scopes());
                }
            }
        }
        retired.push(self.snapshot_scope.clone());
        births.push(RevisionScope::Text);
        let revision = self
            .revisions
            .prepare_scopes(operation, &retired, &births)
            .map_err(revision_to_pyerr)?;
        let mut states = Vec::new();
        states
            .try_reserve_exact(inserted_scopes + kept_scopes)
            .map_err(fail)?;
        let mut born = 0;
        let mut inserted_state = None;
        for next in plan.next_owners(current) {
            let count = scopes(&next);
            match next {
                NextWordOwner::Kept(index) if keep => {
                    states.extend(self.owners[index].scope_states().cloned());
                    continue;
                }
                NextWordOwner::Inserted(_) => {
                    inserted_state.get_or_insert(states.len());
                }
                NextWordOwner::Kept(_) => {}
            }
            for _ in 0..count {
                states.push(revision.born(born).expect("prepared owner birth"));
                born += 1;
            }
        }
        let snapshot_scope = revision.born(born).expect("prepared snapshot birth");
        // Capacity for the commit, which moves owners and indexes them again.
        let added = plan.added();
        self.owners.try_reserve(added).map_err(fail)?;
        if let Some(paragraphs) = &mut self.indexes.body_paragraphs {
            paragraphs.try_reserve(1).map_err(fail)?;
        }
        self.indexes.body_tables.try_reserve(1).map_err(fail)?;
        Ok(PreparedBodySplice {
            revision,
            plan,
            states,
            inserted_state: inserted_state.unwrap_or(0),
            snapshot_scope,
        })
    }

    /// Capture the first owner a prepared splice inserts, the paragraph or
    /// table it places, as [`Self::capture_prepared_owner`] captures a
    /// prepared owner.
    pub(crate) fn capture_prepared_splice_owner(
        &self,
        prepared: &PreparedBodySplice,
        segments: smallvec::SmallVec<[PathSeg; 5]>,
    ) -> PyResult<(ContentPath, RevisionState)> {
        self.revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let unbound = || PyValueError::new_err("element has no current checked physical owner");
        let address = prepared
            .plan
            .blocks
            .first()
            .and_then(|(_, block)| block.first())
            .filter(|address| address.parent.is_none())
            .ok_or_else(unbound)?;
        let at = prepared.inserted_state;
        let states = prepared
            .states
            .get(at..at + address.location.scopes().len())
            .ok_or_else(unbound)?;
        let owners = word_owners_with_states(
            vec![WordOwnerAddress {
                location: address.location.clone(),
                parent: None,
                body_slot: address.body_slot,
                projection_ordinal: address.projection_ordinal,
                read_facts: WordReadFacts::None,
            }],
            states,
        );
        let (mut path, owner) = word_owner_capture(
            &self.revisions,
            &owners,
            (prepared.revision.document_scope(), &self.removal_scope),
            0,
            segments,
            WordView::Owner,
            None,
        )?;
        path.revision = prepared.revision.committed_serial();
        Ok((path, owner))
    }

    /// Publish a prepared splice after its native edit: move the kept owners
    /// in place, place the inserted ones, and index the inventory again.
    pub(crate) fn commit_body_splice(
        &mut self,
        prepared: PreparedBodySplice,
        publication: &WordPublication,
    ) -> PyResult<()> {
        #[cfg(debug_assertions)]
        if matches!(publication, WordPublication::AppendBody) {
            self.assert_body_append_matches(&prepared.plan, publication);
        }
        #[cfg(not(debug_assertions))]
        let _ = publication;
        let PreparedBodySplice {
            revision,
            mut plan,
            states,
            snapshot_scope,
            ..
        } = prepared;
        self.revisions
            .commit_scopes(revision)
            .map_err(revision_to_pyerr)?;
        let part = self
            .owners
            .iter()
            .find_map(|owner| {
                word_owner_story(&owner.location)
                    .filter(|story| story.kind() == rdocx::StoryKind::Body)
                    .map(|story| story.part_name().to_owned())
            })
            .unwrap_or_default();
        let current = self.owners.len();
        // The scope states of each inserted owner, in inventory order.
        let mut inserted = Vec::with_capacity(plan.added());
        let mut cursor = 0;
        for next in plan.next_owners(current) {
            let (count, address) = match next {
                NextWordOwner::Kept(index) => (self.owners[index].location.scopes().len(), None),
                NextWordOwner::Inserted(address) => {
                    (address.location.scopes().len(), Some(address))
                }
            };
            if let Some(address) = address {
                inserted.extend(word_owners_with_states(
                    vec![WordOwnerAddress {
                        location: address.location.clone(),
                        parent: address.parent,
                        body_slot: address.body_slot,
                        projection_ordinal: address.projection_ordinal,
                        read_facts: address.read_facts.clone(),
                    }],
                    &states[cursor..cursor + count],
                ));
            }
            cursor += count;
        }
        self.owners.drain(plan.removed_owners.clone());
        let mut inserted = inserted.into_iter();
        let mut spans = Vec::with_capacity(plan.blocks.len());
        let mut placed = 0;
        for (at, block) in &plan.blocks {
            let start = *at + placed;
            self.owners
                .splice(start..start, inserted.by_ref().take(block.len()));
            spans.push(start..start + block.len());
            placed += block.len();
        }
        let after = plan.first + plan.removed;
        let paragraph_delta = isize::from(plan.inserted_paragraph) - plan.removed as isize;
        let moved_ordinal = |ordinal: usize, delta: isize| ordinal.saturating_add_signed(delta);
        // Kept owners take their next states, in inventory order, by move.
        let mut states = states.into_iter();
        for (index, owner) in self.owners.iter_mut().enumerate() {
            note_fx260_word(1);
            if spans.iter().any(|span| span.contains(&index)) {
                for _ in 0..owner.location.scopes().len() {
                    states.next();
                }
                continue;
            }
            let item = word_owner_body_item(&owner.location);
            let body_route = matches!(
                &owner.location,
                WordOwnerLocation::Paragraph(location) | WordOwnerLocation::Table(location)
                    if location.story().kind() == rdocx::StoryKind::Body
            );
            if let Some(item) = item.filter(|item| *item >= after) {
                move_word_owner_body_item(&mut owner.location, plan.moved(item));
                if body_route {
                    let delta = match owner.location {
                        WordOwnerLocation::Paragraph(_) => paragraph_delta,
                        _ => isize::from(plan.inserted_table),
                    };
                    owner.projection_ordinal = owner
                        .projection_ordinal
                        .map(|ordinal| moved_ordinal(ordinal, delta));
                }
            }
            if body_route {
                owner.body_slot = owner.body_slot.map(|slot| {
                    if slot >= plan.slot + plan.removed {
                        slot + usize::from(plan.inserted_paragraph || plan.inserted_table)
                            - plan.removed
                    } else {
                        slot
                    }
                });
            }
            plan.move_cell_stories(&mut owner.location, &part);
            owner.parent = owner.parent.map(|parent| {
                plan.next_index(parent)
                    .expect("a kept owner's parent is kept")
            });
            let mut next = || states.next().expect("a prepared state for every scope");
            owner.owner = next();
            for state in [&mut owner.content, &mut owner.properties, &mut owner.grid]
                .into_iter()
                .flatten()
            {
                *state = next();
            }
        }
        self.indexes.index_owners_again(&self.owners);
        let moved = |location: &mut rdocx::ContentLocation| {
            if location.story().kind() == rdocx::StoryKind::Body {
                let item = location.index_path()[0];
                location.set_index_path_head(plan.moved(item));
            }
        };
        if let Some(paragraphs) = &mut self.indexes.body_paragraphs {
            paragraphs.drain(plan.paragraph..plan.paragraph + plan.removed);
            paragraphs.iter_mut().flatten().for_each(moved);
            if plan.inserted_paragraph
                && let Some(WordOwnerLocation::Paragraph(location)) = plan
                    .blocks
                    .first()
                    .and_then(|(_, block)| block.first())
                    .map(|address| &address.location)
            {
                paragraphs.insert(plan.paragraph, Some(location.clone()));
            }
        }
        self.indexes
            .body_tables
            .iter_mut()
            .flatten()
            .for_each(moved);
        if let Some((location, cells)) = &plan.table_entry {
            self.indexes
                .body_tables
                .insert(plan.table, Some(location.clone()));
            let mut shifted =
                HashMap::with_capacity(self.indexes.cell_paragraphs.len() + cells.len());
            for ((table, row, cell), mut locations) in self.indexes.cell_paragraphs.drain() {
                for location in &mut locations {
                    let mut route = WordOwnerLocation::Paragraph(location.clone());
                    plan.move_cell_stories(&mut route, &part);
                    if let WordOwnerLocation::Paragraph(route) = route {
                        *location = route;
                    }
                }
                let table = if table >= plan.table {
                    table + 1
                } else {
                    table
                };
                shifted.insert((table, row, cell), locations);
            }
            for ((row, cell), locations) in cells {
                shifted.insert((plan.table, *row, *cell), locations.clone());
            }
            self.indexes.cell_paragraphs = shifted;
        }
        self.snapshot_scope = snapshot_scope;
        let inserted = plan.piece.take();
        self.pieces
            .splice(&self.inner, plan.slot, plan.removed, inserted);
        #[cfg(debug_assertions)]
        self.assert_word_inventory_rebuilds();
        Ok(())
    }

    /// Under debug assertions, prove that an append keeps exactly the owners
    /// the rebuild's matcher keeps. Runs after the native edit.
    #[cfg(debug_assertions)]
    fn assert_body_append_matches(&self, plan: &BodySplicePlan, publication: &WordPublication) {
        #[cfg(test)]
        let visits = fx260_word_visits();
        let rebuilt = guarded(|| word_owner_locations(&self.inner, None))
            .expect("a spliced inventory has a rebuild");
        let mut expected = vec![None; self.owners.len() + plan.added()];
        for index in 0..self.owners.len() {
            if let Some(next) = plan.next_index(index) {
                expected[next] = Some(index);
            }
        }
        assert_eq!(
            indexed_inventory_matches(&self.owners, &rebuilt, publication),
            expected,
            "spliced Word owner matches"
        );
        #[cfg(test)]
        FX260_WORD_VISITS.with(|cell| cell.set(visits));
    }

    /// Under debug assertions, prove the inventory and its document indexes
    /// equal their rebuild from the current document, story fingerprints
    /// aside, which local edits leave stale in the inventory too.
    #[cfg(debug_assertions)]
    fn assert_word_inventory_rebuilds(&self) {
        #[cfg(test)]
        let visits = fx260_word_visits();
        let address = |location: &WordOwnerLocation,
                       parent: Option<usize>,
                       body_slot: Option<usize>,
                       projection_ordinal: Option<usize>,
                       read_facts: &WordReadFacts| WordOwnerAddress {
            location: unfingerprinted_owner_location(location),
            parent,
            body_slot,
            projection_ordinal,
            read_facts: read_facts.clone(),
        };
        let actual = self
            .owners
            .iter()
            .map(|owner| {
                address(
                    &owner.location,
                    owner.parent,
                    owner.body_slot,
                    owner.projection_ordinal,
                    &owner.read_facts,
                )
            })
            .collect::<Vec<_>>();
        let expected = guarded(|| word_owner_locations(&self.inner, None))
            .expect("a spliced inventory has a rebuild")
            .iter()
            .map(|owner| {
                address(
                    &owner.location,
                    owner.parent,
                    owner.body_slot,
                    owner.projection_ordinal,
                    &owner.read_facts,
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "spliced Word owner inventory");
        let indexes =
            WordOwnerIndexes::prepare(&self.owners, WordIndexProjections::of(&self.inner))
                .expect("a spliced inventory has rebuilt indexes");
        let locations = |locations: &[Option<rdocx::ContentLocation>]| {
            locations
                .iter()
                .map(|location| location.as_ref().map(unfingerprinted_location))
                .collect::<Vec<_>>()
        };
        assert_eq!(
            self.indexes.body_paragraphs.as_deref().map(locations),
            indexes.body_paragraphs.as_deref().map(locations),
            "spliced body paragraph index"
        );
        assert_eq!(
            locations(&self.indexes.body_tables),
            locations(&indexes.body_tables),
            "spliced body table index"
        );
        let cells = |cells: &HashMap<(usize, usize, usize), Vec<rdocx::ContentLocation>>| {
            cells
                .iter()
                .map(|(key, locations)| {
                    (
                        *key,
                        locations
                            .iter()
                            .map(unfingerprinted_location)
                            .collect::<Vec<_>>(),
                    )
                })
                .collect::<BTreeMap<_, _>>()
        };
        assert_eq!(
            cells(&self.indexes.cell_paragraphs),
            cells(&indexes.cell_paragraphs),
            "spliced cell paragraph index"
        );
        assert_eq!(
            self.indexes.direct_body_paragraphs, indexes.direct_body_paragraphs,
            "spliced direct body paragraph index"
        );
        #[cfg(test)]
        FX260_WORD_VISITS.with(|cell| cell.set(visits));
    }
}

/// The inserted blocks of a detached table placed at body story item
/// `first` of slot `slot`: the table and its cells before current owner
/// `table_at`, its cell paragraphs and runs before `cells_at`, with cell
/// story owners counted from `cells_before`. Also the number of cells, the
/// table route with the routes of each cell's paragraphs, and the table's
/// owners as the inventory by block reads them, when its read facts give
/// them.
#[allow(clippy::type_complexity)] // The blocks, the cell count, the index entry and the owners of one table.
fn spliced_table_blocks(
    xml: Vec<u8>,
    snapshots: WordOwnerSnapshots,
    story: &rdocx::StoryId,
    (slot, first, ordinal): (usize, usize, usize),
    (table_at, cells_at, cells_before): (usize, usize, usize),
) -> Option<(
    Vec<(usize, Vec<WordOwnerAddress>)>,
    usize,
    (rdocx::ContentLocation, Vec<CellParagraphs>),
    Option<WordBlockOwners>,
)> {
    // The detached body holds this table as its only item.
    let mut items = snapshots
        .0
        .iter()
        .filter(|item| item.location().story().kind() == rdocx::StoryKind::Body);
    let item = items.next()?;
    if items.next().is_some()
        || item.location().item_kind() != rdocx::StoryItemKind::Table
        || item.location().index_path() != [0]
    {
        return None;
    }
    let table = item.location().clone();
    let facts = word_owner_facts(snapshots, |snapshot| {
        if snapshot.location() == &table {
            word_table_facts(&xml).unwrap_or_default()
        } else {
            WordReadFacts::None
        }
    })
    .ok()?;
    let piece = facts.clone();
    let addresses = word_owner_addresses(facts, 0);
    let cells = addresses
        .iter()
        .skip(1)
        .take_while(|address| matches!(address.location, WordOwnerLocation::Cell { .. }))
        .count();
    // The table, its cells in story order, then only their paragraphs and runs.
    let cell_story = |address: &WordOwnerAddress| match &address.location {
        WordOwnerLocation::Cell { story, .. } => Some(story.owner_index()),
        _ => None,
    };
    let in_cells = |location: &rdocx::ContentLocation| {
        location.story().kind() == rdocx::StoryKind::TableCell
            && location.story().owner_index() < cells
    };
    let cell_parent =
        |parent: Option<usize>| parent.is_some_and(|parent| (1..=cells).contains(&parent));
    let shaped = matches!(addresses.first()?.location, WordOwnerLocation::Table(_))
        && addresses[1..=cells]
            .iter()
            .enumerate()
            .all(|(index, address)| {
                cell_story(address) == Some(index) && address.parent == Some(0)
            })
        && addresses[cells + 1..]
            .iter()
            .all(|address| match &address.location {
                WordOwnerLocation::Paragraph(location) => {
                    in_cells(location) && cell_parent(address.parent)
                }
                WordOwnerLocation::Run(location, _) => {
                    in_cells(location) && address.parent.is_some_and(|parent| parent > cells)
                }
                _ => false,
            });
    if !shaped {
        return None;
    }
    // Detached owner `index` takes this next-inventory index.
    let next = |index: usize| {
        if index <= cells {
            table_at + index
        } else {
            cells_at + index
        }
    };
    let moved = |location: &rdocx::ContentLocation| {
        let mut story = location.story().clone();
        story.set_owner_index(story.owner_index() + cells_before);
        rdocx::ContentLocation::new(story, location.item_kind(), location.index_path().to_vec())
    };
    let table =
        rdocx::ContentLocation::new(story.clone(), rdocx::StoryItemKind::Table, vec![first]);
    let mut entries = Vec::with_capacity(cells);
    let mut head = Vec::with_capacity(cells + 1);
    let mut tail = Vec::with_capacity(addresses.len() - cells - 1);
    for (index, address) in addresses.into_iter().enumerate() {
        let parent = address.parent.map(next);
        let location = match &address.location {
            WordOwnerLocation::Table(_) => WordOwnerLocation::Table(table.clone()),
            WordOwnerLocation::Cell {
                row, cell, story, ..
            } => {
                let mut story = story.clone();
                story.set_owner_index(story.owner_index() + cells_before);
                entries.push(((*row, *cell), Vec::new()));
                WordOwnerLocation::Cell {
                    table: table.clone(),
                    row: *row,
                    cell: *cell,
                    story,
                }
            }
            WordOwnerLocation::Paragraph(location) => {
                let location = moved(location);
                entries[address.parent? - 1].1.push(location.clone());
                WordOwnerLocation::Paragraph(location)
            }
            WordOwnerLocation::Run(location, path) => {
                WordOwnerLocation::Run(moved(location), path.clone())
            }
            _ => return None,
        };
        let address = if index == 0 {
            WordOwnerAddress {
                location,
                parent: None,
                body_slot: Some(slot),
                projection_ordinal: Some(ordinal),
                read_facts: address.read_facts,
            }
        } else {
            WordOwnerAddress {
                location,
                parent,
                ..address
            }
        };
        if index <= cells {
            head.push(address);
        } else {
            tail.push(address);
        }
    }
    let mut blocks = vec![(table_at, head)];
    if !tail.is_empty() {
        blocks.push((cells_at, tail));
    }
    Some((blocks, cells, (table, entries), table_owners(piece, cells)))
}

#[derive(Clone)]
pub(crate) enum WordPublication {
    Conservative,
    FormValues,
    Signing,
    AppendBody,
    HeaderFooterAppend(usize),
    ParagraphText(rdocx::ContentLocation),
    ParagraphAppend(rdocx::ContentLocation),
    TableGrid(rdocx::ContentLocation),
    HyperlinkRemove(WordOwnerLocation),
    CellText(rdocx::ContentLocation, usize, usize),
    ElementXml(WordOwnerLocation),
    RunText(WordOwnerLocation),
}

fn form_owner_certificates(
    old: &[WordOwner],
    new: &[WordOwnerAddress],
    matched: &[Option<usize>],
    before: &FormSources,
    after: &FormSources,
) -> std::collections::HashSet<(usize, usize)> {
    let mut old_routes = HashMap::<FormOwnerRoute, Vec<usize>>::new();
    let mut new_routes = HashMap::<FormOwnerRoute, Vec<usize>>::new();
    for (index, owner) in old.iter().enumerate() {
        if let Some(route) = form_owner_route(&owner.location) {
            old_routes.entry(route).or_default().push(index);
        }
    }
    for (index, owner) in new.iter().enumerate() {
        if let Some(route) = form_owner_route(&owner.location) {
            new_routes.entry(route).or_default().push(index);
        }
    }
    let mut proved = std::collections::HashSet::new();
    let mut ready = vec![false; new.len()];
    for target in 0..new.len() {
        let mut stack = vec![(target, false)];
        while let Some((index, parent_ready)) = stack.pop() {
            if ready[index] {
                continue;
            }
            if !parent_ready && let Some(parent) = new[index].parent {
                stack.push((index, true));
                if !ready[parent] {
                    stack.push((parent, false));
                }
                continue;
            }
            ready[index] = true;
            let Some(route) = form_owner_route(&new[index].location) else {
                continue;
            };
            let Some(indices) = old_routes.get(&route) else {
                continue;
            };
            let [previous] = indices.as_slice() else {
                continue;
            };
            if new_routes.get(&route).map(Vec::as_slice) != Some(&[index])
                || matched[index] != Some(*previous)
            {
                continue;
            }
            let source_proved = match &route {
                FormOwnerRoute::Cell(_, _, _, story) => {
                    before.cells.get(&route) == Some(&1)
                        && after.cells.get(&route) == Some(&1)
                        && before.cell_stories.get(story) == Some(&1)
                        && after.cell_stories.get(story) == Some(&1)
                }
                FormOwnerRoute::Paragraph(key) | FormOwnerRoute::Table(key)
                    if matches!(key.0, rdocx::StoryKind::Comment | rdocx::StoryKind::TextBox) =>
                {
                    false
                }
                _ => {
                    matches!((before.xml.get(&route), after.xml.get(&route)), (Some(Some(a)), Some(Some(b))) if a == b)
                }
            };
            if !source_proved {
                continue;
            }
            let ancestors_proved = match (old[*previous].parent, new[index].parent) {
                (Some(a), Some(b)) => proved.contains(&(a, b)),
                (None, None) => !matches!(
                    &route,
                    FormOwnerRoute::Cell(..)
                        | FormOwnerRoute::Paragraph((rdocx::StoryKind::TableCell, ..))
                        | FormOwnerRoute::Table((rdocx::StoryKind::TableCell, ..))
                ),
                _ => false,
            };
            if ancestors_proved {
                proved.insert((*previous, index));
            }
        }
    }
    proved
}

// Signing changes package infrastructure. Preserve content authority only when
// the physical source and every required ancestor survived serialization unchanged.
fn signing_owner_certificates(
    old: &[WordOwner],
    new: &[WordOwnerAddress],
    matched: &[Option<usize>],
    before: &FormSources,
    after: &FormSources,
) -> std::collections::HashSet<(usize, usize)> {
    let mut proved = form_owner_certificates(old, new, matched, before, after);
    let mut paragraphs = HashMap::<FormSourceKey, Vec<(usize, usize)>>::new();
    for &(previous, next) in &proved {
        if let WordOwnerLocation::Paragraph(location) = &new[next].location {
            paragraphs
                .entry(form_source_key(location))
                .or_default()
                .push((previous, next));
        }
    }
    for (next, address) in new.iter().enumerate() {
        let Some(previous) = matched[next] else {
            continue;
        };
        let survives = match (&old[previous].location, &address.location) {
            (WordOwnerLocation::Run(a, left), WordOwnerLocation::Run(b, right)) => {
                left == right
                    && form_source_key(a) == form_source_key(b)
                    && matches!((old[previous].parent, address.parent), (Some(a), Some(b)) if proved.contains(&(a, b)))
            }
            (WordOwnerLocation::Hyperlink(a, left), WordOwnerLocation::Hyperlink(b, right)) => {
                left == right
                    && form_source_key(a) == form_source_key(b)
                    && paragraphs
                        .get(&form_source_key(b))
                        .is_some_and(|pairs| pairs.len() == 1)
            }
            _ => false,
        };
        if survives {
            proved.insert((previous, next));
        }
    }
    proved
}

fn same_inventory_route(
    old: &[WordOwner],
    mut before: usize,
    new: &[WordOwnerAddress],
    mut after: usize,
) -> bool {
    loop {
        #[cfg(test)]
        FX237_ROUTE_WORK.set({
            let mut work = FX237_ROUTE_WORK.get();
            work[4] += 1;
            work
        });
        let record = &old[before];
        let address = &new[after];
        let b = &address.location;
        let parent = address.parent;
        let equal = match (&record.location, b) {
            (WordOwnerLocation::Paragraph(a), WordOwnerLocation::Paragraph(b))
                if a.story().kind() == rdocx::StoryKind::Body
                    && b.story().kind() == rdocx::StoryKind::Body
                    && record.body_slot.is_some()
                    && address.body_slot.is_some() =>
            {
                record.body_slot == address.body_slot
                    && a.index_path().len() <= 2
                    && b.index_path().len() <= 2
                    && a.index_path().get(1..) == b.index_path().get(1..)
            }
            (WordOwnerLocation::Table(a), WordOwnerLocation::Table(b))
                if a.story().kind() == rdocx::StoryKind::Body
                    && b.story().kind() == rdocx::StoryKind::Body
                    && record.body_slot.is_some()
                    && address.body_slot.is_some() =>
            {
                record.body_slot == address.body_slot
            }
            (WordOwnerLocation::Paragraph(a), WordOwnerLocation::Paragraph(b))
            | (WordOwnerLocation::Table(a), WordOwnerLocation::Table(b))
                if record.parent.is_some_and(|index| {
                    matches!(old[index].location, WordOwnerLocation::Cell { .. })
                }) && parent.is_some_and(|index| {
                    matches!(new[index].location, WordOwnerLocation::Cell { .. })
                }) =>
            {
                a.item_kind() == b.item_kind()
                    && record.projection_ordinal == address.projection_ordinal
            }
            (WordOwnerLocation::Paragraph(a), WordOwnerLocation::Paragraph(b))
            | (WordOwnerLocation::Table(a), WordOwnerLocation::Table(b)) => {
                same_story_owner(a.story(), b.story())
                    && a.item_kind() == b.item_kind()
                    && record.projection_ordinal == address.projection_ordinal
            }
            (
                WordOwnerLocation::Cell {
                    row: a, cell: ac, ..
                },
                WordOwnerLocation::Cell {
                    row: b, cell: bc, ..
                },
            ) => a == b && ac == bc,
            (WordOwnerLocation::Run(_, a), WordOwnerLocation::Run(_, b)) => a == b,
            (a, b) => a.same_route(b),
        };
        if !equal {
            return false;
        }
        match (record.parent, parent) {
            (Some(a), Some(b)) => {
                before = a;
                after = b;
            }
            (None, None) => return true,
            _ => return false,
        }
    }
}

// The legacy relation is pairwise, not transitive. Index both slot and
// projection routes, partitioning the old operands before selecting a branch.
#[derive(Clone, PartialEq, Eq, Hash)]
enum WordRouteKey {
    Slot(rdocx::StoryItemKind, usize, Vec<usize>),
    Projection(
        rdocx::StoryKind,
        String,
        usize,
        rdocx::StoryItemKind,
        Option<usize>,
    ),
    CellProjection(rdocx::StoryItemKind, Option<usize>),
    Cell(usize, usize),
    Run(rdocx_oxml::text::AcceptedRunPath),
    Section(usize),
    Hyperlink(
        rdocx::StoryKind,
        String,
        usize,
        rdocx::StoryItemKind,
        Vec<usize>,
        usize,
    ),
}

fn word_route_key(
    address: &WordOwnerAddress,
    class: usize,
    parent_cell: bool,
) -> Option<WordRouteKey> {
    match &address.location {
        WordOwnerLocation::Paragraph(location) | WordOwnerLocation::Table(location) => {
            let body = location.story().kind() == rdocx::StoryKind::Body;
            if body
                && class & 1 != 0
                && let Some(slot) = address.body_slot
            {
                let suffix = if matches!(address.location, WordOwnerLocation::Paragraph(_)) {
                    if location.index_path().len() > 2 {
                        return None;
                    }
                    location.index_path().get(1..).unwrap_or_default().to_vec()
                } else {
                    Vec::new()
                };
                let kind = if matches!(address.location, WordOwnerLocation::Paragraph(_)) {
                    rdocx::StoryItemKind::Paragraph
                } else {
                    rdocx::StoryItemKind::Table
                };
                Some(WordRouteKey::Slot(kind, slot, suffix))
            } else if parent_cell && class & 2 != 0 {
                Some(WordRouteKey::CellProjection(
                    location.item_kind(),
                    address.projection_ordinal,
                ))
            } else {
                let story = location.story();
                Some(WordRouteKey::Projection(
                    story.kind(),
                    story.part_name().to_owned(),
                    story.owner_index(),
                    location.item_kind(),
                    address.projection_ordinal,
                ))
            }
        }
        WordOwnerLocation::Cell { row, cell, .. } => Some(WordRouteKey::Cell(*row, *cell)),
        WordOwnerLocation::Run(_, path) => Some(WordRouteKey::Run(path.clone())),
        WordOwnerLocation::Section(index) => Some(WordRouteKey::Section(*index)),
        WordOwnerLocation::Hyperlink(location, index) => {
            let story = location.story();
            Some(WordRouteKey::Hyperlink(
                story.kind(),
                story.part_name().to_owned(),
                story.owner_index(),
                location.item_kind(),
                location.index_path().to_vec(),
                *index,
            ))
        }
    }
}

#[cfg(test)]
thread_local! {
    static FX237_ROUTE_WORK: std::cell::Cell<[usize; 8]> = const { std::cell::Cell::new([0; 8]) };
}

#[derive(Clone, PartialEq, Eq, Hash)]
enum PropertyOwnerKey {
    Table {
        story_kind: rdocx::StoryKind,
        part_name: String,
        path: Vec<usize>,
        cell: Option<(usize, usize)>,
    },
    Cell {
        story_kind: rdocx::StoryKind,
        part_name: String,
        path: Vec<usize>,
        row: usize,
        column: usize,
        parent_cell: Option<(usize, usize)>,
    },
    InCell {
        story_kind: rdocx::StoryKind,
        part_name: String,
        depth: usize,
        cell: (usize, usize),
        run: Option<rdocx_oxml::text::AcceptedRunPath>,
    },
    BodyItem {
        story_kind: rdocx::StoryKind,
        part_name: String,
        kind: rdocx::StoryItemKind,
        path: Vec<usize>,
    },
    BodyRun {
        story_kind: rdocx::StoryKind,
        part_name: String,
        path: Vec<usize>,
        run: rdocx_oxml::text::AcceptedRunPath,
    },
    Section(usize),
    Hyperlink {
        story_kind: rdocx::StoryKind,
        part_name: String,
        path: Vec<usize>,
        index: usize,
    },
}

fn containing_cell(
    addresses: &[WordOwnerAddress],
    mut parent: Option<usize>,
) -> Option<(usize, usize)> {
    while let Some(index) = parent {
        if let WordOwnerLocation::Cell { row, cell, .. } = &addresses[index].location {
            return Some((*row, *cell));
        }
        parent = addresses[index].parent;
    }
    None
}

fn story_part(location: &rdocx::ContentLocation) -> (rdocx::StoryKind, String) {
    let story = location.story();
    (story.kind(), story.part_name().to_owned())
}

fn property_owner_key(addresses: &[WordOwnerAddress], index: usize) -> PropertyOwnerKey {
    let address = &addresses[index];
    let cell = containing_cell(addresses, address.parent);
    match &address.location {
        WordOwnerLocation::Table(location) => {
            let (story_kind, part_name) = story_part(location);
            PropertyOwnerKey::Table {
                story_kind,
                part_name,
                path: location.index_path().to_vec(),
                cell,
            }
        }
        WordOwnerLocation::Cell {
            table,
            row,
            cell: column,
            ..
        } => {
            let (story_kind, part_name) = story_part(table);
            PropertyOwnerKey::Cell {
                story_kind,
                part_name,
                path: table.index_path().to_vec(),
                row: *row,
                column: *column,
                parent_cell: cell,
            }
        }
        WordOwnerLocation::Paragraph(location) => {
            let (story_kind, part_name) = story_part(location);
            if let Some(cell) = cell {
                PropertyOwnerKey::InCell {
                    story_kind,
                    part_name,
                    depth: location.index_path().len(),
                    cell,
                    run: None,
                }
            } else {
                PropertyOwnerKey::BodyItem {
                    story_kind,
                    part_name,
                    kind: location.item_kind(),
                    path: location.index_path().to_vec(),
                }
            }
        }
        WordOwnerLocation::Run(location, path) => {
            let (story_kind, part_name) = story_part(location);
            if let Some(cell) = cell {
                PropertyOwnerKey::InCell {
                    story_kind,
                    part_name,
                    depth: location.index_path().len(),
                    cell,
                    run: Some(path.clone()),
                }
            } else {
                PropertyOwnerKey::BodyRun {
                    story_kind,
                    part_name,
                    path: location.index_path().to_vec(),
                    run: path.clone(),
                }
            }
        }
        WordOwnerLocation::Section(index) => PropertyOwnerKey::Section(*index),
        WordOwnerLocation::Hyperlink(location, index) => {
            let (story_kind, part_name) = story_part(location);
            PropertyOwnerKey::Hyperlink {
                story_kind,
                part_name,
                path: location.index_path().to_vec(),
                index: *index,
            }
        }
    }
}

fn relocated_property_owners(
    old: &[WordOwner],
    new_addresses: &[WordOwnerAddress],
) -> Option<Vec<WordOwner>> {
    if old.len() != new_addresses.len() {
        return None;
    }
    let old_addresses = old
        .iter()
        .map(|owner| WordOwnerAddress {
            location: owner.location.clone(),
            parent: owner.parent,
            body_slot: owner.body_slot,
            projection_ordinal: owner.projection_ordinal,
            read_facts: owner.read_facts.clone(),
        })
        .collect::<Vec<_>>();
    let mut buckets: HashMap<PropertyOwnerKey, Vec<usize>> = HashMap::new();
    for index in 0..old_addresses.len() {
        buckets
            .entry(property_owner_key(&old_addresses, index))
            .or_default()
            .push(index);
    }
    let mut cursors: HashMap<PropertyOwnerKey, usize> = HashMap::new();
    let mut owners = Vec::with_capacity(new_addresses.len());
    for index in 0..new_addresses.len() {
        let key = property_owner_key(new_addresses, index);
        let bucket = buckets.get(&key)?;
        let cursor = cursors.entry(key).or_insert(0);
        if *cursor >= bucket.len() {
            return None;
        }
        let previous = &old[bucket[*cursor]];
        *cursor += 1;
        let address = &new_addresses[index];
        owners.push(WordOwner {
            location: address.location.clone(),
            parent: address.parent,
            body_slot: address.body_slot,
            projection_ordinal: address.projection_ordinal,
            read_facts: address.read_facts.clone(),
            owner: previous.owner.clone(),
            content: previous.content.clone(),
            properties: previous.properties.clone(),
            grid: previous.grid.clone(),
        });
    }
    if cursors.len() != buckets.len()
        || cursors
            .iter()
            .any(|(key, cursor)| buckets.get(key).map(Vec::len) != Some(*cursor))
    {
        return None;
    }
    Some(owners)
}

fn indexed_inventory_matches(
    old: &[WordOwner],
    new: &[WordOwnerAddress],
    publication: &WordPublication,
) -> Vec<Option<usize>> {
    let mut index = HashMap::<(Option<usize>, usize, WordRouteKey), Vec<usize>>::new();
    let mut links = HashMap::<WordRouteKey, Vec<usize>>::new();
    for (before, owner) in old.iter().enumerate() {
        let address = WordOwnerAddress {
            location: owner.location.clone(),
            parent: owner.parent,
            body_slot: owner.body_slot,
            projection_ordinal: owner.projection_ordinal,
            read_facts: owner.read_facts.clone(),
        };
        let parent_cell = owner
            .parent
            .is_some_and(|parent| matches!(old[parent].location, WordOwnerLocation::Cell { .. }));
        let body_slot = match &owner.location {
            WordOwnerLocation::Paragraph(location) | WordOwnerLocation::Table(location) => {
                location.story().kind() == rdocx::StoryKind::Body && owner.body_slot.is_some()
            }
            _ => false,
        };
        let class = usize::from(body_slot) | (usize::from(parent_cell) << 1);
        // The candidate determines which pairwise branch applies. Supply each
        // branch's old key in its own namespace, without merging parent routes.
        let mut keys = Vec::new();
        for candidate_class in 0..4 {
            if let Some(key) = word_route_key(&address, candidate_class, candidate_class & 2 != 0)
                && !keys.contains(&key)
            {
                keys.push(key);
            }
        }
        for key in keys {
            #[cfg(test)]
            FX237_ROUTE_WORK.set({
                let mut work = FX237_ROUTE_WORK.get();
                work[6] += 1;
                work
            });
            index
                .entry((owner.parent, class, key.clone()))
                .or_default()
                .push(before);
            if matches!(owner.location, WordOwnerLocation::Hyperlink(..)) {
                links.entry(key).or_default().push(before);
            }
        }
        #[cfg(test)]
        FX237_ROUTE_WORK.set({
            let mut work = FX237_ROUTE_WORK.get();
            work[0] += 1;
            work
        });
    }
    #[cfg(test)]
    FX237_ROUTE_WORK.set({
        let mut work = FX237_ROUTE_WORK.get();
        work[7] = index.values().map(Vec::len).max().unwrap_or(0);
        work
    });
    let mut matches = vec![None::<Vec<usize>>; new.len()];
    for target in 0..new.len() {
        let mut stack = vec![(target, false)];
        while let Some((after, ready)) = stack.pop() {
            if matches[after].is_some() {
                continue;
            }
            if !ready && let Some(parent) = new[after].parent {
                stack.push((after, true));
                if matches[parent].is_none() {
                    stack.push((parent, false));
                }
                continue;
            }
            let address = &new[after];
            let mut found = Vec::new();
            // Removal's original override uses the source content route alone,
            // independently of the normal parent-chain matcher.
            let removed_link = if let (
                WordPublication::HyperlinkRemove(WordOwnerLocation::Hyperlink(
                    removed,
                    removed_index,
                )),
                WordOwnerLocation::Hyperlink(location, position),
            ) = (publication, &address.location)
            {
                if same_content_owner(removed, location) {
                    let previous = if position >= removed_index {
                        position.checked_add(1)
                    } else {
                        Some(*position)
                    };
                    if let Some(previous) = previous {
                        let shifted = WordOwnerAddress {
                            location: WordOwnerLocation::Hyperlink(location.clone(), previous),
                            parent: None,
                            body_slot: None,
                            projection_ordinal: None,
                            read_facts: WordReadFacts::None,
                        };
                        if let Some(key) = word_route_key(&shifted, 0, false) {
                            found.extend(links.get(&key).into_iter().flatten().copied());
                        }
                    }
                    true
                } else {
                    false
                }
            } else {
                false
            };
            if !removed_link {
                let parents = address.parent.map_or_else(
                    || vec![None],
                    |parent| {
                        matches[parent]
                            .as_ref()
                            .unwrap()
                            .iter()
                            .copied()
                            .map(Some)
                            .collect()
                    },
                );
                let parent_cell = address.parent.is_some_and(|parent| {
                    matches!(new[parent].location, WordOwnerLocation::Cell { .. })
                });
                for parent in parents {
                    #[cfg(test)]
                    FX237_ROUTE_WORK.set({
                        let mut work = FX237_ROUTE_WORK.get();
                        work[5] += 1;
                        work
                    });
                    for class in 0..4 {
                        if let Some(key) = word_route_key(address, class, parent_cell)
                            && let Some(candidates) = index.get(&(parent, class, key))
                        {
                            for &before in candidates {
                                #[cfg(test)]
                                FX237_ROUTE_WORK.set({
                                    let mut work = FX237_ROUTE_WORK.get();
                                    work[2] += 1;
                                    work
                                });
                                if same_inventory_route(old, before, new, after) {
                                    found.push(before);
                                }
                            }
                        }
                    }
                }
            }
            found.sort_unstable();
            found.dedup();
            #[cfg(test)]
            FX237_ROUTE_WORK.set({
                let mut work = FX237_ROUTE_WORK.get();
                work[1] += 1;
                work[3] = work[3].max(found.len());
                work
            });
            matches[after] = Some(found);
        }
    }
    matches
        .into_iter()
        .map(|matched| matched.unwrap().into_iter().next())
        .collect()
}

impl WordOwner {
    fn states(&self) -> Vec<RevisionState> {
        self.scope_states().cloned().collect()
    }

    /// The owner's scope states, in [`WordOwnerLocation::scopes`] order.
    fn scope_states(&self) -> impl Iterator<Item = &RevisionState> {
        std::iter::once(&self.owner)
            .chain(&self.content)
            .chain(&self.properties)
            .chain(&self.grid)
    }
}

fn word_owners_with_states(
    locations: Vec<WordOwnerAddress>,
    states: &[RevisionState],
) -> Vec<WordOwner> {
    let mut offset = 0;
    locations
        .into_iter()
        .map(|address| {
            let WordOwnerAddress {
                location,
                parent,
                body_slot,
                projection_ordinal,
                read_facts,
            } = address;
            let scopes = &states[offset..offset + location.scopes().len()];
            offset += scopes.len();
            let (content, properties, grid) = match &location {
                WordOwnerLocation::Paragraph(_) => {
                    (Some(scopes[1].clone()), Some(scopes[2].clone()), None)
                }
                WordOwnerLocation::Table(_) => (None, None, Some(scopes[1].clone())),
                WordOwnerLocation::Hyperlink(..) => (None, None, None),
                _ => (Some(scopes[1].clone()), None, None),
            };
            WordOwner {
                location,
                parent,
                body_slot,
                projection_ordinal,
                read_facts,
                owner: scopes[0].clone(),
                content,
                properties,
                grid,
            }
        })
        .collect()
}

#[pyclass(name = "Document")]
pub struct PyDocument {
    pub(crate) inner: rdocx::Document,
    pub(crate) revisions: RevisionCounter,
    owners: Vec<WordOwner>,
    indexes: WordOwnerIndexes,
    removal_scope: RevisionState,
    snapshot_scope: RevisionState,
    /// The owner inventory by body block and story part, from which the
    /// next inventory reads only what changed.
    pieces: WordPieces,
}

#[pyclass(name = "FragmentConflictPolicy", frozen, skip_from_py_object)]
#[derive(Clone, Copy)]
pub struct PyFragmentConflictPolicy {
    inner: rdocx::FragmentConflictPolicy,
}
#[pymethods]
impl PyFragmentConflictPolicy {
    #[staticmethod]
    fn reuse_equivalent() -> Self {
        Self {
            inner: rdocx::FragmentConflictPolicy::reuse_equivalent(),
        }
    }
    #[staticmethod]
    fn rename_all() -> Self {
        Self {
            inner: rdocx::FragmentConflictPolicy::rename_all(),
        }
    }
    fn with_style_reuse(&self, reuse: bool) -> Self {
        Self {
            inner: self.inner.with_style_reuse(reuse),
        }
    }
    fn with_numbering_reuse(&self, reuse: bool) -> Self {
        Self {
            inner: self.inner.with_numbering_reuse(reuse),
        }
    }
    fn with_related_part_reuse(&self, reuse: bool) -> Self {
        Self {
            inner: self.inner.with_related_part_reuse(reuse),
        }
    }
}

#[pyclass(name = "ContentLocation", frozen, skip_from_py_object)]
#[derive(Clone)]
pub struct PyContentLocation {
    story: PyStory,
    kind: String,
    index_path: Vec<usize>,
    is_end: bool,
}
impl PyContentLocation {
    fn native(&self, py: Python<'_>, document: &PyDocument) -> PyResult<rdocx::ContentLocation> {
        let story = document.native_story(py, &self.story)?;
        if self.is_end {
            Ok(rdocx::ContentLocation::end(story))
        } else {
            Ok(rdocx::ContentLocation::new(
                story,
                story_item_kind_from_name(&self.kind)?,
                self.index_path.clone(),
            ))
        }
    }
}
#[pymethods]
impl PyContentLocation {
    #[new]
    fn new(story: PyRef<'_, PyStory>, item_kind: String, index_path: Vec<usize>) -> PyResult<Self> {
        story_item_kind_from_name(&item_kind)?;
        Ok(Self {
            story: story.clone(),
            kind: item_kind,
            index_path,
            is_end: false,
        })
    }
    #[staticmethod]
    fn end(story: PyRef<'_, PyStory>) -> Self {
        Self {
            story: story.clone(),
            kind: "preserved_node".into(),
            index_path: Vec::new(),
            is_end: true,
        }
    }
    #[getter]
    fn story(&self) -> PyStory {
        self.story.clone()
    }
    #[getter]
    fn item_kind(&self) -> &str {
        &self.kind
    }
    #[getter]
    fn index_path<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.index_path)
    }
}

#[pyclass(name = "DocumentFragment", frozen, skip_from_py_object)]
pub struct PyDocumentFragment {
    inner: rdocx::DocumentFragment,
}
#[pymethods]
impl PyDocumentFragment {
    #[staticmethod]
    #[pyo3(signature = (document, start, end, *, include_final_section_properties=false))]
    fn from_range(
        py: Python<'_>,
        document: PyRef<'_, PyDocument>,
        start: PyRef<'_, PyContentLocation>,
        end: PyRef<'_, PyContentLocation>,
        include_final_section_properties: bool,
    ) -> PyResult<Self> {
        let start = start.native(py, &document)?;
        let end = end.native(py, &document)?;
        rdocx::DocumentFragment::from_range(
            &document.inner,
            &start,
            &end,
            include_final_section_properties,
        )
        .map(|inner| Self { inner })
        .map_err(|e| rdocx_to_pyerr(py, e))
    }
}

/// An owned immutable legacy field inventory entry.
#[pyclass(name = "LegacyFormField", frozen)]
pub struct PyLegacyFormField {
    #[pyo3(get)]
    source_part: String,
    #[pyo3(get)]
    ordinal: usize,
    #[pyo3(get)]
    name: Option<String>,
    #[pyo3(get)]
    enabled: bool,
    #[pyo3(get)]
    calculate_on_exit: bool,
    #[pyo3(get, name = "type")]
    field_type: &'static str,
    #[pyo3(get)]
    value: Py<PyAny>,
    #[pyo3(get)]
    choices: Py<PyTuple>,
}

impl PyLegacyFormField {
    fn from_info(py: Python<'_>, field: rdocx::LegacyFormFieldInfo) -> PyResult<Self> {
        let value = match field.value {
            rdocx::LegacyFormFieldValue::Text(value) => {
                value.into_pyobject(py)?.into_any().unbind()
            }
            rdocx::LegacyFormFieldValue::Checked(value) => {
                value.into_pyobject(py)?.to_owned().into_any().unbind()
            }
            rdocx::LegacyFormFieldValue::SelectedIndex(value) => {
                value.into_pyobject(py)?.into_any().unbind()
            }
        };
        Ok(Self {
            source_part: field.source_part,
            ordinal: field.ordinal,
            name: field.name,
            enabled: field.enabled,
            calculate_on_exit: field.calculate_on_exit,
            field_type: match field.kind {
                rdocx::LegacyFormFieldKind::TextInput => "text",
                rdocx::LegacyFormFieldKind::CheckBox => "checkbox",
                rdocx::LegacyFormFieldKind::DropDownList => "dropdown",
            },
            value,
            choices: PyTuple::new(py, field.choices)?.unbind(),
        })
    }
}

/// Type payload for `Document.set_content_control_type`.
#[pyclass(name = "ContentControlTypePayload", frozen, skip_from_py_object)]
#[derive(Clone)]
pub struct PyContentControlTypePayload {
    inner: rdocx::ContentControlTypePayload,
}

#[pymethods]
impl PyContentControlTypePayload {
    #[staticmethod]
    fn rich_text() -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::RichText,
        }
    }

    #[staticmethod]
    fn plain_text() -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::PlainText,
        }
    }

    #[staticmethod]
    fn date(format: String) -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::Date { format },
        }
    }

    #[staticmethod]
    fn checkbox() -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::CheckBox,
        }
    }

    #[staticmethod]
    fn combo_box(items: Vec<(String, String)>) -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::ComboBox { items },
        }
    }

    #[staticmethod]
    fn drop_down_list(items: Vec<(String, String)>) -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::DropDownList { items },
        }
    }

    #[staticmethod]
    fn picture(relationship_id: String) -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::Picture { relationship_id },
        }
    }

    #[staticmethod]
    fn group() -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::Group,
        }
    }

    #[staticmethod]
    fn repeating_section() -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::RepeatingSection,
        }
    }

    #[staticmethod]
    fn repeating_section_item() -> Self {
        Self {
            inner: rdocx::ContentControlTypePayload::RepeatingSectionItem,
        }
    }
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

fn content_control_kind_from_name(name: &str) -> PyResult<rdocx_oxml::content_control::SdtType> {
    use rdocx_oxml::content_control::SdtType;
    match name {
        "rich_text" => Ok(SdtType::RichText),
        "plain_text" => Ok(SdtType::PlainText),
        "picture" => Ok(SdtType::Picture),
        "checkbox" => Ok(SdtType::CheckBox),
        "combo_box" => Ok(SdtType::ComboBox),
        "dropdown_list" => Ok(SdtType::DropDownList),
        "date" => Ok(SdtType::Date),
        "document_part_list" => Ok(SdtType::DocumentPartList),
        "document_part_object" => Ok(SdtType::DocumentPartObject),
        "group" => Ok(SdtType::Group),
        "repeating_section" => Ok(SdtType::RepeatingSection),
        "repeating_section_item" => Ok(SdtType::RepeatingSectionItem),
        "citation" => Ok(SdtType::Citation),
        "equation" => Ok(SdtType::Equation),
        "bibliography" => Ok(SdtType::Bibliography),
        _ => Err(PyValueError::new_err(format!(
            "unknown content control kind {name:?}"
        ))),
    }
}

fn content_control_placement_from_name(name: &str) -> PyResult<rdocx::ContentControlPlacement> {
    match name {
        "block" => Ok(rdocx::ContentControlPlacement::Block),
        "row" => Ok(rdocx::ContentControlPlacement::Row),
        "cell" => Ok(rdocx::ContentControlPlacement::Cell),
        "paragraph" => Ok(rdocx::ContentControlPlacement::Paragraph),
        "inline" => Ok(rdocx::ContentControlPlacement::Inline),
        _ => Err(PyValueError::new_err(format!(
            "content control placement must be block, row, cell, paragraph, or inline, got {name:?}"
        ))),
    }
}

fn content_control_lock_from_name(name: &str) -> PyResult<rdocx::ContentControlLock> {
    match name {
        "unlocked" => Ok(rdocx::ContentControlLock::Unlocked),
        "sdtLocked" => Ok(rdocx::ContentControlLock::SdtLocked),
        "contentLocked" => Ok(rdocx::ContentControlLock::ContentLocked),
        "sdtContentLocked" => Ok(rdocx::ContentControlLock::SdtContentLocked),
        _ => Err(PyValueError::new_err(format!(
            "unknown content control lock {name:?}"
        ))),
    }
}

fn content_control_appearance_from_name(name: &str) -> PyResult<rdocx::ContentControlAppearance> {
    match name {
        "boundingBox" => Ok(rdocx::ContentControlAppearance::BoundingBox),
        "tags" => Ok(rdocx::ContentControlAppearance::Tags),
        "hidden" => Ok(rdocx::ContentControlAppearance::Hidden),
        _ => Err(PyValueError::new_err(format!(
            "unknown content control appearance {name:?}"
        ))),
    }
}

fn optional_python_bool(value: Option<&Bound<'_, PyAny>>, name: &str) -> PyResult<Option<bool>> {
    let Some(value) = value else {
        return Ok(None);
    };
    if value.is_none() {
        return Ok(None);
    }
    if !value.is_instance_of::<PyBool>() {
        return Err(PyTypeError::new_err(format!(
            "{name} must be a bool or None"
        )));
    }
    value.extract::<bool>().map(Some)
}

#[pyclass(name = "HtmlDiagnostic", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyHtmlDiagnostic {
    pub location: String,
    pub property: Option<String>,
    pub message: String,
}

fn html_diagnostics<'py>(
    py: Python<'py>,
    diagnostics: Vec<rdocx::HtmlDiagnostic>,
) -> PyResult<Bound<'py, PyTuple>> {
    let records = diagnostics
        .into_iter()
        .map(|value| {
            Py::new(
                py,
                PyHtmlDiagnostic {
                    location: value.location,
                    property: value.property,
                    message: value.message,
                },
            )
        })
        .collect::<PyResult<Vec<_>>>()?;
    PyTuple::new(py, records)
}

fn read_import_text(path: PathBuf) -> PyResult<String> {
    use std::io::Read;
    const LIMIT: u64 = 64 * 1024 * 1024;
    let file = std::fs::File::open(path)?;
    if file.metadata()?.len() > LIMIT {
        return Err(PyValueError::new_err(
            "HTML or Markdown input exceeds the 64 MiB limit",
        ));
    }
    let mut bytes = Vec::new();
    file.take(LIMIT + 1).read_to_end(&mut bytes)?;
    if bytes.len() as u64 > LIMIT {
        return Err(PyValueError::new_err(
            "HTML or Markdown input exceeds the 64 MiB limit",
        ));
    }
    String::from_utf8(bytes)
        .map_err(|_| PyValueError::new_err("HTML or Markdown input must be UTF-8"))
}

impl PyDocument {
    fn imported_text<'py>(
        py: Python<'py>,
        text: &str,
        markdown: bool,
        template: Option<&Bound<'py, PyAny>>,
        images: &[(String, Py<PyBytes>, String)],
    ) -> PyResult<(Py<Self>, Bound<'py, PyTuple>)> {
        let resources = images
            .iter()
            .map(|(source, bytes, filename)| rdocx::HtmlImageResource {
                source,
                bytes: bytes.bind(py).as_bytes(),
                filename,
            })
            .collect::<Vec<_>>();
        let loaded;
        let borrowed;
        let template = match template {
            None => None,
            Some(value) if value.is_instance_of::<PyDocument>() => {
                borrowed = value.extract::<PyRef<'_, Self>>()?;
                Some(&borrowed.inner)
            }
            Some(value) => {
                let path = value.extract::<PathBuf>()?;
                loaded = rdocx::Document::open(path).map_err(|error| rdocx_to_pyerr(py, error))?;
                Some(&loaded)
            }
        };
        let result = py
            .detach(|| {
                if markdown {
                    rdocx::Document::from_markdown_with_options(text, template, &resources)
                } else {
                    rdocx::Document::from_html_with_options(text, template, &resources)
                }
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let diagnostics = html_diagnostics(py, result.diagnostics)?;
        Ok((
            Py::new(py, Self::from_document(result.document, py)?)?,
            diagnostics,
        ))
    }

    fn section_child_scope(&self, index: usize) -> PyResult<RevisionState> {
        self.owners
            .iter()
            .find(
                |owner| matches!(owner.location,WordOwnerLocation::Section(value) if value==index),
            )
            .and_then(|owner| owner.content.clone())
            .ok_or_else(|| PyValueError::new_err("section furniture owner is absent"))
    }

    fn prepared_section_child_scope(
        &self,
        prepared: &PreparedWordRevision,
        index: usize,
    ) -> PyResult<RevisionState> {
        self.revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        prepared
            .owners
            .iter()
            .find(
                |owner| matches!(owner.location,WordOwnerLocation::Section(value) if value==index),
            )
            .and_then(|owner| owner.content.clone())
            .ok_or_else(|| PyValueError::new_err("prepared section furniture owner is absent"))
    }

    pub(crate) fn prepared_furniture_scope(
        &self,
        prepared: &PreparedWordRevision,
        scope: Option<&RevisionState>,
    ) -> PyResult<Option<RevisionState>> {
        let Some(scope) = scope else {
            return Ok(None);
        };
        self.revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let index = self
            .owners
            .iter()
            .find_map(|owner| match owner.location {
                WordOwnerLocation::Section(index)
                    if owner
                        .content
                        .as_ref()
                        .is_some_and(|content| content.same_state(scope)) =>
                {
                    Some(index)
                }
                _ => None,
            })
            .ok_or_else(|| {
                PyValueError::new_err("furniture authority has no current section owner")
            })?;
        self.prepared_section_child_scope(prepared, index).map(Some)
    }

    fn indexed_owner(&self, state: &RevisionState) -> Option<(usize, &WordOwner)> {
        note_fx260_word(1);
        let index = *self.indexes.by_generation.get(&state.generation())?;
        let owner = self.owners.get(index)?;
        owner.owner.same_state(state).then_some((index, owner))
    }

    /// The route of the physical owner of a paragraph handle's location,
    /// from the inventory of `prepared` when given, or the current one.
    pub(crate) fn indexed_paragraph_location(
        &self,
        location: &ParagraphLocation,
        prepared: Option<&PreparedWordRevision>,
    ) -> Option<rdocx::ContentLocation> {
        note_fx260_word(1);
        let (owners, indexes) = prepared.map_or((&self.owners, &self.indexes), |prepared| {
            (&prepared.owners, &prepared.indexes)
        });
        let native = match location {
            ParagraphLocation::Body(index) => {
                indexes.body_paragraphs.as_ref()?.get(*index)?.as_ref()?
            }
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => indexes
                .cell_paragraphs
                .get(&(*table, *row, *cell))?
                .get(*paragraph)?,
            ParagraphLocation::Related(location) => location,
        };
        let index = indexes.find(owners, &WordOwnerLocation::Paragraph(native.clone()))?;
        let WordOwnerLocation::Paragraph(location) = &owners[index].location else {
            return None;
        };
        Some(location.clone())
    }

    /// The route of the physical owner of a paragraph handle's location in
    /// the current document, or in `candidate` as `prepared` publishes it:
    /// from that inventory, which holds the route the document resolves, or
    /// as the document resolves it.
    pub(crate) fn paragraph_route(
        &self,
        location: &ParagraphLocation,
        prepared: Option<(&PreparedWordRevision, &rdocx::Document)>,
    ) -> rdocx::Result<rdocx::ContentLocation> {
        let document = prepared.map_or(&self.inner, |(_, candidate)| candidate);
        match self.indexed_paragraph_location(location, prepared.map(|(prepared, _)| prepared)) {
            Some(indexed) => {
                #[cfg(debug_assertions)]
                assert_indexed_route(&indexed, || location.native(document), "paragraph");
                Ok(indexed)
            }
            None => location.native(document),
        }
    }

    /// The route of the physical owner of a table handle's location, from
    /// the inventory of `prepared` when given, or the current one.
    pub(crate) fn indexed_table_location(
        &self,
        index: usize,
        related: &Option<rdocx::ContentLocation>,
        prepared: Option<&PreparedWordRevision>,
    ) -> Option<rdocx::ContentLocation> {
        note_fx260_word(1);
        let (owners, indexes) = prepared.map_or((&self.owners, &self.indexes), |prepared| {
            (&prepared.owners, &prepared.indexes)
        });
        let native = match related {
            Some(location) => location,
            None => indexes.body_tables.get(index)?.as_ref()?,
        };
        let index = indexes.find(owners, &WordOwnerLocation::Table(native.clone()))?;
        let WordOwnerLocation::Table(location) = &owners[index].location else {
            return None;
        };
        Some(location.clone())
    }

    pub(crate) fn indexed_table_owner_location(
        &self,
        state: &RevisionState,
    ) -> Option<rdocx::ContentLocation> {
        let (_, owner) = self.indexed_owner(state)?;
        let WordOwnerLocation::Table(location) = &owner.location else {
            return None;
        };
        Some(location.clone())
    }

    pub(crate) fn indexed_cell_parent_location(
        &self,
        state: &RevisionState,
    ) -> Option<rdocx::ContentLocation> {
        let (_, owner) = self.indexed_owner(state)?;
        let WordOwnerLocation::Cell { table, .. } = &owner.location else {
            return None;
        };
        Some(table.clone())
    }

    pub(crate) fn indexed_cell_location(
        &self,
        table: &rdocx::ContentLocation,
        row: usize,
        cell: usize,
    ) -> Option<WordOwnerLocation> {
        // Cell route equivalence intentionally ignores the story fingerprint.
        // Return the actual current inventory value, never the lookup placeholder.
        let key = WordOwnerLocation::Cell {
            table: table.clone(),
            row,
            cell,
            story: rdocx::StoryId::body(),
        };
        let index = self.indexes.find(&self.owners, &key)?;
        Some(self.owners[index].location.clone())
    }

    pub(crate) fn capture_current_view(
        &self,
        state: &RevisionState,
        segments: smallvec::SmallVec<[PathSeg; 5]>,
        view: WordView,
        furniture: Option<&RevisionState>,
    ) -> PyResult<ContentPath> {
        let (index, _) = self
            .indexed_owner(state)
            .ok_or_else(|| PyValueError::new_err("physical owner is no longer current"))?;
        word_owner_capture(
            &self.revisions,
            &self.owners,
            (self.revisions.current_document_scope(), &self.removal_scope),
            index,
            segments,
            view,
            furniture,
        )
        .map(|(path, _)| path)
    }

    /// The direct body slot of a current paragraph owner, from the inventory.
    pub(crate) fn direct_body_paragraph_slot(&self, state: &RevisionState) -> Option<usize> {
        let (_, owner) = self.indexed_owner(state)?;
        match &owner.location {
            WordOwnerLocation::Paragraph(location)
                if location.story().kind() == rdocx::StoryKind::Body
                    && location.index_path().len() == 1 =>
            {
                owner.body_slot
            }
            _ => None,
        }
    }

    /// The direct body slot of a current run owner's paragraph, and the run
    /// route in it, from the inventory.
    pub(crate) fn direct_body_run_slot(
        &self,
        state: &RevisionState,
    ) -> Option<(usize, rdocx_oxml::text::AcceptedRunPath)> {
        let (_, owner) = self.indexed_owner(state)?;
        let WordOwnerLocation::Run(location, path) = &owner.location else {
            return None;
        };
        let paragraph = self.owners.get(owner.parent?)?;
        match &paragraph.location {
            WordOwnerLocation::Paragraph(parent)
                if parent.story().kind() == rdocx::StoryKind::Body
                    && parent.index_path().len() == 1
                    && parent.index_path() == location.index_path() =>
            {
                Some((paragraph.body_slot?, path.clone()))
            }
            _ => None,
        }
    }

    /// The direct body slot of a current cell owner's table, and the cell
    /// coordinates, from the inventory.
    pub(crate) fn direct_body_cell_slot(
        &self,
        state: &RevisionState,
    ) -> Option<(usize, usize, usize)> {
        let (_, owner) = self.indexed_owner(state)?;
        let WordOwnerLocation::Cell {
            table, row, cell, ..
        } = &owner.location
        else {
            return None;
        };
        let parent = self.owners.get(owner.parent?)?;
        match &parent.location {
            WordOwnerLocation::Table(location)
                if location.story().kind() == rdocx::StoryKind::Body
                    && location.index_path().len() == 1
                    && location.index_path() == table.index_path() =>
            {
                Some((parent.body_slot?, *row, *cell))
            }
            _ => None,
        }
    }

    /// The direct body slot of a current table owner, from the inventory.
    pub(crate) fn direct_body_table_slot(&self, state: &RevisionState) -> Option<usize> {
        let (_, owner) = self.indexed_owner(state)?;
        match &owner.location {
            WordOwnerLocation::Table(location)
                if location.story().kind() == rdocx::StoryKind::Body
                    && location.index_path().len() == 1 =>
            {
                owner.body_slot
            }
            _ => None,
        }
    }

    pub(crate) fn cached_paragraph_text(&self, state: &RevisionState) -> Option<&str> {
        let (_, owner) = self.indexed_owner(state)?;
        match &owner.read_facts {
            WordReadFacts::Paragraph(text) => Some(text),
            _ => None,
        }
    }

    pub(crate) fn publish_run_content_candidate(
        &mut self,
        candidate: rdocx::Document,
        owner: Option<&RevisionState>,
        text: String,
    ) -> PyResult<()> {
        let paragraph = owner
            .map(|state| {
                let (_, run) = self.indexed_owner(state).ok_or_else(|| {
                    PyValueError::new_err("run physical owner is no longer current")
                })?;
                if !matches!(run.location, WordOwnerLocation::Run(..)) {
                    return Err(PyValueError::new_err("run owner has the wrong kind"));
                }
                run.parent
                    .filter(|parent| {
                        self.owners.get(*parent).is_some_and(|owner| {
                            matches!(owner.location, WordOwnerLocation::Paragraph(_))
                        })
                    })
                    .ok_or_else(|| PyValueError::new_err("run has no current paragraph owner"))
            })
            .transpose()?;
        // These publishers retain physical topology and all existing scopes.
        // Candidate text and parent authority are ready before either value moves.
        if let Some(paragraph) = paragraph {
            self.owners[paragraph].read_facts = WordReadFacts::Paragraph(text);
        }
        self.inner = candidate;
        Ok(())
    }

    pub(crate) fn cached_body_paragraph_count(&self) -> Option<usize> {
        Some(self.indexes.body_paragraphs.as_ref()?.len())
    }

    pub(crate) fn cached_body_table_count(&self) -> usize {
        self.indexes.body_tables.len()
    }

    pub(crate) fn cached_table_rows(&self, state: &RevisionState) -> Option<&[Vec<usize>]> {
        let (_, owner) = self.indexed_owner(state)?;
        match &owner.read_facts {
            WordReadFacts::Table(rows) => Some(rows),
            _ => None,
        }
    }

    pub(crate) fn cached_cell_paragraph_count(&self, state: &RevisionState) -> Option<usize> {
        let (_, owner) = self.indexed_owner(state)?;
        let WordOwnerLocation::Cell { row, cell, .. } = &owner.location else {
            return None;
        };
        let parent = self.owners.get(owner.parent?)?;
        let WordReadFacts::Table(rows) = &parent.read_facts else {
            return None;
        };
        rows.get(*row)?.get(*cell).copied()
    }

    pub(crate) fn cached_cell_paragraph(
        &self,
        state: &RevisionState,
        paragraph: usize,
    ) -> Option<rdocx::ContentLocation> {
        let (index, _) = self.indexed_owner(state)?;
        let items = self.indexes.paragraphs_by_parent.get(&index)?;
        // A source/model mismatch stays on the original native admission path.
        if items.len() != self.cached_cell_paragraph_count(state)? {
            return None;
        }
        let WordOwnerLocation::Paragraph(location) = &self.owners[*items.get(paragraph)?].location
        else {
            return None;
        };
        Some(location.clone())
    }

    pub(crate) fn cached_cell_tables(
        &self,
        state: &RevisionState,
    ) -> Option<Vec<rdocx::ContentLocation>> {
        let (index, owner) = self.indexed_owner(state)?;
        if !matches!(owner.location, WordOwnerLocation::Cell { .. }) {
            return None;
        }
        Some(
            self.indexes
                .tables_by_parent
                .get(&index)
                .into_iter()
                .flatten()
                .map(|index| {
                    let WordOwnerLocation::Table(location) = &self.owners[*index].location else {
                        unreachable!("table index contains tables");
                    };
                    location.clone()
                })
                .collect(),
        )
    }

    pub(crate) fn capture_owner(
        &self,
        location: &WordOwnerLocation,
        segments: smallvec::SmallVec<[PathSeg; 5]>,
        view: WordView,
        furniture_scope: Option<&RevisionState>,
    ) -> PyResult<(ContentPath, RevisionState)> {
        word_owner_capture(
            &self.revisions,
            &self.owners,
            (self.revisions.current_document_scope(), &self.removal_scope),
            self.indexes.find(&self.owners, location).ok_or_else(|| {
                PyValueError::new_err("element has no current checked physical owner")
            })?,
            segments,
            view,
            furniture_scope,
        )
    }

    pub(crate) fn capture_prepared_owner(
        &self,
        prepared: &PreparedWordRevision,
        location: &WordOwnerLocation,
        segments: smallvec::SmallVec<[PathSeg; 5]>,
        view: WordView,
        furniture_scope: Option<&RevisionState>,
    ) -> PyResult<(ContentPath, RevisionState)> {
        // Validate origin and expected serial before using any prepared state.
        self.revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let (mut path, owner) = word_owner_capture(
            &self.revisions,
            &prepared.owners,
            (prepared.revision.document_scope(), &self.removal_scope),
            prepared
                .indexes
                .find(&prepared.owners, location)
                .ok_or_else(|| {
                    PyValueError::new_err("element has no current checked physical owner")
                })?,
            segments,
            view,
            furniture_scope,
        )?;
        path.revision = prepared.revision.committed_serial();
        Ok((path, owner))
    }

    pub(crate) fn capture_snapshot(
        &self,
        segments: smallvec::SmallVec<[PathSeg; 5]>,
    ) -> PyResult<ContentPath> {
        self.revisions
            .capture_scopes(
                segments,
                &[
                    self.revisions.current_document_scope(),
                    self.snapshot_scope.clone(),
                ],
            )
            .map_err(revision_to_pyerr)
    }

    pub(crate) fn capture_prepared_snapshot(
        &self,
        prepared: &PreparedWordRevision,
        segments: smallvec::SmallVec<[PathSeg; 5]>,
    ) -> PyResult<ContentPath> {
        self.revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let mut path = self
            .revisions
            .capture_scopes(
                segments,
                &[
                    prepared.revision.document_scope(),
                    prepared.snapshot_scope.clone(),
                ],
            )
            .map_err(revision_to_pyerr)?;
        path.revision = prepared.revision.committed_serial();
        Ok(path)
    }

    pub(crate) fn plain_paragraph_owner(&self, state: &RevisionState) -> Option<(usize, usize)> {
        let selected = *self.indexes.by_generation.get(&state.generation())?;
        let owner = self.owners.get(selected)?;
        if !owner.owner.same_state(state) {
            return None;
        }
        let slot = owner.body_slot?;
        (self.indexes.direct_body_paragraphs.get(&slot) == Some(&selected))
            .then_some((selected, slot))
    }

    fn prepare_tail_body_removal(
        &self,
        slot: usize,
    ) -> PyResult<Option<(std::ops::Range<usize>, PreparedRevision)>> {
        if slot.checked_add(1) != Some(self.inner.content_count()) {
            return Ok(None);
        }
        let Some(&selected) = self.indexes.direct_body_paragraphs.get(&slot) else {
            return Ok(None);
        };
        let mut end = selected + 1;
        while end < self.owners.len() && self.owners[end].parent == Some(selected) {
            note_fx260_word(1);
            end += 1;
        }
        for owner in &self.owners[end..] {
            note_fx260_word(1);
            let independent_run = matches!(owner.location, WordOwnerLocation::Run(..))
                && owner.parent.is_some_and(|parent| parent < selected);
            if !matches!(owner.location, WordOwnerLocation::Section(_)) && !independent_run {
                return Ok(None);
            }
        }
        let revision = self
            .revisions
            .prepare_scopes(
                "Document.remove_content",
                &[self.removal_scope.clone(), self.snapshot_scope.clone()],
                &[RevisionScope::Document, RevisionScope::Text],
            )
            .map_err(revision_to_pyerr)?;
        Ok(Some((selected..end, revision)))
    }

    fn commit_tail_body_removal(
        &mut self,
        slot: usize,
        range: std::ops::Range<usize>,
        revision: PreparedRevision,
    ) -> PyResult<()> {
        let removal_scope = revision.born(0).expect("removal epoch birth");
        let snapshot_scope = revision.born(1).expect("removal snapshot birth");
        self.revisions
            .commit_scopes(revision)
            .map_err(revision_to_pyerr)?;
        let start = range.start;
        let end = range.end;
        for (offset, owner) in self.owners.drain(range).enumerate() {
            note_fx260_word(1);
            self.indexes.by_generation.remove(&owner.owner.generation());
            let hash = owner.location.route_hash();
            let bucket = self
                .indexes
                .by_route
                .get_mut(&hash)
                .expect("removed owner route is indexed");
            bucket.retain(|index| *index != start + offset);
            if bucket.is_empty() {
                self.indexes.by_route.remove(&hash);
            }
            self.indexes.paragraphs_by_parent.remove(&(start + offset));
            self.indexes.tables_by_parent.remove(&(start + offset));
        }
        if let Some(paragraphs) = &mut self.indexes.body_paragraphs {
            paragraphs.pop();
        }
        // Suffix sections and locally born runs of earlier paragraphs have no
        // direct-body route or parent index to shift. Earlier owners stay put.
        for (offset, owner) in self.owners[start..].iter().enumerate() {
            note_fx260_word(1);
            for index in self
                .indexes
                .by_route
                .get_mut(&owner.location.route_hash())
                .expect("surviving suffix route is indexed")
            {
                if *index == end + offset {
                    *index = start + offset;
                }
            }
            *self
                .indexes
                .by_generation
                .get_mut(&owner.owner.generation())
                .expect("surviving suffix owner is indexed") = start + offset;
        }
        self.indexes.direct_body_paragraphs.remove(&slot);
        self.removal_scope = removal_scope;
        self.snapshot_scope = snapshot_scope;
        self.pieces.splice(&self.inner, slot, 1, None);
        Ok(())
    }

    pub(crate) fn current_paragraph_owner(
        &self,
        state: &RevisionState,
    ) -> PyResult<rdocx::ContentLocation> {
        self.owners
            .iter()
            .find_map(|owner| {
                if !owner.owner.same_state(state) {
                    return None;
                }
                match &owner.location {
                    WordOwnerLocation::Paragraph(location) => Some(location.clone()),
                    _ => None,
                }
            })
            .ok_or_else(|| PyValueError::new_err("paragraph physical owner is no longer current"))
            .and_then(|location| {
                ParagraphLocation::Related(location)
                    .native(&self.inner)
                    .map_err(|error| PyValueError::new_err(error.to_string()))
            })
    }

    pub(crate) fn current_run_owner(
        &self,
        state: &RevisionState,
    ) -> PyResult<(rdocx::ContentLocation, rdocx_oxml::text::AcceptedRunPath)> {
        self.owners
            .iter()
            .find_map(|owner| {
                if !owner.owner.same_state(state) {
                    return None;
                }
                match &owner.location {
                    WordOwnerLocation::Run(location, path) => {
                        Some((location.clone(), path.clone()))
                    }
                    _ => None,
                }
            })
            .ok_or_else(|| PyValueError::new_err("run physical owner is no longer current"))
            .and_then(|(location, path)| {
                ParagraphLocation::Related(location)
                    .native(&self.inner)
                    .map(|location| (location, path))
                    .map_err(|error| PyValueError::new_err(error.to_string()))
            })
    }

    pub(crate) fn current_table_owner(
        &self,
        state: &RevisionState,
    ) -> PyResult<rdocx::ContentLocation> {
        self.owners
            .iter()
            .find_map(|owner| {
                if !owner.owner.same_state(state) {
                    return None;
                }
                match &owner.location {
                    WordOwnerLocation::Table(location) => Some(location.clone()),
                    _ => None,
                }
            })
            .ok_or_else(|| PyValueError::new_err("table physical owner is no longer current"))
            .and_then(|location| {
                crate::table::TableLocation::Related(location)
                    .native(&self.inner)
                    .map_err(|error| PyValueError::new_err(error.to_string()))
            })
    }

    pub(crate) fn current_cell_owner(
        &self,
        state: &RevisionState,
    ) -> PyResult<(rdocx::ContentLocation, usize, usize)> {
        self.owners
            .iter()
            .find_map(|owner| {
                if !owner.owner.same_state(state) {
                    return None;
                }
                match &owner.location {
                    WordOwnerLocation::Cell {
                        table, row, cell, ..
                    } => Some((table.clone(), *row, *cell)),
                    _ => None,
                }
            })
            .ok_or_else(|| PyValueError::new_err("cell physical owner is no longer current"))
            .and_then(|(table, row, cell)| {
                crate::table::TableLocation::Related(table)
                    .native(&self.inner)
                    .map(|table| (table, row, cell))
                    .map_err(|error| PyValueError::new_err(error.to_string()))
            })
    }

    pub(crate) fn prepare_plain_paragraph_xml(
        &mut self,
        selected: usize,
        paragraph: &rdocx_oxml::CT_P,
    ) -> PyResult<PreparedPlainParagraphXml> {
        note_fx260_word(1);
        let WordOwnerLocation::Paragraph(location) = &self.owners[selected].location else {
            return Err(PyValueError::new_err("local XML owner is not a paragraph"));
        };
        let mut runs = Vec::new();
        let mut addresses = Vec::new();
        for path in paragraph.accepted_run_paths() {
            let location = WordOwnerLocation::Run(location.clone(), path);
            if let Some(index) = self.indexes.find(&self.owners, &location) {
                runs.push(index);
            } else {
                addresses.push(WordOwnerAddress {
                    location,
                    parent: Some(selected),
                    body_slot: None,
                    projection_ordinal: None,
                    read_facts: WordReadFacts::None,
                });
            }
        }
        let fail =
            |_| pyo3::exceptions::PyMemoryError::new_err("Word local owner allocation failed");
        self.owners.try_reserve(addresses.len()).map_err(fail)?;
        self.indexes
            .by_generation
            .try_reserve(runs.len() + addresses.len())
            .map_err(fail)?;
        self.indexes
            .by_route
            .try_reserve(addresses.len())
            .map_err(fail)?;
        let mut retired = Vec::new();
        if let Some(content) = &self.owners[selected].content {
            retired.push(content.clone());
        }
        if let Some(properties) = &self.owners[selected].properties {
            retired.push(properties.clone());
        }
        for &index in &runs {
            note_fx260_word(1);
            retired.extend(self.owners[index].states());
        }
        retired.push(self.snapshot_scope.clone());
        let mut scopes = retired.iter().map(RevisionState::scope).collect::<Vec<_>>();
        for address in &addresses {
            scopes.extend_from_slice(address.location.scopes());
        }
        let revision = self
            .revisions
            .prepare_local_scopes("Paragraph.replace_xml", &retired, &scopes)
            .map_err(revision_to_pyerr)?;
        let states = (0..scopes.len())
            .map(|index| revision.born(index).expect("prepared paragraph scope"))
            .collect::<Vec<_>>();
        let newborn = word_owners_with_states(addresses, &states[retired.len()..]);
        let mut route_buckets = HashMap::<u64, Vec<usize>>::new();
        for (offset, owner) in newborn.iter().enumerate() {
            let hash = owner.location.route_hash();
            route_buckets
                .entry(hash)
                .or_insert_with(|| {
                    self.indexes
                        .by_route
                        .get(&hash)
                        .cloned()
                        .unwrap_or_default()
                })
                .push(self.owners.len() + offset);
        }
        Ok(PreparedPlainParagraphXml {
            revision,
            states,
            runs,
            newborn,
            route_buckets,
        })
    }

    pub(crate) fn commit_plain_paragraph_xml(
        &mut self,
        selected: usize,
        prepared: PreparedPlainParagraphXml,
        text: String,
    ) -> PyResult<()> {
        let PreparedPlainParagraphXml {
            revision,
            states,
            runs,
            newborn,
            route_buckets,
        } = prepared;
        self.revisions
            .commit_scopes(revision)
            .map_err(revision_to_pyerr)?;
        let mut born = 0;
        if self.owners[selected].content.is_some() {
            self.owners[selected].content = Some(states[born].clone());
            born += 1;
        }
        if self.owners[selected].properties.is_some() {
            self.owners[selected].properties = Some(states[born].clone());
            born += 1;
        }
        for index in runs {
            note_fx260_word(1);
            let owner = &mut self.owners[index];
            self.indexes.by_generation.remove(&owner.owner.generation());
            owner.owner = states[born].clone();
            born += 1;
            if owner.content.is_some() {
                owner.content = Some(states[born].clone());
                born += 1;
            }
            if owner.properties.is_some() {
                owner.properties = Some(states[born].clone());
                born += 1;
            }
            if owner.grid.is_some() {
                owner.grid = Some(states[born].clone());
                born += 1;
            }
            self.indexes
                .by_generation
                .insert(owner.owner.generation(), index);
        }
        self.owners[selected].read_facts = WordReadFacts::Paragraph(text);
        self.snapshot_scope = states[born].clone();
        for owner in newborn {
            self.indexes
                .by_generation
                .insert(owner.owner.generation(), self.owners.len());
            self.owners.push(owner);
        }
        for (hash, bucket) in route_buckets {
            self.indexes.by_route.insert(hash, bucket);
        }
        Ok(())
    }

    pub(crate) fn prepare_word_revision(
        &self,
        py: Python<'_>,
        candidate: &rdocx::Document,
        operation: &'static str,
        publication: WordPublication,
    ) -> PyResult<PreparedWordRevision> {
        let forms = matches!(
            publication,
            WordPublication::FormValues | WordPublication::Signing
        );
        let mut old_sources = FormSources::default();
        let mut new_sources = FormSources::default();
        // A form publication compares the sources that only the rebuild
        // records.
        let (pieces, inventory) = if forms {
            (None, None)
        } else {
            let (pieces, inventory) = self.pieces.assemble(candidate);
            (Some(pieces), inventory)
        };
        let (locations, projections) = match inventory {
            Some((locations, projections)) => (locations, Some(projections)),
            None => (
                word_owner_locations(candidate, forms.then_some(&mut new_sources))
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
                None,
            ),
        };
        if forms {
            word_owner_locations(&self.inner, Some(&mut old_sources))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
        }
        let plan = self.word_revision_plan(
            py,
            candidate,
            &locations,
            &publication,
            (&old_sources, &new_sources),
        )?;
        #[cfg(debug_assertions)]
        if let Some(projections) = &projections {
            assert_word_pieces_rebuild(candidate, &locations, projections);
            self.assert_word_plan_rebuilds(py, candidate, &publication, &plan);
        }
        let WordRevisionPlan {
            retained,
            births,
            retired,
            preserve_snapshot,
        } = plan;
        let revision = if matches!(
            publication,
            WordPublication::ElementXml(_) | WordPublication::RunText(_) | WordPublication::Signing
        ) {
            self.revisions
                .prepare_local_scopes(operation, &retired, &births)
        } else {
            self.revisions.prepare_scopes(operation, &retired, &births)
        }
        .map_err(revision_to_pyerr)?;
        let mut born = 0;
        let mut states = Vec::new();
        for record in retained {
            for state in record {
                states.push(state.unwrap_or_else(|| {
                    let state = revision.born(born).expect("prepared owner birth");
                    born += 1;
                    state
                }));
            }
        }
        let snapshot_scope = if preserve_snapshot {
            self.snapshot_scope.clone()
        } else {
            revision.born(born).expect("prepared snapshot birth")
        };
        let owners = word_owners_with_states(locations, &states);
        let projections = projections.unwrap_or_else(|| WordIndexProjections::of(candidate));
        let indexes = WordOwnerIndexes::prepare(&owners, projections)?;
        Ok(PreparedWordRevision {
            revision,
            owners,
            indexes,
            snapshot_scope,
            pieces,
        })
    }

    /// Which owner states of the current inventory the next inventory
    /// `locations` of `candidate` keeps for `publication`, which it retires,
    /// and the scopes of the states it gives birth to, in order.
    fn word_revision_plan(
        &self,
        py: Python<'_>,
        candidate: &rdocx::Document,
        locations: &[WordOwnerAddress],
        publication: &WordPublication,
        (old_sources, new_sources): (&FormSources, &FormSources),
    ) -> PyResult<WordRevisionPlan> {
        let forms = matches!(
            publication,
            WordPublication::FormValues | WordPublication::Signing
        );
        let selected = match publication {
            WordPublication::HeaderFooterAppend(index) => Some(WordOwnerLocation::Section(*index)),
            WordPublication::ParagraphText(location)
            | WordPublication::ParagraphAppend(location) => {
                Some(WordOwnerLocation::Paragraph(location.clone()))
            }
            WordPublication::TableGrid(location) => {
                Some(WordOwnerLocation::Table(location.clone()))
            }
            WordPublication::CellText(table, row, cell) => {
                self.owners.iter().find_map(|owner| match &owner.location {
                    WordOwnerLocation::Cell {
                        table: owned,
                        row: r,
                        cell: c,
                        ..
                    } if same_content_owner(owned, table) && r == row && c == cell => {
                        Some(owner.location.clone())
                    }
                    _ => None,
                })
            }
            WordPublication::ElementXml(location)
            | WordPublication::RunText(location)
            | WordPublication::HyperlinkRemove(location) => Some(location.clone()),
            _ => None,
        };
        let selected_index = selected.as_ref().and_then(|selected| {
            self.owners
                .iter()
                .position(|owner| owner.location.same_route(selected))
        });
        if !matches!(
            publication,
            WordPublication::Conservative
                | WordPublication::FormValues
                | WordPublication::Signing
                | WordPublication::AppendBody
        ) && selected_index.is_none()
        {
            return Err(PyValueError::new_err(
                "checked XML owner has no current physical authority",
            ));
        }
        let matches = indexed_inventory_matches(&self.owners, locations, publication);
        let form_certificates = if matches!(publication, WordPublication::Signing) {
            signing_owner_certificates(&self.owners, locations, &matches, old_sources, new_sources)
        } else if forms {
            form_owner_certificates(&self.owners, locations, &matches, old_sources, new_sources)
        } else {
            std::collections::HashSet::new()
        };
        let preserve_snapshot = matches!(publication, WordPublication::Signing)
            && self
                .inner
                .stories()
                .map_err(|error| rdocx_to_pyerr(py, error))?
                == candidate
                    .stories()
                    .map_err(|error| rdocx_to_pyerr(py, error))?
            && self
                .inner
                .story_item_snapshots()
                .map_err(|error| rdocx_to_pyerr(py, error))?
                == candidate
                    .story_item_snapshots()
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut retained = Vec::new();
        let mut births = Vec::new();
        for (index, address) in locations.iter().enumerate() {
            let location = &address.location;
            let old = matches[index].map(|before| (before, &self.owners[before]));
            let mut states = Vec::new();
            let mut descendant = false;
            if let (Some((before, _)), Some(selected)) = (old, selected_index) {
                let mut parent = self.owners[before].parent;
                while let Some(index) = parent {
                    if index == selected {
                        descendant = true;
                        break;
                    }
                    parent = self.owners[index].parent;
                }
            }
            for (slot, scope) in location.scopes().iter().enumerate() {
                let route_proved = old.is_some_and(|(before, owner)| {
                    if matches!(publication, WordPublication::AppendBody) {
                        return true;
                    }
                    if Some(before) == selected_index {
                        // A text-box paragraph is rewritten through its physical
                        // part. Its owner scope retires with the edit, so the
                        // handle that made the change cannot be reused.
                        if matches!(
                            &self.owners[before].location,
                            WordOwnerLocation::Paragraph(location)
                                if location.story().kind() == rdocx::StoryKind::TextBox
                        ) {
                            return false;
                        }
                        return true;
                    }
                    let mut root = before;
                    while let Some(parent) = self.owners[root].parent {
                        root = parent;
                    }
                    match &self.owners[root].location {
                        WordOwnerLocation::Section(_) => true,
                        WordOwnerLocation::Hyperlink(..) => matches!(
                            publication,
                            WordPublication::ParagraphAppend(_)
                                | WordPublication::HyperlinkRemove(_)
                        ),
                        WordOwnerLocation::Paragraph(location)
                        | WordOwnerLocation::Table(location) => match location.story().kind() {
                            rdocx::StoryKind::Body => self.owners[root].body_slot.is_some(),
                            rdocx::StoryKind::TextBox | rdocx::StoryKind::Comment => false,
                            rdocx::StoryKind::Header
                            | rdocx::StoryKind::Footer
                            | rdocx::StoryKind::Footnote
                            | rdocx::StoryKind::Endnote => true,
                            _ => false,
                        },
                        _ => {
                            let _ = owner;
                            false
                        }
                    }
                });
                let preserve = old
                    .is_some_and(|(before, _)| form_certificates.contains(&(before, index)))
                    || (route_proved
                        && !matches!(
                            publication,
                            WordPublication::Conservative
                                | WordPublication::FormValues
                                | WordPublication::Signing
                        )
                        && !descendant
                        && !(matches!(publication, WordPublication::HyperlinkRemove(_))
                            && matches!(location, WordOwnerLocation::Paragraph(_))
                            && slot == 1)
                        && !(old.is_some_and(|(before, _)| Some(before) == selected_index)
                            && slot != 0
                            && !matches!(publication, WordPublication::RunText(_))
                            && (!matches!(
                                publication,
                                WordPublication::ParagraphText(_)
                                    | WordPublication::ParagraphAppend(_)
                            ) || slot == 1)));
                let state = if preserve {
                    old.map(|(_, owner)| owner.states()[slot].clone())
                } else {
                    None
                };
                if state.is_none() {
                    births.push(*scope);
                }
                states.push(state);
            }
            retained.push(states);
        }
        let surviving = retained
            .iter()
            .flatten()
            .filter_map(Option::as_ref)
            .map(RevisionState::generation)
            .collect::<std::collections::HashSet<_>>();
        let mut retired = self
            .owners
            .iter()
            .flat_map(WordOwner::states)
            .filter(|state| !surviving.contains(&state.generation()))
            .collect::<Vec<_>>();
        if !preserve_snapshot {
            retired.push(self.snapshot_scope.clone());
            births.push(RevisionScope::Text);
        }
        Ok(WordRevisionPlan {
            retained,
            births,
            retired,
            preserve_snapshot,
        })
    }

    /// Under debug assertions, prove that a revision plans the same
    /// retained states, births and retirements from the rebuilt inventory
    /// of `candidate` as from the inventory read by block and part.
    #[cfg(debug_assertions)]
    fn assert_word_plan_rebuilds(
        &self,
        py: Python<'_>,
        candidate: &rdocx::Document,
        publication: &WordPublication,
        plan: &WordRevisionPlan,
    ) {
        #[cfg(test)]
        let visits = fx260_word_visits();
        let rebuilt = guarded(|| word_owner_locations(candidate, None))
            .expect("an inventory from pieces has a rebuild");
        let sources = FormSources::default();
        let expected = guarded(|| {
            self.word_revision_plan(py, candidate, &rebuilt, publication, (&sources, &sources))
        })
        .expect("a revision from pieces has one from the rebuild");
        let retained = |plan: &WordRevisionPlan| {
            plan.retained
                .iter()
                .map(|states| {
                    states
                        .iter()
                        .map(|state| state.as_ref().map(RevisionState::generation))
                        .collect::<Vec<_>>()
                })
                .collect::<Vec<_>>()
        };
        let retired = |plan: &WordRevisionPlan| {
            plan.retired
                .iter()
                .map(RevisionState::generation)
                .collect::<Vec<_>>()
        };
        assert_eq!(retained(plan), retained(&expected), "retained owner states");
        assert_eq!(plan.births, expected.births, "owner birth order");
        assert_eq!(retired(plan), retired(&expected), "retired owner states");
        assert_eq!(plan.preserve_snapshot, expected.preserve_snapshot);
        #[cfg(test)]
        FX260_WORD_VISITS.with(|cell| cell.set(visits));
    }

    pub(crate) fn commit_word_revision(&mut self, prepared: PreparedWordRevision) -> PyResult<()> {
        self.revisions
            .commit_scopes(prepared.revision)
            .map_err(revision_to_pyerr)?;
        self.owners = prepared.owners;
        self.indexes = prepared.indexes;
        self.snapshot_scope = prepared.snapshot_scope;
        if let Some(pieces) = prepared.pieces {
            self.pieces = pieces;
        }
        Ok(())
    }

    pub(crate) fn publish_owner_candidate(
        &mut self,
        py: Python<'_>,
        candidate: rdocx::Document,
        operation: &'static str,
        publication: WordPublication,
    ) -> PyResult<()> {
        if self
            .inner
            .saves_as(&candidate)
            .map_err(|error| rdocx_to_pyerr(py, error))?
        {
            return Ok(());
        }
        let prepared = self.prepare_word_revision(py, &candidate, operation, publication)?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    /// Keep live handles when a property edit only moves existing physical
    /// addresses. A real insert, removal or replacement keeps the conservative
    /// publication.
    pub(crate) fn adopt_property_edit(
        &mut self,
        py: Python<'_>,
        candidate: rdocx::Document,
        operation: &'static str,
    ) -> PyResult<()> {
        let (pieces, inventory) = self.pieces.assemble(&candidate);
        let (locations, projections) = match inventory {
            Some((locations, projections)) => {
                #[cfg(debug_assertions)]
                assert_word_pieces_rebuild(&candidate, &locations, &projections);
                (locations, Some(projections))
            }
            None => (
                word_owner_locations(&candidate, None)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
                None,
            ),
        };
        // The pieces hold for whichever document holds their values, so the
        // conservative publication below reuses them.
        self.pieces = pieces;
        if let Some(owners) = relocated_property_owners(&self.owners, &locations) {
            let projections = projections.unwrap_or_else(|| WordIndexProjections::of(&candidate));
            let indexes = WordOwnerIndexes::prepare(&owners, projections)?;
            self.owners = owners;
            self.indexes = indexes;
            self.inner = candidate;
            return Ok(());
        }
        self.publish_owner_candidate(py, candidate, operation, WordPublication::Conservative)
    }

    fn remove_note_by_id(
        &mut self,
        py: Python<'_>,
        id: &Bound<'_, PyAny>,
        endnote: bool,
    ) -> PyResult<()> {
        if !id.is_instance_of::<PyInt>() || id.is_instance_of::<PyBool>() {
            return Err(PyTypeError::new_err("note id must be a non-bool integer"));
        }
        let id = id.extract::<i32>()?;
        let mut candidate = self.inner.clone_for_staging();
        if endnote {
            candidate.remove_endnote(id)
        } else {
            candidate.remove_footnote(id)
        }
        .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.remove_note")
    }

    fn edit_content_control(
        &mut self,
        py: Python<'_>,
        location: &PyContentLocation,
        operation: &'static str,
        edit: impl FnOnce(&mut rdocx::Document, &rdocx::ContentLocation) -> rdocx::Result<()> + Send,
    ) -> PyResult<()> {
        let location = location.native(py, self)?;
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| edit(&mut candidate, &location))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.publish_api_candidate(py, candidate, operation)
    }

    fn publish_api_candidate(
        &mut self,
        py: Python<'_>,
        candidate: rdocx::Document,
        operation: &'static str,
    ) -> PyResult<()> {
        if self
            .inner
            .saves_as(&candidate)
            .map_err(|e| rdocx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            operation,
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }
    fn publish_form_candidate(
        &mut self,
        py: Python<'_>,
        candidate: rdocx::Document,
        operation: &'static str,
    ) -> PyResult<()> {
        if self
            .inner
            .saves_as(&candidate)
            .map_err(|e| rdocx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let prepared =
            self.prepare_word_revision(py, &candidate, operation, WordPublication::FormValues)?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }
    fn set_control_value(
        &mut self,
        py: Python<'_>,
        name: &str,
        value: &str,
        by_tag: bool,
    ) -> PyResult<usize> {
        let missing = if by_tag {
            self.inner.content_controls_by_tag(name).is_empty()
        } else {
            self.inner.content_controls_by_alias(name).is_empty()
        };
        if missing {
            return Err(PyKeyError::new_err(name.to_owned()));
        }
        let mut candidate = self.inner.clone_for_staging();
        let count = if by_tag {
            candidate.set_content_control_value_by_tag(name, value)
        } else {
            candidate.set_content_control_value_by_alias(name, value)
        }
        .map_err(|error| match error {
            rdocx::Error::Oxml(rdocx_oxml::OxmlError::InvalidValue(message)) => {
                PyValueError::new_err(message)
            }
            other => rdocx_to_pyerr(py, other),
        })?;
        self.publish_api_candidate(py, candidate, "Document.set_content_control_value")?;
        Ok(count)
    }

    fn create_note(
        slf: &Bound<'_, Self>,
        py: Python<'_>,
        reference: &Bound<'_, PyAny>,
        text: &str,
        endnote: bool,
    ) -> PyResult<Py<PyNote>> {
        let owner = slf.clone().unbind();
        let (location, run_index) = if let Ok(reference) = reference.cast::<PyParagraph>() {
            let paragraph = reference.borrow();
            if !paragraph.belongs_to(py, &owner) {
                return Err(PyValueError::new_err(
                    "note reference belongs to a different document",
                ));
            }
            let location = paragraph.validate(py)?;
            let document = slf.borrow();
            let run_index = location
                .read(&document.inner, |paragraph| paragraph.run_count())
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            (
                location
                    .native(&document.inner)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
                run_index,
            )
        } else if let Ok(reference) = reference.cast::<crate::run::PyRun>() {
            let run = reference.borrow();
            if !run.belongs_to(py, &owner) {
                return Err(PyValueError::new_err(
                    "note reference belongs to a different document",
                ));
            }
            let (location, index) = run.validate(py)?;
            let document = slf.borrow();
            (
                location
                    .native(&document.inner)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
                index + 1,
            )
        } else {
            return Err(PyTypeError::new_err(
                "note reference must be a Paragraph or Run",
            ));
        };
        let mut document = slf.borrow_mut();
        let mut candidate = document.inner.clone_for_staging();
        let position = rdocx::StoryRunPosition {
            location,
            run_index,
        };
        let id = if endnote {
            candidate.create_endnote_at(&position, text)
        } else {
            candidate.create_footnote_at(&position, text)
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_note",
            crate::document::WordPublication::ParagraphAppend(position.location.clone()),
        )?;
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![])
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyNote {
                document: owner,
                id,
                endnote,
                path,
            },
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    fn from_document(inner: rdocx::Document, py: Python<'_>) -> PyResult<Self> {
        let (pieces, inventory) = WordPieces::default().assemble(&inner);
        let (locations, projections) = match inventory {
            Some((locations, projections)) => {
                #[cfg(debug_assertions)]
                assert_word_pieces_rebuild(&inner, &locations, &projections);
                (locations, Some(projections))
            }
            None => (
                word_owner_locations(&inner, None).map_err(|error| rdocx_to_pyerr(py, error))?,
                None,
            ),
        };
        let mut scopes = locations
            .iter()
            .flat_map(|address| address.location.scopes().iter().copied())
            .collect::<Vec<_>>();
        scopes.push(RevisionScope::Text);
        scopes.push(RevisionScope::Document);
        let mut revisions = RevisionCounter::new();
        let states = revisions
            .initialize_scopes(&scopes)
            .map_err(revision_to_pyerr)?;
        let removal_scope = states.last().expect("removal epoch allocated").clone();
        let snapshot_scope = states[states.len() - 2].clone();
        let owners = word_owners_with_states(locations, &states[..states.len() - 2]);
        let projections = projections.unwrap_or_else(|| WordIndexProjections::of(&inner));
        let indexes = WordOwnerIndexes::prepare(&owners, projections)?;
        Ok(Self {
            inner,
            revisions,
            owners,
            snapshot_scope,
            removal_scope,
            indexes,
            pieces,
        })
    }

    /// The live native owner that a `Story` snapshot names.
    ///
    /// A snapshot carries no fingerprint, so it resolves by kind, part name
    /// and owner index against the document as it is now.
    fn native_story(&self, py: Python<'_>, story: &PyStory) -> PyResult<rdocx::StoryId> {
        self.validate_owned_story(py, story)?;
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
    /// A captured snapshot resolves through its current checked physical owner
    /// state, independently of the source Story snapshot's global epoch.
    /// A constructed record resolves only when exactly one current link matches.
    fn native_hyperlink(
        &self,
        py: Python<'_>,
        hyperlink: &PyHyperlink,
    ) -> PyResult<(rdocx::StoryId, usize)> {
        if let Some((path, state)) = &hyperlink.owner {
            if !path.belongs_to(&self.revisions) {
                return Err(PyValueError::new_err(
                    "hyperlink belongs to a different document",
                ));
            }
            path.validate_scopes(
                &self.revisions,
                "hyperlink",
                "re-fetch it with document.hyperlinks.",
            )
            .map_err(|error| crate::stale_to_pyerr(py, error))?;
            let owner = self
                .owners
                .iter()
                .find(|owner| owner.owner.same_state(state))
                .ok_or_else(|| {
                    PyValueError::new_err("hyperlink physical owner is no longer current")
                })?;
            let WordOwnerLocation::Hyperlink(location, _) = &owner.location else {
                return Err(PyValueError::new_err(
                    "hyperlink physical owner has the wrong kind",
                ));
            };
            let story = self
                .inner
                .stories()
                .map_err(|error| rdocx_to_pyerr(py, error))?
                .into_iter()
                .find(|story| same_story_owner(story, location.story()))
                .ok_or_else(|| {
                    PyValueError::new_err("hyperlink story owner is no longer current")
                })?;
            let index = owner
                .projection_ordinal
                .ok_or_else(|| PyValueError::new_err("hyperlink source position is absent"))?;
            return Ok((story, index));
        }
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
        self.validate_owned_item(py, item)?;
        let story = self.native_story(py, &item.story)?;
        if item.binding.is_none()
            && item.story.binding.is_none()
            && item.revision != self.revisions.current()
        {
            return Err(crate::stale_to_pyerr(
                py,
                StaleElementError {
                    element_kind: "story item".to_owned(),
                    captured_revision: item.revision,
                    current_revision: self.revisions.current(),
                    recovery_hint: "Causal provenance is unavailable for this explicit numeric selector. Re-fetch it with document.story_items.".to_owned(),
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
        let mut locations = Vec::with_capacity(indices.len());
        for index in indices {
            if *index == content_count {
                locations.push(rdocx::ContentLocation::end(story.clone()));
                continue;
            }
            // The first body item in that direct child is the child itself.
            let location = self
                .inner
                .direct_body_item_location(*index)
                .map_err(|error| rdocx_to_pyerr(py, error))?
                .filter(|location| location.story() == &story)
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
        document: &std::sync::Arc<Py<PyDocument>>,
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
                story: self.owned_story_snapshot(item.location().story())?,
                kind: story_item_kind_name(item.location().item_kind()).to_owned(),
                index_path: item.location().index_path().to_vec(),
                direct_body_index: item.direct_body_index(),
                text: item.text().map(str::to_owned),
                xml: item.xml().to_vec(),
                revision: self.revisions.current(),
                binding: Some(self.capture_snapshot(smallvec![])?),
                document: Some(document.clone()),
            },
            run_index: position.run_index,
        })
    }

    fn story_item_snapshot(
        &self,
        py: Python<'_>,
        location: &rdocx::ContentLocation,
        document: std::sync::Arc<Py<PyDocument>>,
    ) -> PyResult<PyStoryItem> {
        let item = self
            .inner
            .standalone_story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .find(|item| item.location() == location)
            .ok_or_else(|| PyIndexError::new_err("inserted story item was not found"))?;
        Ok(PyStoryItem {
            story: self.owned_story_snapshot(item.location().story())?,
            kind: story_item_kind_name(item.location().item_kind()).to_owned(),
            index_path: item.location().index_path().to_vec(),
            direct_body_index: item.direct_body_index(),
            text: item.text().map(str::to_owned),
            xml: item.xml().to_vec(),
            revision: self.revisions.current(),
            binding: Some(self.capture_snapshot(smallvec![])?),
            document: Some(document),
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
        let owner = std::sync::Arc::new(paragraph.document.clone_ref(py));
        let [_, _] = location.index_path() else {
            return document.story_item_snapshot(py, &location, owner);
        };
        let item = document
            .inner
            .story_range_paragraph_snapshot(&location)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(PyStoryItem {
            story: document.owned_story_snapshot(item.location().story())?,
            kind: story_item_kind_name(item.location().item_kind()).to_owned(),
            index_path: item.location().index_path().to_vec(),
            direct_body_index: item.direct_body_index(),
            text: item.text().map(str::to_owned),
            xml: item.xml().to_vec(),
            revision: document.revisions.current(),
            binding: Some(document.capture_snapshot(smallvec![])?),
            document: Some(owner),
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
                | crate::paragraph::ParagraphLocation::Related(_) => {
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
            return table.direct_body_index(py);
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
        let mut candidate = self.inner.clone_for_staging();
        let boundary = candidate
            .split_run(body_index, run_index, character_offset)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if run_count(&candidate) != before {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.split_run",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(boundary)
    }

    /// Run a native mutation that reports how many things it changed.
    ///
    /// The GIL is released while it runs, and live handles are staled only
    /// when the count is nonzero.
    fn counted_mutation<F>(
        &mut self,
        py: Python<'_>,
        operation: &'static str,
        mutation: F,
    ) -> PyResult<usize>
    where
        F: FnOnce(&mut rdocx::Document) -> rdocx::Result<usize> + Send,
    {
        let mut candidate = self.inner.clone_for_staging();
        let count = py
            .detach(|| mutation(&mut candidate))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if count > 0 {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                operation,
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
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
    ) -> PyResult<(rdocx::Document, Vec<usize>)> {
        let mut candidate = self.inner.clone_for_staging();
        let counts = py
            .detach(|| candidate.try_replace_all_expected(pairs))
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .map_err(|mismatch| crate::replacement_count_to_pyerr(py, &mismatch, batch))?;
        Ok((candidate, counts))
    }

    // Both checked item and physical cell routes share singular count/error and
    // publication policy. The concrete closures are the two actual consumers.
    pub(crate) fn scoped_replacement<F>(
        &mut self,
        py: Python<'_>,
        operation: &'static str,
        mutation: F,
    ) -> PyResult<usize>
    where
        F: FnOnce(
                &mut rdocx::Document,
            ) -> rdocx::Result<Result<usize, rdocx::ReplacementCountMismatch>>
            + Send,
    {
        let mut candidate = self.inner.clone_for_staging();
        let count = py
            .detach(|| mutation(&mut candidate))
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .map_err(|mismatch| crate::replacement_count_to_pyerr(py, &mismatch, false))?;
        if count > 0 {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                operation,
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
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
        binding: None,
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

pub(crate) fn checked_section_length(
    value: Option<&Bound<'_, PyAny>>,
    name: &str,
) -> PyResult<Option<i64>> {
    value
        .map(|value| {
            if value.is_instance_of::<PyBool>() {
                return Err(PyTypeError::new_err(format!(
                    "{name} must be an integer EMU length, not bool"
                )));
            }
            value.extract::<i64>()
        })
        .transpose()
}

// Public update_section has an explicit-unit contract. Other consumers of
// checked_section_length, including live Section setters, keep their EMU API.
fn checked_update_section_length(
    py: Python<'_>,
    value: Option<&Bound<'_, PyAny>>,
    name: &str,
) -> PyResult<Option<i64>> {
    if let Some(value) = value
        && !value.is_instance_of::<pyo3::types::PyInt>()
    {
        return Err(PyTypeError::new_err(format!(
            "{name} requires Length or zero; use Twips or Emu"
        )));
    }
    let emu = checked_section_length(value, name)?;
    if let Some(value) = value
        && emu != Some(0)
        && !value.is_instance(&py.import("rdocx.shared")?.getattr("Length")?)?
    {
        return Err(PyValueError::new_err(format!(
            "update_section {name} nonzero geometry requires explicit Length; use Twips(1440) or Emu(value)"
        )));
    }
    Ok(emu)
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
    #[getter]
    fn default_font_name(&self) -> Option<String> {
        self.inner.default_font_name().map(str::to_owned)
    }

    #[setter]
    fn set_default_font_name(&mut self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.inner
            .set_default_font_name_value(value)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[getter]
    fn default_font_size(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.inner
            .default_font_size()
            .map(|value| crate::length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_default_font_size(&mut self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        let value = value.map(rdocx::Length::emu);
        if value.is_some_and(|value| {
            value.to_emu() < 0 || u32::try_from(value.to_emu() / 6350).is_err()
        }) {
            return Err(PyValueError::new_err(
                "Document.default_font_size must fit nonnegative half-points",
            ));
        }
        self.inner
            .set_default_font_size_value(value)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[new]
    #[pyo3(signature = (path = None))]
    fn new(path: Option<&Bound<'_, PyAny>>, py: Python<'_>) -> PyResult<Self> {
        let inner = match path {
            Some(path) => open_document_input(py, path)?,
            None => rdocx::Document::new(),
        };
        Self::from_document(inner, py)
    }

    #[staticmethod]
    fn open(path: &Bound<'_, PyAny>, py: Python<'_>) -> PyResult<Self> {
        Self::from_document(open_document_input(py, path)?, py)
    }

    #[staticmethod]
    fn from_bytes(bytes: &[u8], py: Python<'_>) -> PyResult<Self> {
        rdocx::Document::from_bytes(bytes)
            .map_err(|error| rdocx_to_pyerr(py, error))
            .and_then(|inner| Self::from_document(inner, py))
    }

    #[staticmethod]
    #[pyo3(signature = (html, *, images=Vec::new()))]
    fn from_html<'py>(
        py: Python<'py>,
        html: &str,
        images: Vec<(String, Py<PyBytes>, String)>,
    ) -> PyResult<(Py<Self>, Bound<'py, PyTuple>)> {
        Self::imported_text(py, html, false, None, &images)
    }

    #[staticmethod]
    #[pyo3(signature = (path, *, images=Vec::new()))]
    fn open_html<'py>(
        py: Python<'py>,
        path: PathBuf,
        images: Vec<(String, Py<PyBytes>, String)>,
    ) -> PyResult<(Py<Self>, Bound<'py, PyTuple>)> {
        Self::imported_text(py, &read_import_text(path)?, false, None, &images)
    }

    #[staticmethod]
    #[pyo3(signature = (markdown, *, template=None, images=Vec::new()))]
    fn from_markdown<'py>(
        py: Python<'py>,
        markdown: &str,
        template: Option<&Bound<'py, PyAny>>,
        images: Vec<(String, Py<PyBytes>, String)>,
    ) -> PyResult<(Py<Self>, Bound<'py, PyTuple>)> {
        Self::imported_text(py, markdown, true, template, &images)
    }

    #[staticmethod]
    #[pyo3(signature = (path, *, template=None, images=Vec::new()))]
    fn open_markdown<'py>(
        py: Python<'py>,
        path: PathBuf,
        template: Option<&Bound<'py, PyAny>>,
        images: Vec<(String, Py<PyBytes>, String)>,
    ) -> PyResult<(Py<Self>, Bound<'py, PyTuple>)> {
        Self::imported_text(py, &read_import_text(path)?, true, template, &images)
    }

    #[pyo3(signature = (html, *, after, images=Vec::new()))]
    fn insert_html<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        html: &str,
        after: PyRef<'_, PyParagraph>,
        images: Vec<(String, Py<PyBytes>, String)>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        if !after.belongs_to(py, &slf) {
            return Err(PyValueError::new_err(
                "after must be a live paragraph of this document",
            ));
        }
        let paragraph = after.validate(py)?;
        let mut document = slf.borrow_mut(py);
        let location = paragraph
            .native(&document.inner)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let snapshots = document
            .inner
            .story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let direct = snapshots
            .iter()
            .filter(|item| {
                item.location().story() == location.story()
                    && item.is_direct_child()
                    && (location.story().kind() != rdocx::StoryKind::Body
                        || item.direct_body_index().is_some())
            })
            .collect::<Vec<_>>();
        let index = direct
            .iter()
            .position(|item| item.location() == &location)
            .ok_or_else(|| {
                PyValueError::new_err("after must be a supported direct-child paragraph")
            })?;
        let destination = direct.get(index + 1).map_or_else(
            || rdocx::ContentLocation::end(location.story().clone()),
            |item| item.location().clone(),
        );
        let resources = images
            .iter()
            .map(|(source, bytes, filename)| rdocx::HtmlImageResource {
                source,
                bytes: bytes.bind(py).as_bytes(),
                filename,
            })
            .collect::<Vec<_>>();
        let mut candidate = document.inner.clone_for_staging();
        let result = py
            .detach(|| candidate.insert_html_fragment(&destination, html, &resources))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let diagnostics = html_diagnostics(py, result.diagnostics)?;
        document.publish_api_candidate(py, candidate, "Document.insert_html")?;
        Ok(diagnostics)
    }

    /// Sign current typed content with DER or one unencrypted PKCS#8/X.509 PEM block.
    fn sign(
        &mut self,
        private_key: &[u8],
        certificate: &[u8],
        py: Python<'_>,
    ) -> PyResult<Py<pyo3::types::PyDict>> {
        let mut candidate = self.inner.clone_for_staging();
        let report = py
            .detach(|| candidate.sign(private_key, certificate))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let result = oxml_py_support::signature_report_to_py(py, &report)?;
        let prepared =
            self.prepare_word_revision(py, &candidate, "Document.sign", WordPublication::Signing)?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(result)
    }

    /// Verify serialized current content, without publishing state or asserting chain trust.
    fn verify_signatures(&self, py: Python<'_>) -> PyResult<Py<PyTuple>> {
        let reports = py
            .detach(|| {
                let mut current = self.inner.clone_for_staging();
                let bytes = current.to_bytes()?;
                rdocx::Document::from_bytes(&bytes)?
                    .verify_signatures()
                    .map_err(rdocx::Error::from)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let reports = reports
            .iter()
            .map(|report| oxml_py_support::signature_report_to_py(py, report))
            .collect::<PyResult<Vec<_>>>()?;
        Ok(PyTuple::new(py, reports)?.unbind())
    }

    fn save(&mut self, path: &Bound<'_, PyAny>, py: Python<'_>) -> PyResult<()> {
        if let Ok(path) = path.extract::<PathBuf>() {
            if path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("pdf"))
            {
                return Err(PyValueError::new_err(
                    "use to_pdf() to export a PDF instead of save()",
                ));
            }
            return py
                .detach(|| self.inner.save(path))
                .map_err(|error| rdocx_to_pyerr(py, error));
        }
        let write = path.getattr("write").map_err(|_| PyTypeError::new_err(
            "save requires a filesystem path or binary writable stream; use stream.write(document.to_bytes())",
        ))?;
        let bytes = self
            .inner
            .to_bytes()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let written = write.call1((PyBytes::new(py, &bytes),))?;
        if written.extract::<usize>().ok() != Some(bytes.len()) {
            return Err(pyo3::exceptions::PyOSError::new_err(
                "binary stream write returned a short or invalid byte count",
            ));
        }
        Ok(())
    }

    #[pyo3(name = "to_bytes")]
    fn serialize<'py>(&mut self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        py.detach(|| self.inner.to_bytes())
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn to_html(&self, py: Python<'_>) -> String {
        py.detach(|| self.inner.to_html())
    }
    fn to_markdown(&self, py: Python<'_>) -> String {
        py.detach(|| self.inner.to_markdown())
    }
    fn text(&self, py: Python<'_>) -> String {
        py.detach(|| self.inner.text_report().text)
    }
    fn text_report(&self, py: Python<'_>) -> PyTextReport {
        py.detach(|| self.inner.text_report()).into()
    }
    fn word_count(&self) -> usize {
        self.inner.word_count()
    }
    #[pyo3(signature = (*, include_spaces = true))]
    fn character_count(&self, include_spaces: bool) -> usize {
        self.inner.character_count(include_spaces)
    }
    fn validate(&self, py: Python<'_>) -> PyResult<PyValidationReport> {
        py.detach(|| self.inner.validate())
            .map(Into::into)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }
    fn to_odt(&self, py: Python<'_>) -> PyResult<PyOdtExportResult> {
        let result = py
            .detach(|| self.inner.to_odt_bytes())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(PyOdtExportResult {
            bytes: result.bytes,
            diagnostics: result
                .diagnostics
                .into_iter()
                .map(|d| PyOdtDiagnostic {
                    path: d.path,
                    message: d.message,
                })
                .collect(),
        })
    }
    fn to_rtf(&self, py: Python<'_>) -> PyResult<PyRtfExportResult> {
        let result = py
            .detach(|| self.inner.to_rtf_bytes())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(PyRtfExportResult {
            bytes: result.bytes,
            diagnostics: result
                .diagnostics
                .into_iter()
                .map(|d| PyRtfDiagnostic {
                    offset: d.offset,
                    destination: d.destination,
                    message: d.message,
                })
                .collect(),
        })
    }
    fn to_epub(&self, py: Python<'_>) -> PyResult<PyEpubExportResult> {
        let result = py
            .detach(|| self.inner.to_epub_bytes())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        Ok(PyEpubExportResult {
            bytes: result.bytes,
            diagnostics: result
                .diagnostics
                .into_iter()
                .map(|d| PyEpubDiagnostic {
                    path: d.path,
                    message: d.message,
                })
                .collect(),
        })
    }
    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None))]
    fn page_count<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<usize> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        py.detach(|| match refs {
            None => self
                .inner
                .layout_deterministic()
                .map(|layout| layout.layout.pages.len()),
            Some(ref fonts) => self
                .inner
                .layout_with_additional_fonts_and_options(fonts, rdocx::RenderOptions::default())
                .map(|layout| layout.layout.pages.len()),
        })
        .map_err(|error| rdocx_to_pyerr(py, error))
    }
    #[pyo3(signature = (data, *, grammar_version=None))]
    fn render_template(
        &mut self,
        py: Python<'_>,
        data: &Bound<'_, PyAny>,
        grammar_version: Option<u32>,
    ) -> PyResult<usize> {
        if !data.is_instance_of::<PyDict>() {
            return Err(PyTypeError::new_err(
                "render_template data must be a dict with string keys",
            ));
        }
        let data = template_value(data, "data", &mut Vec::new())?;
        let before = self
            .inner
            .to_bytes()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut candidate = self.inner.clone_for_staging();
        let count = py
            .detach(|| match grammar_version {
                None => candidate.render_template(&data),
                Some(version) => candidate.render_template_versioned(&data, version),
            })
            .map_err(|error| template_render_pyerr(py, error))?;
        let after = candidate
            .to_bytes()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if after != before {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.render_template",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(count)
    }

    fn bind_content_controls(
        &mut self,
        py: Python<'_>,
        values: &Bound<'_, PyDict>,
    ) -> PyResult<Py<PyInt>> {
        let values = values
            .iter()
            .map(|(key, value)| Ok((key.extract::<String>()?, value.extract::<String>()?)))
            .collect::<PyResult<HashMap<_, _>>>()?;
        let mut candidate = self.inner.clone_for_staging();
        let count = py
            .detach(|| candidate.bind_content_controls(&values))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let result = count.into_pyobject(py)?.unbind();
        self.publish_form_candidate(py, candidate, "Document.bind_content_controls")?;
        Ok(result)
    }

    fn legacy_form_fields<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let fields = self
            .inner
            .legacy_form_fields()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let fields = fields
            .into_iter()
            .map(|field| Py::new(py, PyLegacyFormField::from_info(py, field)?))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, fields)
    }

    #[pyo3(signature = (name, value, *, source_part=None, ordinal=None))]
    fn set_legacy_form_field_value(
        &mut self,
        py: Python<'_>,
        name: Option<&str>,
        value: &Bound<'_, PyAny>,
        source_part: Option<&str>,
        ordinal: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<Py<PyLegacyFormField>> {
        let ordinal = ordinal
            .map(|value| {
                if !value.is_instance_of::<PyInt>() || value.is_instance_of::<PyBool>() {
                    return Err(PyTypeError::new_err(
                        "legacy field ordinal must be a nonnegative int, not bool",
                    ));
                }
                value.extract::<usize>()
            })
            .transpose()?;
        if source_part.is_some() != ordinal.is_some() {
            return Err(PyValueError::new_err(
                "source_part and ordinal must be supplied together",
            ));
        }
        if source_part.is_none() && name.is_none() {
            return Err(PyValueError::new_err(
                "name is required without source_part and ordinal",
            ));
        }
        let fields = self
            .inner
            .legacy_form_fields()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut matches = fields.into_iter().filter(|field| {
            if let (Some(part), Some(ordinal)) = (source_part, ordinal) {
                field.source_part == part
                    && field.ordinal == ordinal
                    && name.is_none_or(|name| field.name.as_deref() == Some(name))
            } else {
                field.name.as_deref() == name
            }
        });
        let field = matches
            .next()
            .ok_or_else(|| PyValueError::new_err("unknown or stale legacy form field identity"))?;
        if matches.next().is_some() {
            return Err(PyValueError::new_err(
                "ambiguous legacy form field name; supply source_part and ordinal",
            ));
        }
        let value = match field.kind {
            rdocx::LegacyFormFieldKind::TextInput => {
                rdocx::LegacyFormFieldValue::Text(value.extract::<String>()?)
            }
            rdocx::LegacyFormFieldKind::CheckBox => {
                if !value.is_instance_of::<PyBool>() {
                    return Err(PyTypeError::new_err("legacy checkbox requires bool"));
                }
                rdocx::LegacyFormFieldValue::Checked(value.extract::<bool>()?)
            }
            rdocx::LegacyFormFieldKind::DropDownList => {
                if !value.is_instance_of::<PyInt>() || value.is_instance_of::<PyBool>() {
                    return Err(PyTypeError::new_err(format!(
                        "legacy dropdown received {}. Use a nonnegative integer index, not bool, from document.legacy_form_fields() snapshot choices",
                        value.get_type().name()?,
                    )));
                }
                rdocx::LegacyFormFieldValue::SelectedIndex(value.extract::<usize>()?)
            }
        };
        let mut candidate = self.inner.clone_for_staging();
        let reopened = py
            .detach(|| {
                candidate.set_legacy_form_field_value(&field.source_part, field.ordinal, value)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let result = Py::new(py, PyLegacyFormField::from_info(py, reopened)?)?;
        self.publish_form_candidate(py, candidate, "Document.set_legacy_form_field_value")?;
        Ok(result)
    }

    fn content_controls<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.inner
                .content_controls()
                .into_iter()
                .map(PyContentControl::from),
        )
    }
    fn set_content_control_value_by_tag(
        &mut self,
        py: Python<'_>,
        tag: &str,
        value: &str,
    ) -> PyResult<usize> {
        self.set_control_value(py, tag, value, true)
    }
    fn fill_content_control_tags<'py>(
        &mut self,
        py: Python<'py>,
        pairs: Vec<(String, String)>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let pairs = pairs
            .iter()
            .map(|(tag, value)| (tag.as_str(), value.as_str()))
            .collect::<Vec<_>>();
        let mut candidate = self.inner.clone_for_staging();
        let counts = py
            .detach(|| candidate.fill_content_control_tags(&pairs))
            .map_err(|error| rdocx_to_pyerr(py, error))?;

        let result = PyTuple::new(py, counts)?;
        self.publish_api_candidate(py, candidate, "Document.fill_content_control_tags")?;

        Ok(result)
    }

    #[allow(clippy::too_many_arguments)] // Q21 keeps each public operand explicit.
    #[pyo3(signature = (destination, placement, kind, tag=None, alias=None, lock=None, appearance=None, placeholder=None))]
    fn insert_content_control(
        &mut self,
        py: Python<'_>,
        destination: PyRef<'_, PyContentLocation>,
        placement: &str,
        kind: &str,
        tag: Option<String>,
        alias: Option<String>,
        lock: Option<String>,
        appearance: Option<String>,
        placeholder: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let destination = destination.native(py, self)?;
        let spec = rdocx::ContentControlSpec {
            placement: content_control_placement_from_name(placement)?,
            kind: content_control_kind_from_name(kind)?,
            tag,
            alias,
            lock: lock
                .as_deref()
                .map(content_control_lock_from_name)
                .transpose()?,
            appearance: appearance
                .as_deref()
                .map(content_control_appearance_from_name)
                .transpose()?,
            placeholder: optional_python_bool(placeholder, "placeholder")?,
        };
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.insert_content_control(&destination, spec))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.publish_api_candidate(py, candidate, "Document.insert_content_control")
    }

    #[pyo3(signature = (location, tag=None))]
    fn set_content_control_tag(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        tag: Option<String>,
    ) -> PyResult<()> {
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_tag",
            |document, location| document.set_content_control_tag(location, tag.as_deref()),
        )
    }

    #[pyo3(signature = (location, alias=None))]
    fn set_content_control_alias(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        alias: Option<String>,
    ) -> PyResult<()> {
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_alias",
            |document, location| document.set_content_control_alias(location, alias.as_deref()),
        )
    }

    #[pyo3(signature = (location, lock=None))]
    fn set_content_control_lock(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        lock: Option<String>,
    ) -> PyResult<()> {
        let lock = lock
            .as_deref()
            .map(content_control_lock_from_name)
            .transpose()?;
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_lock",
            |document, location| document.set_content_control_lock(location, lock),
        )
    }

    #[pyo3(signature = (location, appearance=None))]
    fn set_content_control_appearance(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        appearance: Option<String>,
    ) -> PyResult<()> {
        let appearance = appearance
            .as_deref()
            .map(content_control_appearance_from_name)
            .transpose()?;
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_appearance",
            |document, location| document.set_content_control_appearance(location, appearance),
        )
    }

    #[pyo3(signature = (location, placeholder=None))]
    fn set_content_control_placeholder(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        placeholder: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let placeholder = optional_python_bool(placeholder, "placeholder")?;
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_placeholder",
            |document, location| document.set_content_control_placeholder(location, placeholder),
        )
    }

    fn remove_content_control(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
    ) -> PyResult<()> {
        self.edit_content_control(
            py,
            &location,
            "Document.remove_content_control",
            |document, location| document.remove_content_control(location),
        )
    }

    fn unwrap_content_control(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
    ) -> PyResult<()> {
        self.edit_content_control(
            py,
            &location,
            "Document.unwrap_content_control",
            |document, location| document.unwrap_content_control(location),
        )
    }

    fn move_content_control(
        &mut self,
        py: Python<'_>,
        source: PyRef<'_, PyContentLocation>,
        destination: PyRef<'_, PyContentLocation>,
    ) -> PyResult<()> {
        let source = source.native(py, self)?;
        let destination = destination.native(py, self)?;
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.move_content_control(&source, &destination))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.publish_api_candidate(py, candidate, "Document.move_content_control")
    }

    fn clone_content_control(
        &mut self,
        py: Python<'_>,
        source: PyRef<'_, PyContentLocation>,
        destination: PyRef<'_, PyContentLocation>,
    ) -> PyResult<()> {
        let source = source.native(py, self)?;
        let destination = destination.native(py, self)?;
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.clone_content_control(&source, &destination))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.publish_api_candidate(py, candidate, "Document.clone_content_control")
    }

    fn set_content_control_type(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        payload: PyRef<'_, PyContentControlTypePayload>,
    ) -> PyResult<()> {
        let payload = payload.inner.clone();
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_type",
            move |document, location| document.set_content_control_type(location, payload),
        )
    }

    fn repeat_content_control(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
    ) -> PyResult<()> {
        self.edit_content_control(
            py,
            &location,
            "Document.repeat_content_control",
            |document, location| document.repeat_content_control(location),
        )
    }

    #[pyo3(signature = (item_xml, schema_uris=Vec::new()), text_signature = "($self, item_xml, schema_uris=())")]
    fn add_custom_xml_store(
        &mut self,
        py: Python<'_>,
        item_xml: &[u8],
        schema_uris: Vec<String>,
    ) -> PyResult<String> {
        let mut candidate = self.inner.clone_for_staging();
        let item_id = py
            .detach(|| {
                let uris = schema_uris.iter().map(String::as_str).collect::<Vec<_>>();
                candidate.add_custom_xml_store(item_xml, &uris)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        self.publish_api_candidate(py, candidate, "Document.add_custom_xml_store")?;
        Ok(item_id)
    }

    #[pyo3(signature = (location, store_item_id, xpath, prefix_mappings=None))]
    fn set_content_control_binding(
        &mut self,
        py: Python<'_>,
        location: PyRef<'_, PyContentLocation>,
        store_item_id: String,
        xpath: String,
        prefix_mappings: Option<String>,
    ) -> PyResult<()> {
        self.edit_content_control(
            py,
            &location,
            "Document.set_content_control_binding",
            move |document, location| {
                document.set_content_control_binding(
                    location,
                    &store_item_id,
                    &xpath,
                    prefix_mappings.as_deref(),
                )
            },
        )
    }

    fn set_content_control_value_by_alias(
        &mut self,
        py: Python<'_>,
        alias: &str,
        value: &str,
    ) -> PyResult<usize> {
        self.set_control_value(py, alias, value, false)
    }
    #[pyo3(signature = (destination, fragment, *, policy=None))]
    fn import_fragment(
        &mut self,
        py: Python<'_>,
        destination: PyRef<'_, PyContentLocation>,
        fragment: PyRef<'_, PyDocumentFragment>,
        policy: Option<PyRef<'_, PyFragmentConflictPolicy>>,
    ) -> PyResult<()> {
        let destination = destination.native(py, self)?;
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .import_fragment(
                &destination,
                &fragment.inner,
                policy.map(|p| p.inner).unwrap_or_default(),
            )
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.import_fragment")
    }
    #[pyo3(signature = (other, *, at, policy=None))]
    fn insert_document(
        &mut self,
        py: Python<'_>,
        other: PyRef<'_, PyDocument>,
        at: &Bound<'_, PyAny>,
        policy: Option<PyRef<'_, PyFragmentConflictPolicy>>,
    ) -> PyResult<()> {
        if at.is_instance_of::<PyBool>() {
            return Err(PyTypeError::new_err(
                "at must be a physical body index, not bool",
            ));
        }
        let at = at.extract::<usize>()?;
        let destination = self.body_location(py, at)?;
        if other.inner.content_count() == 0 {
            return Ok(());
        }
        let start = other.body_location(py, 0)?;
        let end = rdocx::ContentLocation::end(other.body_story(py)?);
        let fragment = rdocx::DocumentFragment::from_range(&other.inner, &start, &end, false)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .import_fragment(
                &destination,
                &fragment,
                policy.map(|p| p.inner).unwrap_or_default(),
            )
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.insert_document")
    }
    #[getter]
    fn custom_properties(slf: Py<Self>) -> PyCustomProperties {
        PyCustomProperties { document: slf }
    }
    #[getter]
    fn app_properties(slf: Py<Self>) -> PyAppProperties {
        PyAppProperties { document: slf }
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

    #[pyo3(signature = (paragraph, heading, text, *, tooltip=None))]
    fn add_heading_hyperlink(
        slf: Py<Self>,
        py: Python<'_>,
        paragraph: PyRef<'_, PyParagraph>,
        heading: PyRef<'_, PyParagraph>,
        text: &str,
        tooltip: Option<&str>,
    ) -> PyResult<Py<crate::run::PyRun>> {
        if !paragraph.belongs_to(py, &slf) || !heading.belongs_to(py, &slf) {
            return Err(PyValueError::new_err(
                "heading hyperlink paragraphs belong to a different document",
            ));
        }
        let source = paragraph.validate(py)?;
        let destination = heading.validate(py)?;
        let mut document = slf.borrow_mut(py);
        let source = source
            .native(&document.inner)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let destination = destination
            .native(&document.inner)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut candidate = document.inner.clone_for_staging();
        let position = candidate
            .add_heading_hyperlink_at(&source, &destination, text, tooltip)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let run_path = candidate
            .read_story_paragraph(&position.location, |p| p.run_path(position.run_index))
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("heading link run was not found"))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_heading_hyperlink",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(
                &prepared.revision,
                smallvec![
                    PathSeg::Para(*position.location.index_path().last().unwrap_or(&0)),
                    PathSeg::Run(position.run_index)
                ],
            )
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            crate::run::PyRun::new(slf.clone_ref(py), path, run_path, Some(position.location))
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    /// Frozen snapshots distinguish physical drawings of the same media.
    fn pictures<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let values = self
            .inner
            .pictures()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .map(|inner| {
                Py::new(
                    py,
                    PyPicture {
                        inner,
                        binding: self.capture_snapshot(smallvec![])?,
                    },
                )
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, values)
    }

    fn set_picture_size(
        &mut self,
        py: Python<'_>,
        picture: PyRef<'_, PyPicture>,
        width: &Bound<'_, PyAny>,
        height: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        picture
            .binding
            .validate_scopes(
                &self.revisions,
                "picture",
                "Re-fetch it with document.pictures().",
            )
            .map_err(|error| crate::stale_to_pyerr(py, error))?;
        if picture.binding.revision != self.revisions.current() {
            return Err(crate::stale_to_pyerr(
                py,
                StaleElementError {
                    element_kind: "picture".into(),
                    captured_revision: picture.binding.revision,
                    current_revision: self.revisions.current(),
                    recovery_hint: "Re-fetch it with document.pictures().".into(),
                },
            ));
        }
        let width = checked_section_length(Some(width), "picture width")?.unwrap();
        let height = checked_section_length(Some(height), "picture height")?.unwrap();
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_picture_size_at(
                &picture.inner,
                rdocx::Length::emu(width),
                rdocx::Length::emu(height),
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if (width, height) == (picture.inner.width, picture.inner.height) {
            return Ok(());
        }
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_picture_size",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn set_picture_sizes_for_relationship(
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
        let mut candidate = document.inner.clone_for_staging();
        let boundary = candidate
            .paragraph_mut(paragraph_index)
            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
            .split_run(run_index, character_offset)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if run_count(&candidate) != before {
            let prepared = document.prepare_word_revision(
                py,
                &candidate,
                "Document.split_run",
                crate::document::WordPublication::Conservative,
            )?;
            document.commit_word_revision(prepared)?;
            document.inner = candidate;
        }
        Ok(boundary)
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    fn to_pdf<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        py.detach(|| match &font_refs {
            None => self.inner.to_pdf_with_options(options),
            Some(fonts) => self.inner.to_pdf_with_fonts_and_options(fonts, options),
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    fn to_pdf_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<(Bound<'py, PyBytes>, Bound<'py, PyTuple>)> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (payload, diagnostics) = py
            .detach(|| self.inner.to_pdf_with_diagnostics(font_refs, options))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((PyBytes::new(py, &payload), messages))
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

    #[pyo3(signature = (page_index, dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    #[allow(clippy::too_many_arguments)]
    fn render_page_to_png<'py>(
        &self,
        py: Python<'py>,
        page_index: usize,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<Option<Bound<'py, PyBytes>>> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        py.detach(|| match &font_refs {
            None => self
                .inner
                .render_page_to_png_with_options(page_index, dpi, options),
            Some(fonts) => self
                .inner
                .render_page_to_png_with_fonts_and_options(page_index, dpi, fonts, options),
        })
        .map(|bytes| bytes.map(|bytes| PyBytes::new(py, &bytes)))
        .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[pyo3(signature = (page_index, dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    #[allow(clippy::too_many_arguments)]
    fn render_page_to_png_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        page_index: usize,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<(Option<Bound<'py, PyBytes>>, Bound<'py, PyTuple>)> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (payload, diagnostics) = py
            .detach(|| {
                self.inner
                    .render_page_to_png_with_diagnostics(page_index, dpi, font_refs, options)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((payload.map(|bytes| PyBytes::new(py, &bytes)), messages))
    }
    #[pyo3(signature = (page_index, *, fonts = None, font_dir = None, font_dirs = None))]
    fn render_page_to_svg<'py>(
        &self,
        py: Python<'py>,
        page_index: usize,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Option<PySvgRenderResult>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let rendered = py
            .detach(|| match &font_refs {
                None => self.inner.render_page_to_svg(page_index),
                Some(fonts) => self.inner.render_page_to_svg_with_fonts_and_options(
                    page_index,
                    fonts,
                    rdocx::RenderOptions::default(),
                ),
            })
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

    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    #[allow(clippy::too_many_arguments)] // Explicit public rendering keyword operands.
    fn render_all_pages<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<Bound<'py, PyList>> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let pages = py
            .detach(|| match &font_refs {
                None => self.inner.render_all_pages_with_options(dpi, options),
                Some(fonts) => self
                    .inner
                    .render_all_pages_with_fonts_and_options(dpi, fonts, options),
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        PyList::new(py, pages.iter().map(|page| PyBytes::new(py, page)))
    }

    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    #[allow(clippy::too_many_arguments)] // Explicit public rendering keyword operands.
    fn render_all_pages_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<(Bound<'py, PyList>, Bound<'py, PyTuple>)> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (payload, diagnostics) = py
            .detach(|| {
                self.inner
                    .render_all_pages_with_diagnostics(dpi, font_refs, options)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((
            PyList::new(py, payload.iter().map(|bytes| PyBytes::new(py, bytes)))?,
            messages,
        ))
    }
    #[pyo3(signature = (*, dpi = 150.0, format = "png", quality = 90, transparent = false, pages = None, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    #[allow(clippy::too_many_arguments)]
    fn render_pages<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        format: &str,
        quality: u8,
        transparent: bool,
        pages: Option<Vec<usize>>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<Py<PyAny>> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let rendered = py
            .detach(|| {
                let format = parse_raster_format(format, quality, transparent)?;
                if let Some(fonts) = &font_refs {
                    let layout = self
                        .inner
                        .layout_with_additional_fonts_and_options(fonts, options)?;
                    let selected = match &pages {
                        Some(pages) => pages.clone(),
                        None => (0..layout.layout.pages.len()).collect(),
                    };
                    return oxml_pdf::render_pages(
                        &layout.layout,
                        &selected,
                        rdocx::RasterOptions { dpi, format },
                    )
                    .map_err(rdocx::Error::Raster);
                }
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

    #[pyo3(signature = (*, dpi = 150.0, format = "png", quality = 90, transparent = false, pages = None, fonts = None, font_dir = None, font_dirs = None, revision_view = "accepted", revision_style = "plain"))]
    #[allow(clippy::too_many_arguments)]
    fn render_pages_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        format: &str,
        quality: u8,
        transparent: bool,
        pages: Option<Vec<usize>>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
        revision_view: &str,
        revision_style: &str,
    ) -> PyResult<(Py<PyAny>, Bound<'py, PyTuple>)> {
        let options = parse_render_options(revision_view, revision_style)?;
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (rendered, diagnostics) = py
            .detach(|| {
                let format = parse_raster_format(format, quality, transparent)?;
                self.inner.render_pages_with_diagnostics(
                    pages.as_deref(),
                    rdocx::RasterOptions { dpi, format },
                    font_refs,
                    options,
                )
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let payload = match rendered {
            rdocx::RasterOutput::SeparatePages(pages) => {
                PyList::new(py, pages.iter().map(|bytes| PyBytes::new(py, bytes)))?
                    .into_any()
                    .unbind()
            }
            rdocx::RasterOutput::MultiPageTiff(bytes) => {
                PyBytes::new(py, &bytes).into_any().unbind()
            }
        };
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((payload, messages))
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
        let mut candidate = self.inner.clone_for_staging();
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
                let before = candidate.to_bytes()?;
                let diagnostics =
                    candidate.compare_with_options(&edited.inner, author, timestamp, &options)?;
                let changed = candidate.to_bytes()? != before;
                Ok::<_, rdocx::Error>((diagnostics, changed))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let result = PyTuple::new(
            py,
            diagnostics.into_iter().map(|item| PyComparisonDiagnostic {
                location: item.location,
                message: item.message,
            }),
        )?;
        if changed {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.compare",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(result)
    }

    #[getter]
    fn comments<'py>(slf: &Bound<'py, Self>) -> PyResult<Bound<'py, PyTuple>> {
        let py = slf.py();
        let owner = std::sync::Arc::new(slf.clone().unbind());
        let this = slf.borrow();
        let mut anchors = this
            .inner
            .comment_anchor_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut items = this
            .inner
            .standalone_story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .map(|item| (item.location().clone(), item))
            .collect::<HashMap<_, _>>();
        let records = this
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
                            start: this.comment_story_position(
                                py,
                                &range.start,
                                &mut items,
                                &owner,
                            )?,
                            end: this.comment_story_position(py, &range.end, &mut items, &owner)?,
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
        page_width: Option<Bound<'_, PyAny>>,
        page_height: Option<Bound<'_, PyAny>>,
        margin_top: Option<Bound<'_, PyAny>>,
        margin_right: Option<Bound<'_, PyAny>>,
        margin_bottom: Option<Bound<'_, PyAny>>,
        margin_left: Option<Bound<'_, PyAny>>,
        gutter: Option<Bound<'_, PyAny>>,
        column_count: Option<u32>,
        column_spacing: Option<Bound<'_, PyAny>>,
        page_number_start: Option<u32>,
        header_distance: Option<Bound<'_, PyAny>>,
        footer_distance: Option<Bound<'_, PyAny>>,
        different_first_page: Option<bool>,
        break_type: Option<&str>,
    ) -> PyResult<PySection> {
        let page_width = checked_update_section_length(py, page_width.as_ref(), "page_width")?;
        let page_height = checked_update_section_length(py, page_height.as_ref(), "page_height")?;
        let margin_top = checked_update_section_length(py, margin_top.as_ref(), "margin_top")?;
        let margin_right =
            checked_update_section_length(py, margin_right.as_ref(), "margin_right")?;
        let margin_bottom =
            checked_update_section_length(py, margin_bottom.as_ref(), "margin_bottom")?;
        let margin_left = checked_update_section_length(py, margin_left.as_ref(), "margin_left")?;
        let gutter = checked_update_section_length(py, gutter.as_ref(), "gutter")?;
        let column_spacing =
            checked_update_section_length(py, column_spacing.as_ref(), "column_spacing")?;
        let header_distance =
            checked_update_section_length(py, header_distance.as_ref(), "header_distance")?;
        let footer_distance =
            checked_update_section_length(py, footer_distance.as_ref(), "footer_distance")?;
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

    #[pyo3(signature = (reference, text = ""))]
    fn add_footnote(
        slf: &Bound<'_, Self>,
        py: Python<'_>,
        reference: &Bound<'_, PyAny>,
        text: &str,
    ) -> PyResult<Py<PyNote>> {
        Self::create_note(slf, py, reference, text, false)
    }
    #[pyo3(signature = (reference, text = ""))]
    fn add_endnote(
        slf: &Bound<'_, Self>,
        py: Python<'_>,
        reference: &Bound<'_, PyAny>,
        text: &str,
    ) -> PyResult<Py<PyNote>> {
        Self::create_note(slf, py, reference, text, true)
    }

    #[pyo3(signature = (family, *, format="decimal", start=None, restart="continuous", placement=None))]
    fn set_note_policy(
        &mut self,
        py: Python<'_>,
        family: &str,
        format: &str,
        start: Option<&Bound<'_, PyAny>>,
        restart: &str,
        placement: Option<&str>,
    ) -> PyResult<()> {
        let (family, policy) = checked_note_policy(family, format, start, restart, placement)?;
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_note_policy(family, policy)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_note_policy",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn set_page_color(&mut self, py: Python<'_>, color: &str) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_page_color(color)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_page_color",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    /// Author the existing native VML compatibility watermark.
    fn set_text_watermark(&mut self, py: Python<'_>, text: &str) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_text_watermark(text)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_text_watermark",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    /// Author the existing native VML image watermark, with checked EMU dimensions.
    fn set_image_watermark(
        &mut self,
        py: Python<'_>,
        data: &[u8],
        filename: &str,
        width: &Bound<'_, PyAny>,
        height: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let width = checked_section_length(Some(width), "watermark width")?
            .ok_or_else(|| PyTypeError::new_err("watermark width is required"))?;
        let height = checked_section_length(Some(height), "watermark height")?
            .ok_or_else(|| PyTypeError::new_err("watermark height is required"))?;
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_image_watermark(
                data,
                filename,
                rdocx::Length::emu(width),
                rdocx::Length::emu(height),
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_image_watermark",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn section(slf: &Bound<'_, Self>, index: usize) -> PyResult<PySectionHandle> {
        let document = slf.borrow();
        if index >= document.inner.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        let (path, owner) = document.capture_owner(
            &WordOwnerLocation::Section(index),
            smallvec![],
            WordView::Owner,
            None,
        )?;
        Ok(PySectionHandle {
            document: slf.clone().unbind(),
            index,
            path,
            owner,
        })
    }
    #[getter]
    fn settings(slf: &Bound<'_, Self>) -> PySettings {
        PySettings {
            document: slf.clone().unbind(),
        }
    }

    fn insert_section(&mut self, py: Python<'_>, index: usize) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        if index > candidate.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        candidate
            .insert_section(index)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.insert_section",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn remove_section(&mut self, py: Python<'_>, index: usize) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        if index >= candidate.section_count() {
            return Err(PyIndexError::new_err("section index out of range"));
        }
        candidate
            .remove_section(index)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.remove_section",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    #[getter]
    fn styles<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let snapshots = PyTuple::new(py, self.inner.styles().into_iter().map(style_snapshot))?;
        py.import("rdocx")?
            .getattr("_StyleSnapshots")?
            .call1((snapshots,))?
            .cast_into::<PyTuple>()
            .map_err(Into::into)
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
        alignment = None,
        line_spacing = None,
        underline = None,
        keep_with_next = None,
        keep_together = None,
        page_break_before = None,
        borders = None,
        shading = None,
        tab_stops = None,
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
        color: Option<&Bound<'_, PyAny>>,
        space_before: Option<i64>,
        space_after: Option<i64>,
        left_indent: Option<i64>,
        right_indent: Option<i64>,
        first_line_indent: Option<i64>,
        alignment: Option<i32>,
        line_spacing: Option<&Bound<'_, PyAny>>,
        underline: Option<&Bound<'_, PyAny>>,
        keep_with_next: Option<bool>,
        keep_together: Option<bool>,
        page_break_before: Option<bool>,
        borders: Option<&Bound<'_, PyAny>>,
        shading: Option<&Bound<'_, PyAny>>,
        tab_stops: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyStyle> {
        let color = color
            .filter(|value| !value.is_none())
            .map(|value| oxml_py_support::rgb_channels(value, "Document.add_style(color)"))
            .transpose()?;
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
            alignment,
            line_spacing: line_spacing
                .filter(|v| !v.is_none())
                .map(|v| v.clone().unbind()),
            underline: underline
                .filter(|v| !v.is_none())
                .map(|v| v.clone().unbind()),
            keep_with_next,
            keep_together,
            page_break_before,
            borders: borders.filter(|v| !v.is_none()).map(|v| v.clone().unbind()),
            shading: shading.filter(|v| !v.is_none()).map(|v| v.clone().unbind()),
            tab_stops: tab_stops
                .filter(|v| !v.is_none())
                .map(|v| v.clone().unbind()),
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
        if let Some(properties) = formatting.paragraph_properties(py, None)? {
            if native_type == rdocx::StyleType::Character {
                return Err(PyValueError::new_err(
                    "a character style takes no paragraph formatting",
                ));
            }
            builder = builder.replace_paragraph_properties(properties);
        }
        if let Some(properties) = formatting.run_properties(py, None)? {
            builder = builder.run_properties(properties);
            if let Some(name) = &formatting.font_name {
                builder = builder.font(name);
            }
            if let Some((r, g, b)) = formatting.color {
                builder = builder.color(&format!("{r:02X}{g:02X}{b:02X}"));
            }
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
        alignment = None,
        line_spacing = None,
        underline = None,
        keep_with_next = None,
        keep_together = None,
        page_break_before = None,
        borders = None,
        shading = None,
        tab_stops = None,
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
        color: Option<&Bound<'_, PyAny>>,
        space_before: Option<i64>,
        space_after: Option<i64>,
        left_indent: Option<i64>,
        right_indent: Option<i64>,
        first_line_indent: Option<i64>,
        alignment: Option<i32>,
        line_spacing: Option<&Bound<'_, PyAny>>,
        underline: Option<&Bound<'_, PyAny>>,
        keep_with_next: Option<bool>,
        keep_together: Option<bool>,
        page_break_before: Option<bool>,
        borders: Option<&Bound<'_, PyAny>>,
        shading: Option<&Bound<'_, PyAny>>,
        tab_stops: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyStyle> {
        let color = color
            .filter(|value| !value.is_none())
            .map(|value| oxml_py_support::rgb_channels(value, "Document.set_style(color)"))
            .transpose()?;
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
            alignment,
            line_spacing: line_spacing
                .filter(|v| !v.is_none())
                .map(|v| v.clone().unbind()),
            underline: underline
                .filter(|v| !v.is_none())
                .map(|v| v.clone().unbind()),
            keep_with_next,
            keep_together,
            page_break_before,
            borders: borders.filter(|v| !v.is_none()).map(|v| v.clone().unbind()),
            shading: shading.filter(|v| !v.is_none()).map(|v| v.clone().unbind()),
            tab_stops: tab_stops
                .filter(|v| !v.is_none())
                .map(|v| v.clone().unbind()),
        };
        if let Some(properties) =
            formatting.paragraph_properties(py, existing.paragraph_properties())?
        {
            if style_type == rdocx::StyleType::Character {
                return Err(PyValueError::new_err(
                    "a character style takes no paragraph formatting",
                ));
            }
            builder = builder.replace_paragraph_properties(properties);
        }
        if let Some(properties) = formatting.run_properties(py, existing.run_properties())? {
            builder = builder.run_properties(properties);
            if let Some(name) = &formatting.font_name {
                builder = builder.font(name);
            }
            if let Some((r, g, b)) = formatting.color {
                builder = builder.color(&format!("{r:02X}{g:02X}{b:02X}"));
            }
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

    #[getter]
    fn default_paragraph_format(&self, py: Python<'_>) -> PyResult<PyDefaultParagraphFormat> {
        default_paragraph_format_snapshot(py, self.inner.default_paragraph_properties())
    }

    #[pyo3(signature = (**kwargs))]
    fn set_default_paragraph_format(
        &mut self,
        py: Python<'_>,
        kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<()> {
        let Some(kwargs) = kwargs else {
            return Ok(());
        };
        for (key, _) in kwargs.iter() {
            let key = key.extract::<String>()?;
            if !matches!(
                key.as_str(),
                "space_before" | "space_after" | "line_spacing"
            ) {
                return Err(PyTypeError::new_err(format!(
                    "set_default_paragraph_format() got an unexpected keyword argument '{key}'"
                )));
            }
        }
        let mut next = self
            .inner
            .default_paragraph_properties()
            .cloned()
            .unwrap_or_default();
        for name in ["space_before", "space_after", "line_spacing"] {
            let Some(value) = kwargs.get_item(name)? else {
                continue;
            };
            let twips = length_keyword(py, &value, name)?;
            match name {
                "space_before" => next.space_before = twips,
                "space_after" => next.space_after = twips,
                "line_spacing" => {
                    next.line_spacing = twips;
                    next.line_rule = twips.map(|_| "exact".to_owned());
                }
                _ => unreachable!("keyword was checked above"),
            }
        }
        let current = self.inner.default_paragraph_properties().cloned();
        if current.as_ref() == Some(&next)
            || (current.is_none() && next == rdocx::CT_PPr::default())
        {
            return Ok(());
        }
        self.inner
            .set_default_paragraph_properties(Some(next))
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn part_names<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.inner
                .part_names()
                .map_err(|error| rdocx_to_pyerr(py, error))?,
        )
    }

    fn part_xml(&self, py: Python<'_>, name: &str) -> PyResult<String> {
        self.inner
            .part_xml(name)
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

    #[pyo3(signature = (definition_id, *, start = None, level = 0))]
    fn add_numbering_instance(
        &mut self,
        py: Python<'_>,
        definition_id: u32,
        start: Option<i64>,
        level: i64,
    ) -> PyResult<u32> {
        if !(0..=8).contains(&level) {
            return Err(PyValueError::new_err("level must be between 0 and 8"));
        }
        if start.is_none() && level != 0 {
            return Err(PyValueError::new_err("nonzero level requires start"));
        }
        let overrides = start
            .map(|start| {
                let start = u32::try_from(start)
                    .map_err(|_| PyValueError::new_err("start must be between 0 and 4294967295"))?;
                Ok::<_, PyErr>(rdocx::NumberingLevelOverride::new(level as u32).start(start))
            })
            .transpose()?
            .into_iter()
            .collect::<Vec<_>>();
        self.inner
            .add_numbering_instance(definition_id, &overrides)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    #[getter]
    fn charts<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let values = self
            .inner
            .charts()
            .map_err(|e| rdocx_to_pyerr(py, e))?
            .into_iter()
            .map(|(chart_type, data)| Py::new(py, PyChart { chart_type, data }))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, values)
    }

    #[pyo3(signature = (paragraph, *, start=Python::attach(|py| PyInt::new(py, 1).into_any().unbind())), text_signature = "($self, paragraph, *, start=1)")]
    fn restart_numbering(
        slf: Py<Self>,
        py: Python<'_>,
        paragraph: PyRef<'_, PyParagraph>,
        start: Py<PyAny>,
    ) -> PyResult<u32> {
        let start = start.bind(py);
        if !start.is_instance_of::<PyInt>() || start.is_instance_of::<PyBool>() {
            return Err(PyTypeError::new_err("start must be a non-bool integer"));
        }
        let start = start.extract::<u32>()?;
        if !paragraph.belongs_to(py, &slf) {
            return Err(PyValueError::new_err(
                "paragraph belongs to a different document",
            ));
        }
        let ParagraphLocation::Body(index) = paragraph.validate(py)? else {
            return Err(PyValueError::new_err(
                "numbering restart requires a direct body paragraph",
            ));
        };
        let mut document = slf.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let id = candidate
            .restart_numbering(index, start)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        document.publish_api_candidate(py, candidate, "Document.restart_numbering")?;
        Ok(id)
    }

    fn remove_footnote(&mut self, py: Python<'_>, id: &Bound<'_, PyAny>) -> PyResult<()> {
        self.remove_note_by_id(py, id, false)
    }

    fn remove_endnote(&mut self, py: Python<'_>, id: &Bound<'_, PyAny>) -> PyResult<()> {
        self.remove_note_by_id(py, id, true)
    }

    /// Effective legend properties of the live direct-body chart paragraph.
    fn chart_legend(
        slf: Py<Self>,
        py: Python<'_>,
        paragraph: PyRef<'_, PyParagraph>,
    ) -> PyResult<Py<PyDict>> {
        if !paragraph.belongs_to(py, &slf) {
            return Err(PyValueError::new_err(
                "paragraph belongs to a different document",
            ));
        }
        let ParagraphLocation::Body(index) = paragraph.validate(py)? else {
            return Err(PyValueError::new_err(
                "chart legend requires a direct body paragraph",
            ));
        };
        let settings = slf
            .borrow(py)
            .inner
            .chart_legend(index)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let result = PyDict::new(py);
        result.set_item("visible", settings.visible)?;
        result.set_item("position", settings.position.as_str())?;
        result.set_item("overlay", settings.overlay)?;
        Ok(result.unbind())
    }

    /// None leaves a property unchanged. Hidden legends discard their settings.
    #[pyo3(signature=(paragraph, *, visible=None, position=None, overlay=None))]
    fn set_chart_legend(
        slf: Py<Self>,
        py: Python<'_>,
        paragraph: PyRef<'_, PyParagraph>,
        visible: Option<&Bound<'_, PyAny>>,
        position: Option<&str>,
        overlay: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let checked_bool = |value: Option<&Bound<'_, PyAny>>| -> PyResult<Option<bool>> {
            value
                .map(|v| {
                    if !v.is_instance_of::<PyBool>() {
                        return Err(PyTypeError::new_err(
                            "legend visibility and overlay require bool",
                        ));
                    }
                    v.extract()
                })
                .transpose()
        };
        let visible = checked_bool(visible)?;
        let overlay = checked_bool(overlay)?;
        let position = position
            .map(rdocx::LegendPosition::parse)
            .transpose()
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        if !paragraph.belongs_to(py, &slf) {
            return Err(PyValueError::new_err(
                "paragraph belongs to a different document",
            ));
        }
        let ParagraphLocation::Body(index) = paragraph.validate(py)? else {
            return Err(PyValueError::new_err(
                "chart legend requires a direct body paragraph",
            ));
        };
        let location = ParagraphLocation::Body(index)
            .native(&slf.borrow(py).inner)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let mut document = slf.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .set_chart_legend(index, visible, position, overlay)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        document.publish_owner_candidate(
            py,
            candidate,
            "Document.set_chart_legend",
            WordPublication::ElementXml(WordOwnerLocation::Paragraph(location)),
        )
    }

    #[pyo3(signature = (chart_type, chart_data, width, height, *, after=None))]
    fn add_chart(
        slf: Py<Self>,
        py: Python<'_>,
        chart_type: &Bound<'_, PyAny>,
        chart_data: &Bound<'_, PyAny>,
        width: &Bound<'_, PyAny>,
        height: &Bound<'_, PyAny>,
        after: Option<PyRef<'_, PyParagraph>>,
    ) -> PyResult<Py<PyParagraph>> {
        let (kind, data) = checked_chart_inputs(chart_type, chart_data)?;
        let width =
            rdocx::Length::emu(checked_section_length(Some(width), "chart width")?.unwrap());
        let height =
            rdocx::Length::emu(checked_section_length(Some(height), "chart height")?.unwrap());
        let after = after
            .map(|paragraph| {
                if !paragraph.belongs_to(py, &slf) {
                    return Err(PyValueError::new_err(
                        "paragraph belongs to a different document",
                    ));
                }
                let ParagraphLocation::Body(index) = paragraph.validate(py)? else {
                    return Err(PyValueError::new_err(
                        "chart insertion requires a direct body paragraph",
                    ));
                };
                Ok(index)
            })
            .transpose()?;
        let mut document = slf.borrow_mut(py);
        let index = after.map_or(document.inner.paragraph_count(), |index| index + 1);
        let mut candidate = document.inner.clone_for_staging();
        match after {
            Some(after) => candidate.add_chart_after(after, kind, width, height, &data),
            None => candidate.add_chart(kind, width, height, &data),
        }
        .map(drop)
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_chart",
            if after.is_none() {
                WordPublication::AppendBody
            } else {
                WordPublication::Conservative
            },
        )?;
        let path = document
            .revisions
            .capture_prepared(
                &prepared.revision,
                smallvec![PathSeg::Body(0), PathSeg::Para(index)],
            )
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyParagraph::new(slf.clone_ref(py), path)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    #[pyo3(signature = (text, level = 0))]
    fn add_bullet_list_item(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        level: i64,
    ) -> PyResult<Py<PyParagraph>> {
        if !(0..=8).contains(&level) {
            return Err(PyValueError::new_err("level must be between 0 and 8"));
        }
        let mut document = slf.borrow_mut(py);
        let index = document.inner.paragraph_count();
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .try_add_bullet_list_item(text, level as u32)
            .map(|_| ())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_bullet_list_item",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(
                &prepared.revision,
                smallvec![PathSeg::Body(0), PathSeg::Para(index)],
            )
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyParagraph::new(slf.clone_ref(py), path)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    #[pyo3(signature = (text, level = 0, *, restart = false))]
    fn add_numbered_list_item(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        level: i64,
        restart: bool,
    ) -> PyResult<Py<PyParagraph>> {
        if !(0..=8).contains(&level) {
            return Err(PyValueError::new_err("level must be between 0 and 8"));
        }
        let mut document = slf.borrow_mut(py);
        let index = document.inner.paragraph_count();
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .try_add_numbered_list_item_with_restart(text, level as u32, restart)
            .map(|_| ())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_numbered_list_item",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(
                &prepared.revision,
                smallvec![PathSeg::Body(0), PathSeg::Para(index)],
            )
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyParagraph::new(slf.clone_ref(py), path)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
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
        PyTuple::new(
            py,
            stories
                .iter()
                .map(|story| self.owned_story_snapshot(story))
                .collect::<PyResult<Vec<_>>>()?,
        )
    }

    #[getter]
    fn story_items<'py>(slf: &Bound<'py, Self>) -> PyResult<Bound<'py, PyTuple>> {
        let py = slf.py();
        let owner = std::sync::Arc::new(slf.clone().unbind());
        let this = slf.borrow();
        let items = this
            .inner
            .standalone_story_item_snapshots()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let snapshots = items
            .into_iter()
            .map(|item| {
                Ok(PyStoryItem {
                    story: this.owned_story_snapshot(item.location().story())?,
                    kind: story_item_kind_name(item.location().item_kind()).to_owned(),
                    index_path: item.location().index_path().to_vec(),
                    direct_body_index: item.direct_body_index(),
                    text: item.text().map(str::to_owned),
                    xml: item.xml().to_vec(),
                    revision: this.revisions.current(),
                    binding: Some(this.capture_snapshot(smallvec![])?),
                    document: Some(owner.clone()),
                })
            })
            .collect::<PyResult<Vec<_>>>()?;
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
                        story: resolved
                            .as_ref()
                            .map(|value| self.owned_story_snapshot(value.story()))
                            .transpose()?,
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
    ) -> PyResult<Py<PyStory>> {
        let (kind, variant) = self.section_story_target(section_index, kind, variant)?;
        let mut candidate = self.inner.clone_for_staging();
        let story = py
            .detach(|| candidate.create_section_story(section_index, kind, variant))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.create_section_story",
            crate::document::WordPublication::Conservative,
        )?;
        let mut snapshot = story_snapshot(&story);
        snapshot.binding = Some(self.capture_prepared_snapshot(&prepared, smallvec![])?);
        let result = Py::new(py, snapshot)?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(result)
    }

    fn link_section_story(
        &mut self,
        py: Python<'_>,
        section_index: usize,
        kind: &str,
        variant: &str,
        story: PyRef<'_, PyStory>,
    ) -> PyResult<Py<PyStory>> {
        let (kind, variant) = self.section_story_target(section_index, kind, variant)?;
        let story = self.native_story(py, &story)?;
        let mut candidate = self.inner.clone_for_staging();
        let linked = py
            .detach(|| candidate.link_section_story(section_index, kind, variant, &story))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.link_section_story",
            crate::document::WordPublication::Conservative,
        )?;
        let mut snapshot = story_snapshot(&linked);
        snapshot.binding = Some(self.capture_prepared_snapshot(&prepared, smallvec![])?);
        let result = Py::new(py, snapshot)?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(result)
    }

    fn unlink_section_story(
        &mut self,
        py: Python<'_>,
        section_index: usize,
        kind: &str,
        variant: &str,
    ) -> PyResult<Py<PyStory>> {
        let (kind, variant) = self.section_story_target(section_index, kind, variant)?;
        let mut candidate = self.inner.clone_for_staging();
        let story = py
            .detach(|| candidate.unlink_section_story(section_index, kind, variant))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.unlink_section_story",
            crate::document::WordPublication::Conservative,
        )?;
        let mut snapshot = story_snapshot(&story);
        snapshot.binding = Some(self.capture_prepared_snapshot(&prepared, smallvec![])?);
        let result = Py::new(py, snapshot)?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(result)
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
                let owner = self
                    .owners
                    .iter()
                    .find(|owner| {
                        matches!(&owner.location, WordOwnerLocation::Hyperlink(current, _)
                        if same_content_owner(current, &location))
                            && owner.projection_ordinal == Some(*next)
                    })
                    .ok_or_else(|| {
                        PyValueError::new_err("hyperlink has no checked physical source owner")
                    })?;
                let binding = self
                    .revisions
                    .capture_scopes(smallvec![], std::slice::from_ref(&owner.owner))
                    .map_err(revision_to_pyerr)?;
                let owner = Some((binding, owner.owner.clone()));
                *next += 1;
                Ok(PyHyperlink {
                    story: self.owned_story_snapshot(location.story())?,
                    index_path: location.index_path().to_vec(),
                    text: link.text,
                    url: link.url,
                    anchor: link.anchor,
                    relationship_id: link.rel_id,
                    position,
                    owner,
                })
            })
            .collect::<PyResult<Vec<_>>>()?;
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
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.remove_hyperlink(&story, link_index))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let selected = self
            .owners
            .iter()
            .find_map(|owner| match &owner.location {
                WordOwnerLocation::Hyperlink(location, _)
                    if same_story_owner(location.story(), &story)
                        && owner.projection_ordinal == Some(link_index) =>
                {
                    Some(owner.location.clone())
                }
                _ => None,
            })
            .ok_or_else(|| {
                PyValueError::new_err("hyperlink has no checked physical source owner")
            })?;
        self.publish_owner_candidate(
            py,
            candidate,
            "Document.remove_hyperlink",
            WordPublication::HyperlinkRemove(selected),
        )
    }

    fn set_header(&mut self, py: Python<'_>, text: &str) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.try_set_header(text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_header",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn set_footer(&mut self, py: Python<'_>, text: &str) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.try_set_footer(text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_footer",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn set_story_text(
        &mut self,
        py: Python<'_>,
        item: PyRef<'_, PyStoryItem>,
        text: &str,
    ) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        let location = self.native_location(py, &item)?;
        py.detach(|| candidate.set_story_text(&location, text))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.set_story_text",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    fn add_hyperlink_to_story(
        &mut self,
        py: Python<'_>,
        story: PyRef<'_, PyStory>,
        text: &str,
        url: &str,
    ) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        let story = self.native_story(py, &story)?;
        py.detach(|| candidate.add_hyperlink_to_story(&story, text, url))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.add_hyperlink_to_story",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    #[allow(clippy::too_many_arguments)] // Preserve explicit checked public picture operands.
    #[pyo3(signature = (data, filename=None, width=None, height=None, *, after=None, options=None))]
    fn add_picture(
        &mut self,
        py: Python<'_>,
        data: &Bound<'_, PyAny>,
        filename: Option<&str>,
        width: Option<&Bound<'_, PyAny>>,
        height: Option<&Bound<'_, PyAny>>,
        after: Option<&Bound<'_, PyAny>>,
        options: Option<PyRef<'_, PyPictureOptions>>,
    ) -> PyResult<Py<PyPicture>> {
        let (data, filename) = picture_input(data, filename)?;
        let width = checked_section_length(width, "picture width")?.map(rdocx::Length::emu);
        let height = checked_section_length(height, "picture height")?.map(rdocx::Length::emu);
        if width.is_some() != height.is_some() {
            return Err(crate::rdocx_to_pyerr(
                py,
                rdocx::Error::Other(
                    "use add_picture(..., width=..., height=...) with both dimensions; omit both for the native image size".to_owned(),
                ),
            ));
        }
        let after = after
            .map(|value| -> PyResult<rdocx::ContentLocation> {
                if let Ok(item) = value.extract::<PyRef<'_, PyStoryItem>>() {
                    return self.native_location(py, &item);
                }
                if let Ok(picture) = value.extract::<PyRef<'_, PyPicture>>() {
                    picture
                        .binding
                        .validate_scopes(
                            &self.revisions,
                            "picture",
                            "Re-fetch it with document.pictures().",
                        )
                        .map_err(|error| crate::stale_to_pyerr(py, error))?;
                    let current = self
                        .inner
                        .pictures()
                        .map_err(|error| rdocx_to_pyerr(py, error))?;
                    if !current.contains(&picture.inner) {
                        return Err(PyValueError::new_err("picture placement source changed"));
                    }
                    return Ok(picture.inner.location.clone());
                }
                Err(PyTypeError::new_err("after must be a StoryItem or Picture"))
            })
            .transpose()?;
        let story = match &after {
            Some(location) => location.story().clone(),
            None => self.body_story(py)?,
        };
        let mut candidate = self.inner.clone_for_staging();
        let location = candidate
            .insert_picture_to_story_with_options(
                &story,
                after.as_ref(),
                &data,
                &filename,
                width,
                height,
                options.map(|options| options.inner.clone()),
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let picture = candidate
            .pictures()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .find(|picture| picture.location == location)
            .ok_or_else(|| PyIndexError::new_err("inserted picture was not inventoried"))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.add_picture",
            crate::document::WordPublication::Conservative,
        )?;
        let result = Py::new(
            py,
            PyPicture {
                inner: picture,
                binding: self.capture_prepared_snapshot(&prepared, smallvec![])?,
            },
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(result)
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
        let mut candidate = self.inner.clone_for_staging();
        let id = if let Ok(range) = range.cast::<PyRunRange>() {
            candidate.add_comment_with_date((*range.borrow()).into(), author, initials, text, date)
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
            candidate.add_story_comment_with_date(
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
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.add_comment",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
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
        let mut candidate = self.inner.clone_for_staging();
        let id = candidate
            .add_comment_on_text(anchor, occurrence, author, initials, text, date)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.add_comment_on_text",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(id)
    }

    /// Move a root comment to a checked current story range atomically.
    fn move_comment(
        &mut self,
        id: i32,
        range: PyRef<'_, PyStoryRunRange>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        let start = rdocx::StoryRunPosition {
            location: self.native_location(py, &range.start.item)?,
            run_index: range.start.run_index,
        };
        let end = rdocx::StoryRunPosition {
            location: self.native_location(py, &range.end.item)?,
            run_index: range.end.run_index,
        };
        py.detach(|| candidate.move_comment(id, rdocx::StoryRunRange { start, end }))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.move_comment",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
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
        let mut candidate = self.inner.clone_for_staging();
        py.detach(|| candidate.move_comment_to_text(id, anchor, occurrence))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.move_comment_to_text",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
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
        let mut candidate = self.inner.clone_for_staging();
        let id = candidate
            .reply_to_with_date(parent_id, author, text, date)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.reply_to",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
        Ok(id)
    }

    #[pyo3(signature = (id, *, resolved = true))]
    fn resolve_comment(&mut self, id: i32, resolved: bool, py: Python<'_>) -> PyResult<bool> {
        let mut candidate = self.inner.clone_for_staging();
        let updated = candidate
            .resolve_comment(id, resolved)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if updated {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.resolve_comment",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(updated)
    }

    fn remove_comment(&mut self, id: i32, py: Python<'_>) -> PyResult<bool> {
        let mut candidate = self.inner.clone_for_staging();
        let removed = candidate
            .remove_comment(id)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if removed {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.remove_comment",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
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

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None))]
    fn layout<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let fragments = py
            .detach(|| {
                let layout = match &font_refs {
                    None => self.inner.layout_deterministic()?,
                    Some(fonts) => {
                        std::sync::Arc::new(self.inner.layout_with_additional_fonts_and_options(
                            fonts,
                            rdocx::RenderOptions::default(),
                        )?)
                    }
                };
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

    #[pyo3(signature = (page_index, *, fonts = None, font_dir = None, font_dirs = None))]
    fn layout_page<'py>(
        &self,
        page_index: usize,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Option<PyLayoutPage>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        py.detach(|| {
            let layout = match &font_refs {
                None => self.inner.layout_deterministic()?,
                Some(fonts) => {
                    std::sync::Arc::new(self.inner.layout_with_additional_fonts_and_options(
                        fonts,
                        rdocx::RenderOptions::default(),
                    )?)
                }
            };
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

    fn rebuild_toc(&mut self, py: Python<'_>) -> PyResult<Py<PyTocRebuildReport>> {
        let mut candidate = self.inner.clone_for_staging();
        let (report, changed) = py
            .detach(|| {
                let before = candidate.to_bytes()?;
                let report = candidate.rebuild_toc()?;
                let changed = candidate.to_bytes()? != before;
                Ok::<_, rdocx::Error>((report, changed))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let result = Py::new(
            py,
            PyTocRebuildReport {
                entry_count: report.entry_count,
                bookmark_count: report.bookmark_count,
                diagnostics: report.diagnostics,
            },
        )?;
        if changed {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.rebuild_toc",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(result)
    }

    #[getter(revisions)]
    fn revision_snapshots<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let revisions = self
            .inner
            .story_revisions()
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        PyTuple::new(
            py,
            revisions
                .iter()
                .map(|revision| {
                    Ok(PyRevision {
                        id: revision.id(),
                        author: revision.author().to_owned(),
                        timestamp: revision.timestamp().map(str::to_owned),
                        kind: revision_kind_name(revision.kind()).to_owned(),
                        story: Some(self.owned_story_snapshot(revision.story())?),
                    })
                })
                .collect::<PyResult<Vec<_>>>()?,
        )
    }

    fn accept_all(&mut self, py: Python<'_>) -> PyResult<usize> {
        self.counted_mutation(py, "Document.accept_all", rdocx::Document::accept_all)
    }

    fn reject_all(&mut self, py: Python<'_>) -> PyResult<usize> {
        self.counted_mutation(py, "Document.reject_all", rdocx::Document::reject_all)
    }

    fn accept_revisions_by_author(&mut self, py: Python<'_>, author: &str) -> PyResult<usize> {
        self.counted_mutation(py, "Document.accept_revisions_by_author", |document| {
            document.accept_revisions_by_author(author)
        })
    }

    fn reject_revisions_by_author(&mut self, py: Python<'_>, author: &str) -> PyResult<usize> {
        self.counted_mutation(py, "Document.reject_revisions_by_author", |document| {
            document.reject_revisions_by_author(author)
        })
    }

    #[pyo3(signature = (*, start, end))]
    fn accept_revisions_in_date_range(
        &mut self,
        py: Python<'_>,
        start: &str,
        end: &str,
    ) -> PyResult<usize> {
        self.counted_mutation(py, "Document.accept_revisions_in_date_range", |document| {
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
        self.counted_mutation(py, "Document.reject_revisions_in_date_range", |document| {
            document.reject_revisions_in_date_range(start, end)
        })
    }

    fn accept_revision_id(&mut self, py: Python<'_>, id: i32) -> PyResult<usize> {
        self.counted_mutation(py, "Document.accept_revision_id", |document| {
            document.accept_revision_id(id)
        })
    }

    fn reject_revision_id(&mut self, py: Python<'_>, id: i32) -> PyResult<usize> {
        self.counted_mutation(py, "Document.reject_revision_id", |document| {
            document.reject_revision_id(id)
        })
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
        self.scoped_replacement(py, "Document.replace_text_at", |document| {
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
            return self.counted_mutation(py, "Document.try_replace_text", |document| {
                document.try_replace_text(placeholder, replacement)
            });
        }
        let (candidate, counts) =
            self.expected_replacements(py, &[(placeholder, replacement, expect)], false)?;
        if counts[0] > 0 {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.try_replace_text",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
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
        let (candidate, counts) = self.expected_replacements(py, &pairs, true)?;
        let changed = counts.iter().any(|count| *count > 0);
        let result = PyTuple::new(py, counts)?;
        if changed {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.replace_all",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(result)
    }

    fn replace_all_regex_expected<'py>(
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
        let (candidate, counts) = {
            let mut candidate = self.inner.clone_for_staging();
            let counts = py
                .detach(|| candidate.try_replace_all_regex_expected(&pairs))
                .map_err(|error| rdocx_to_pyerr(py, error))?
                .map_err(|mismatch| crate::replacement_count_to_pyerr(py, &mismatch, true))?;
            (candidate, counts)
        };

        let result = PyTuple::new(py, counts)?;
        self.publish_api_candidate(py, candidate, "Document.replace_all_regex_expected")?;

        Ok(result)
    }

    fn replace_all_regex(
        &mut self,
        py: Python<'_>,
        patterns: Vec<(String, String)>,
    ) -> PyResult<usize> {
        self.counted_mutation(py, "Document.replace_all_regex", |document| {
            document.replace_all_regex(&patterns)
        })
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
        self.counted_mutation(py, "Document.update_fields", |document| {
            document.update_fields(&context)
        })
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
    fn update_all_fields(
        &mut self,
        py: Python<'_>,
        now: Option<&Bound<'_, PyAny>>,
        file_name: Option<String>,
        file_path: Option<String>,
        merge_fields: Option<BTreeMap<String, String>>,
        included_text: Option<BTreeMap<String, String>>,
        merge_record_number: Option<u32>,
        merge_sequence_number: Option<u32>,
    ) -> PyResult<Py<PyFieldsUpdateReport>> {
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
        let mut candidate = self.inner.clone_for_staging();
        let report = py
            .detach(|| candidate.update_all_fields(&context))
            .map_err(|error| rdocx_to_pyerr(py, error))?;

        // Every returned object, including nested outcomes and tuples, exists
        // before publishing. Returning this owned Py object needs no allocation.
        let json_loads = py.import("json")?.getattr("loads")?;
        let evaluations = report
            .evaluations
            .iter()
            .map(|evaluation| {
                let outcome = json_loads
                    .call1((evaluation.to_json_value()["outcome"].to_string(),))?
                    .cast_into::<PyDict>()?
                    .unbind();
                Py::new(
                    py,
                    PyFieldEvaluation {
                        field_index: evaluation.field_index,
                        instruction: evaluation.instruction.clone(),
                        cached_result: evaluation.cached_result.clone(),
                        outcome,
                    },
                )
            })
            .collect::<PyResult<Vec<_>>>()?;
        let evaluations = PyTuple::new(py, evaluations)?.unbind();
        let diagnostics = PyTuple::new(py, &report.diagnostics)?.unbind();
        let layout = Py::new(
            py,
            PyLayoutBackedFieldUpdateReport {
                page_fields: report.layout.page_fields,
                num_pages_fields: report.layout.num_pages_fields,
                page_reference_fields: report.layout.page_reference_fields,
                section_fields: report.layout.section_fields,
                section_pages_fields: report.layout.section_pages_fields,
                diagnostics: report.layout.diagnostics,
            },
        )?;
        let result = Py::new(
            py,
            PyFieldsUpdateReport {
                ordinary_updated: report.ordinary_updated,
                layout,
                evaluations,
                diagnostics,
                total_updated: report.total_updated,
                saved_cache_scope: report.saved_cache_scope.to_owned(),
            },
        )?;
        self.publish_api_candidate(py, candidate, "Document.update_all_fields")?;

        Ok(result)
    }

    fn update_page_fields(&mut self, py: Python<'_>) -> PyResult<usize> {
        let mut candidate = self.inner.clone_for_staging();
        let updated = py
            .detach(|| candidate.update_page_fields())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if updated != 0 {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.update_page_fields",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(updated)
    }

    fn update_layout_backed_fields(
        &mut self,
        py: Python<'_>,
    ) -> PyResult<Py<PyLayoutBackedFieldUpdateReport>> {
        let mut candidate = self.inner.clone_for_staging();
        let report = py
            .detach(|| candidate.update_layout_backed_fields())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let changed = report.updated_count() != 0;
        let result = Py::new(
            py,
            PyLayoutBackedFieldUpdateReport {
                page_fields: report.page_fields,
                num_pages_fields: report.num_pages_fields,
                page_reference_fields: report.page_reference_fields,
                section_fields: report.section_fields,
                section_pages_fields: report.section_pages_fields,
                diagnostics: report.diagnostics,
            },
        )?;
        if changed {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.update_layout_backed_fields",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(result)
    }

    #[pyo3(signature = (index, max_level = 3))]
    fn insert_toc(&mut self, py: Python<'_>, index: usize, max_level: u32) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        if index > candidate.content_count() {
            return Err(PyIndexError::new_err("content index out of range"));
        }
        if !(1..=9).contains(&max_level) {
            return Err(PyValueError::new_err("max_level must be between 1 and 9"));
        }
        candidate
            .insert_toc(index, max_level)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.insert_toc",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
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

    #[pyo3(signature = (text = "", style = None))]
    fn add_paragraph(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        style: Option<&str>,
    ) -> PyResult<Py<PyParagraph>> {
        Self::append_paragraph_candidate(slf, py, text, style, false)
    }

    #[pyo3(signature = (text = "", level = 1))]
    fn add_heading(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        #[pyo3(from_py_with = checked_heading_level)] level: i64,
    ) -> PyResult<Py<PyParagraph>> {
        if !(0..=9).contains(&level) {
            return Err(PyValueError::new_err(
                "heading level must be from 0 through 9",
            ));
        }
        let style = if level == 0 {
            "Title".to_owned()
        } else {
            format!("Heading{level}")
        };
        Self::append_paragraph_candidate(slf, py, text, Some(&style), false)
    }

    fn add_page_break(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyParagraph>> {
        Self::append_paragraph_candidate(slf, py, "", None, true)
    }

    #[pyo3(signature = (start_type = None))]
    fn add_section(&self, start_type: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let _ = start_type;
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "use document.insert_section(len(document.sections)) and document.update_section(index, break_type=...); existing native insertion copies geometry only",
        ))
    }

    #[pyo3(signature = (rows, cols, style = None))]
    fn add_table(
        slf: Py<Self>,
        py: Python<'_>,
        rows: usize,
        cols: usize,
        style: Option<&str>,
    ) -> PyResult<Py<PyTable>> {
        let mut document = slf.borrow_mut(py);
        document.refuse_unsavable_body_edit(py)?;
        let index = document.inner.table_count();
        // A style may add its definition, so only a plain table splices.
        if style.is_none() {
            let appended = rdocx::Document::detached_table(rows, cols);
            let slot = document.inner.content_count();
            if let Ok((xml, snapshots)) = document.inner.detached_table_owner_snapshots(&appended)
                && let Some(plan) = document.plan_body_splice(BodySplice {
                    slot,
                    removed: 0,
                    inserted: Some(SplicedBodyItem::Table(xml, snapshots)),
                })?
            {
                // Every fallible step precedes the in-place native append.
                let publication = crate::document::WordPublication::AppendBody;
                let prepared =
                    document.prepare_body_splice(plan, "Document.add_table", &publication)?;
                let path = document
                    .revisions
                    .capture_prepared(&prepared.revision, smallvec![PathSeg::Body(index)])
                    .map_err(revision_to_pyerr)?;
                let result = Py::new(
                    py,
                    PyTable::new(slf.clone_ref(py), path)
                        .bind_prepared_splice(&document, &prepared)?,
                )?;
                document.inner.add_table(rows, cols);
                debug_assert_eq!(prepared.plan.table, index);
                document.commit_body_splice(prepared, &publication)?;
                return Ok(result);
            }
        }
        let mut candidate = document.inner.clone_for_staging();
        let style = style
            .map(|value| crate::table::checked_table_style(py, &mut candidate, value))
            .transpose()?;
        let mut table = candidate.add_table(rows, cols);
        if let Some(style) = style {
            table.set_style(&style);
        }
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_table",
            crate::document::WordPublication::AppendBody,
        )?;
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, smallvec![PathSeg::Body(index)])
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyTable::new(slf.clone_ref(py), path)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    fn remove_content(&mut self, py: Python<'_>, index: usize) -> PyResult<bool> {
        if self.inner.plain_direct_paragraph(index)
            && let Some((range, revision)) = self.prepare_tail_body_removal(index)?
        {
            let removed = py
                .detach(|| self.inner.try_remove_content(index))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            if removed {
                self.commit_tail_body_removal(index, range, revision)?;
            }
            return Ok(removed);
        }
        if let Some(plan) = self.plan_body_splice(BodySplice {
            slot: index,
            removed: 1,
            inserted: None,
        })? {
            let publication = crate::document::WordPublication::Conservative;
            let prepared =
                self.prepare_body_splice(plan, "Document.remove_content", &publication)?;
            let removed = py
                .detach(|| self.inner.try_remove_content(index))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            if removed {
                self.commit_body_splice(prepared, &publication)?;
            }
            return Ok(removed);
        }
        let mut candidate = self.inner.clone_for_staging();
        let removed = py
            .detach(|| candidate.try_remove_content(index))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        if removed {
            let prepared = self.prepare_word_revision(
                py,
                &candidate,
                "Document.remove_content",
                crate::document::WordPublication::Conservative,
            )?;
            self.commit_word_revision(prepared)?;
            self.inner = candidate;
        }
        Ok(removed)
    }

    fn remove_content_range(
        &mut self,
        py: Python<'_>,
        start: usize,
        end: usize,
    ) -> PyResult<usize> {
        let len = self.inner.content_count();
        if start > end || end > len {
            return Err(rdocx_to_pyerr(
                py,
                rdocx::Error::Other("content range is reversed or out of range".to_owned()),
            ));
        }
        if start == end {
            return Ok(0);
        }
        if let Some(plan) = self.plan_body_splice(BodySplice {
            slot: start,
            removed: end - start,
            inserted: None,
        })? {
            let publication = crate::document::WordPublication::Conservative;
            let prepared =
                self.prepare_body_splice(plan, "Document.remove_content_range", &publication)?;
            let removed = py
                .detach(|| self.inner.remove_content_range(start, end))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            self.commit_body_splice(prepared, &publication)?;
            return Ok(removed);
        }
        let mut candidate = self.inner.clone_for_staging();
        let removed = py
            .detach(|| candidate.remove_content_range(start, end))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = self.prepare_word_revision(
            py,
            &candidate,
            "Document.remove_content_range",
            crate::document::WordPublication::Conservative,
        )?;
        self.commit_word_revision(prepared)?;
        self.inner = candidate;
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

    #[pyo3(signature = (index, text, style = None))]
    pub(crate) fn insert_paragraph(
        slf: Py<Self>,
        py: Python<'_>,
        index: usize,
        text: &str,
        style: Option<&str>,
    ) -> PyResult<Py<PyParagraph>> {
        let result = {
            let mut document = slf.borrow_mut(py);
            if index > document.inner.content_count() {
                return Err(PyIndexError::new_err("content index out of range"));
            }
            document.refuse_unsavable_body_edit(py)?;
            let style = style
                .map(|value| style_id_of_type(&document.inner, value, rdocx::StyleType::Paragraph))
                .transpose()?;
            // The same edit builds the paragraph read for the inventory and
            // the paragraph placed.
            let edit = |paragraph: &mut rdocx::Paragraph<'_>| {
                if let Some(style) = &style {
                    paragraph.set_style(style);
                }
            };
            let mut inserted = rdocx::Document::detached_text_paragraph(text);
            edit(&mut rdocx::Paragraph::detached(&mut inserted));
            let splice = document.inner.direct_body_paragraph_facts(&inserted);
            if let Ok(Some(facts)) = splice
                && let Some(plan) = document.plan_body_splice(BodySplice {
                    slot: index,
                    removed: 0,
                    inserted: Some(SplicedBodyItem::Paragraph(facts.0, facts.1)),
                })?
            {
                // Every fallible step precedes the in-place native insertion.
                let publication = crate::document::WordPublication::Conservative;
                let prepared = document.prepare_body_splice(
                    plan,
                    "Document.insert_paragraph",
                    &publication,
                )?;
                let path = document
                    .revisions
                    .capture_prepared(
                        &prepared.revision,
                        smallvec![PathSeg::Body(0), PathSeg::Para(prepared.plan.paragraph)],
                    )
                    .map_err(revision_to_pyerr)?;
                let result = Py::new(
                    py,
                    PyParagraph::new(slf.clone_ref(py), path)
                        .bind_prepared_splice(&document, &prepared)?,
                )?;
                edit(&mut document.inner.insert_paragraph(index, text));
                debug_assert_eq!(
                    document.inner.paragraph_index_of_content(index),
                    Some(prepared.plan.paragraph)
                );
                document.commit_body_splice(prepared, &publication)?;
                return Ok(result);
            }
            let mut candidate = document.inner.clone_for_staging();
            edit(&mut candidate.insert_paragraph(index, text));
            let paragraph = candidate
                .paragraph_index_of_content(index)
                .expect("an inserted body paragraph has a paragraph index");
            let prepared = document.prepare_word_revision(
                py,
                &candidate,
                "Document.insert_paragraph",
                crate::document::WordPublication::Conservative,
            )?;
            let path = document
                .revisions
                .capture_prepared(
                    &prepared.revision,
                    smallvec![PathSeg::Body(0), PathSeg::Para(paragraph)],
                )
                .map_err(revision_to_pyerr)?;
            let result = Py::new(
                py,
                PyParagraph::new(slf.clone_ref(py), path)
                    .bind_prepared(py, &document, &prepared, &candidate, None)?,
            )?;
            document.commit_word_revision(prepared)?;
            document.inner = candidate;
            result
        };
        Ok(result)
    }

    #[pyo3(signature = (index, rows, cols))]
    fn insert_table(
        slf: Py<Self>,
        py: Python<'_>,
        index: usize,
        rows: usize,
        cols: usize,
    ) -> PyResult<Py<PyTable>> {
        let result = {
            let mut document = slf.borrow_mut(py);
            if index > document.inner.content_count() {
                return Err(PyIndexError::new_err("content index out of range"));
            }
            document.refuse_unsavable_body_edit(py)?;
            let inserted = rdocx::Document::detached_table(rows, cols);
            if let Ok((xml, snapshots)) = document.inner.detached_table_owner_snapshots(&inserted)
                && let Some(plan) = document.plan_body_splice(BodySplice {
                    slot: index,
                    removed: 0,
                    inserted: Some(SplicedBodyItem::Table(xml, snapshots)),
                })?
            {
                // Every fallible step precedes the in-place native insertion.
                let publication = crate::document::WordPublication::Conservative;
                let prepared =
                    document.prepare_body_splice(plan, "Document.insert_table", &publication)?;
                let path = document
                    .revisions
                    .capture_prepared(
                        &prepared.revision,
                        smallvec![PathSeg::Body(prepared.plan.table)],
                    )
                    .map_err(revision_to_pyerr)?;
                let result = Py::new(
                    py,
                    PyTable::new(slf.clone_ref(py), path)
                        .bind_prepared_splice(&document, &prepared)?,
                )?;
                document.inner.insert_table(index, rows, cols);
                debug_assert_eq!(
                    document.inner.content_index_of_table(prepared.plan.table),
                    Some(index)
                );
                document.commit_body_splice(prepared, &publication)?;
                return Ok(result);
            }
            let mut candidate = document.inner.clone_for_staging();
            candidate.insert_table(index, rows, cols);
            // Preserve the direct body table ordinal while capturing its physical owner.
            let table = (0..candidate.table_count())
                .find(|table| candidate.content_index_of_table(*table) == Some(index))
                .expect("an inserted body table has a table index");
            let prepared = document.prepare_word_revision(
                py,
                &candidate,
                "Document.insert_table",
                crate::document::WordPublication::Conservative,
            )?;
            let path = document
                .revisions
                .capture_prepared(&prepared.revision, smallvec![PathSeg::Body(table)])
                .map_err(revision_to_pyerr)?;
            let result = Py::new(
                py,
                PyTable::new(slf.clone_ref(py), path)
                    .bind_prepared(py, &document, &prepared, &candidate, None)?,
            )?;
            document.commit_word_revision(prepared)?;
            document.inner = candidate;
            result
        };
        Ok(result)
    }

    fn pop_content(
        slf: Py<Self>,
        py: Python<'_>,
        index: &Bound<'_, PyAny>,
    ) -> PyResult<Py<PyContentFragment>> {
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
        let mut document = slf.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let fragment = candidate
            .remove_content_at(&location)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.pop_content",
            crate::document::WordPublication::Conservative,
        )?;
        let result = Py::new(py, PyContentFragment { inner: fragment })?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
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
        let mut document = slf.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .insert_content(&location, fragment)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.insert_content",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
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
        let mut document = slf.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .clone_content(&source, &destination)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.clone_content",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
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
        let mut document = slf.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .move_content(&source, &destination)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.move_content",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }

    #[getter]
    fn fonts(&self, py: Python<'_>) -> PyResult<Py<PyTuple>> {
        let fonts = self.inner.try_fonts().map_err(|e| rdocx_to_pyerr(py, e))?;
        let mut snapshots = Vec::new();
        snapshots.try_reserve_exact(fonts.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("font listing allocation failed")
        })?;
        for font in fonts {
            snapshots.push(PyFontDefinition::snapshot(py, font)?);
        }
        Ok(PyTuple::new(py, snapshots)?.unbind())
    }
    fn set_font(&mut self, py: Python<'_>, font: PyRef<'_, PyFontDefinition>) -> PyResult<()> {
        if !font.embedded_fonts.bind(py).is_empty() {
            return Err(PyValueError::new_err(
                "set_font accepts descriptive records only, use embed_font",
            ));
        }
        let native = rdocx::FontDefinition {
            name: font.name.clone(),
            alternate_name: font.alternate_name.clone(),
            family: font.family.clone(),
            pitch: font.pitch.clone(),
            embedded_fonts: Vec::new(),
        };
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_font(native)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.set_font")
    }
    fn remove_font(
        &mut self,
        py: Python<'_>,
        name: &str,
    ) -> PyResult<Option<Py<PyFontDefinition>>> {
        let mut candidate = self.inner.clone_for_staging();
        let removed = candidate
            .remove_font(name)
            .map_err(|e| rdocx_to_pyerr(py, e))?
            .map(|font| PyFontDefinition::snapshot(py, font))
            .transpose()?;
        self.publish_api_candidate(py, candidate, "Document.remove_font")?;
        Ok(removed)
    }
    #[pyo3(signature = (name, kind="regular"))]
    fn remove_embedded_font(
        &mut self,
        py: Python<'_>,
        name: &str,
        kind: &str,
    ) -> PyResult<Option<Py<PyEmbeddedFont>>> {
        let kind = checked_font_kind(kind)?;
        let mut candidate = self.inner.clone_for_staging();
        let removed = candidate
            .remove_embedded_font(name, kind)
            .map_err(|e| rdocx_to_pyerr(py, e))?
            .map(|font| PyEmbeddedFont::snapshot(py, font))
            .transpose()?;
        self.publish_api_candidate(py, candidate, "Document.remove_embedded_font")?;
        Ok(removed)
    }
    #[pyo3(signature = (name, font, kind="regular", *, license, font_key=None, subsetted=false))]
    #[allow(clippy::too_many_arguments)] // Explicit operands match the native font authoring contract.
    fn embed_font(
        &mut self,
        py: Python<'_>,
        name: &str,
        font: &Bound<'_, PyAny>,
        kind: &str,
        license: PyRef<'_, PyFontEmbeddingLicense>,
        font_key: Option<String>,
        subsetted: bool,
    ) -> PyResult<()> {
        let kind = checked_font_kind(kind)?;
        let license = checked_font_license(&license)?;
        let data = checked_font_input(font)?;
        let key = match font_key {
            Some(key) => key,
            None => rdocx::derive_font_key(&data).map_err(|e| rdocx_to_pyerr(py, e))?,
        };
        let font = rdocx::EmbeddedFont::new(kind, data, key, license).with_subsetted(subsetted);
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .embed_font(name, font)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.embed_font")
    }

    #[getter]
    fn embed_true_type_fonts(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.inner
            .embed_true_type_fonts()
            .map_err(|e| rdocx_to_pyerr(py, e))
    }
    #[setter]
    fn set_embed_true_type_fonts(&mut self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_embed_true_type_fonts(value)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.set_embed_true_type_fonts")
    }
    #[getter]
    fn embed_system_fonts(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.inner
            .embed_system_fonts()
            .map_err(|e| rdocx_to_pyerr(py, e))
    }
    #[setter]
    fn set_embed_system_fonts(&mut self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_embed_system_fonts(value)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.set_embed_system_fonts")
    }
    #[getter]
    fn save_subset_fonts(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.inner
            .save_subset_fonts()
            .map_err(|e| rdocx_to_pyerr(py, e))
    }
    #[setter]
    fn set_save_subset_fonts(&mut self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let mut candidate = self.inner.clone_for_staging();
        candidate
            .set_save_subset_fonts(value)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        self.publish_api_candidate(py, candidate, "Document.set_save_subset_fonts")
    }
}

fn parse_render_options(
    revision_view: &str,
    revision_style: &str,
) -> PyResult<rdocx::RenderOptions> {
    let revision_view = match revision_view {
        "accepted" => rdocx::RevisionView::Accepted,
        "tracked" => rdocx::RevisionView::Tracked,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown revision view {other:?}, expected accepted or tracked"
            )));
        }
    };
    let revision_style = match revision_style {
        "plain" => rdocx::RevisionStyle::Plain,
        "legal" => rdocx::RevisionStyle::Legal,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown revision style {other:?}, expected plain or legal"
            )));
        }
    };
    if revision_style != rdocx::RevisionStyle::Plain
        && revision_view != rdocx::RevisionView::Tracked
    {
        return Err(PyValueError::new_err(
            "revision_style applies only with revision_view=\"tracked\"",
        ));
    }
    Ok(rdocx::RenderOptions {
        revision_view,
        revision_style,
    })
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

// Native snapshots bind causal state without changing their visible value equality.

impl PyDocument {
    fn owned_story_snapshot(&self, story: &rdocx::StoryId) -> PyResult<PyStory> {
        let mut snapshot = story_snapshot(story);
        snapshot.binding = Some(self.capture_snapshot(smallvec![])?);
        Ok(snapshot)
    }

    fn validate_owned_story(&self, py: Python<'_>, story: &PyStory) -> PyResult<()> {
        let Some(path) = &story.binding else {
            return Ok(());
        };
        if !path.belongs_to(&self.revisions) {
            return Err(PyValueError::new_err(
                "story belongs to a different document",
            ));
        }
        path.validate_scopes(
            &self.revisions,
            "story",
            "Re-fetch it with document.stories.",
        )
        .map_err(|error| crate::stale_to_pyerr(py, error))
    }

    fn validate_owned_item(&self, py: Python<'_>, item: &PyStoryItem) -> PyResult<()> {
        if let Some(path) = &item.binding {
            if !path.belongs_to(&self.revisions) {
                return Err(PyValueError::new_err(
                    "story item belongs to a different document",
                ));
            }
            path.validate_scopes(
                &self.revisions,
                "story item",
                "Re-fetch it with document.story_items.",
            )
            .map_err(|error| crate::stale_to_pyerr(py, error))?;
        }
        self.validate_owned_story(py, &item.story)
    }
}

impl PartialEq for PyStory {
    fn eq(&self, other: &Self) -> bool {
        self.kind == other.kind
            && self.part_name == other.part_name
            && self.owner_index == other.owner_index
    }
}
impl Eq for PyStory {}
impl PartialEq for PyStoryItem {
    fn eq(&self, other: &Self) -> bool {
        self.story == other.story
            && self.kind == other.kind
            && self.index_path == other.index_path
            && self.direct_body_index == other.direct_body_index
            && self.text == other.text
            && self.xml == other.xml
            && self.revision == other.revision
    }
}
impl Eq for PyStoryItem {}

fn collect_caller_fonts<'py>(
    py: Python<'py>,
    fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
    font_dir: Option<PathBuf>,
    font_dirs: Option<Bound<'py, PyAny>>,
) -> PyResult<Option<Vec<oxml_layout::FontFile>>> {
    if fonts.is_none() && font_dir.is_none() && font_dirs.is_none() {
        return Ok(None);
    }
    if font_dir.is_some() && font_dirs.is_some() {
        return Err(PyValueError::new_err(
            "font_dir and font_dirs cannot be supplied together",
        ));
    }
    let explicit = fonts
        .unwrap_or_default()
        .into_iter()
        .map(|(family, data)| oxml_layout::FontFile {
            family,
            data: data.as_bytes().to_vec(),
        })
        .collect::<Vec<_>>();
    let directories = match font_dirs {
        Some(paths) => {
            if paths.is_instance_of::<PyString>() || paths.is_instance_of::<PyBytes>() {
                return Err(PyTypeError::new_err(
                    "font_dirs must be a sequence of paths, not a scalar string or bytes",
                ));
            }
            paths
                .extract::<Vec<PathBuf>>()
                .map_err(|_| PyTypeError::new_err("font_dirs must be a sequence of paths"))?
        }
        None => font_dir.into_iter().collect(),
    };
    py.detach(|| {
        for directory in &directories {
            let metadata = std::fs::metadata(directory).map_err(|error| {
                std::io::Error::new(
                    error.kind(),
                    format!("font directory {}: {error}", directory.display()),
                )
            })?;
            if !metadata.is_dir() {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::NotADirectory,
                    format!("font directory {} is not a directory", directory.display()),
                ));
            }
        }
        Ok::<_, std::io::Error>(())
    })
    .map_err(PyErr::from)?;
    let selected = py
        .detach(|| {
            let mut selected = Vec::new();
            for directory in directories {
                selected.extend(oxml_layout::FontFile::from_dir(&directory)?);
            }
            selected.extend(explicit);
            Ok::<_, oxml_layout::LayoutError>(selected)
        })
        .map_err(|error| rdocx_to_pyerr(py, rdocx::Error::Layout(error)))?;
    Ok((!selected.is_empty()).then_some(selected))
}

#[pyclass(name = "OdtDiagnostic", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyOdtDiagnostic {
    pub path: String,
    pub message: String,
}

#[pyclass(name = "RtfDiagnostic", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyRtfDiagnostic {
    pub offset: usize,
    pub destination: Option<String>,
    pub message: String,
}

#[pyclass(name = "EpubDiagnostic", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyEpubDiagnostic {
    pub path: String,
    pub message: String,
}

#[pyclass(name = "TextDiagnostic", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyTextDiagnostic {
    pub message: String,
}

#[pyclass(name = "ValidationFinding", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyValidationFinding {
    pub severity: String,
    pub message: String,
}

#[pyclass(name = "OdtExportResult", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyOdtExportResult {
    bytes: Vec<u8>,
    diagnostics: Vec<PyOdtDiagnostic>,
}
#[pymethods]
impl PyOdtExportResult {
    #[getter]
    fn bytes<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.bytes)
    }
    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.diagnostics.iter().cloned())
    }
}

#[pyclass(name = "RtfExportResult", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyRtfExportResult {
    bytes: Vec<u8>,
    diagnostics: Vec<PyRtfDiagnostic>,
}
#[pymethods]
impl PyRtfExportResult {
    #[getter]
    fn bytes<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.bytes)
    }
    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.diagnostics.iter().cloned())
    }
}

#[pyclass(name = "EpubExportResult", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyEpubExportResult {
    bytes: Vec<u8>,
    diagnostics: Vec<PyEpubDiagnostic>,
}
#[pymethods]
impl PyEpubExportResult {
    #[getter]
    fn bytes<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.bytes)
    }
    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.diagnostics.iter().cloned())
    }
}

#[pyclass(name = "TextReport", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyTextReport {
    text: String,
    scope: String,
    revision_view: String,
    diagnostics: Vec<PyTextDiagnostic>,
}
#[pymethods]
impl PyTextReport {
    #[getter]
    fn text(&self) -> &str {
        &self.text
    }
    #[getter]
    fn scope(&self) -> &str {
        &self.scope
    }
    #[getter]
    fn revision_view(&self) -> &str {
        &self.revision_view
    }
    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.diagnostics.iter().cloned())
    }
}
impl From<rdocx::TextReport> for PyTextReport {
    fn from(report: rdocx::TextReport) -> Self {
        Self {
            text: report.text,
            scope: report.scope.to_owned(),
            revision_view: report.revision_view.to_owned(),
            diagnostics: report
                .diagnostics
                .into_iter()
                .map(|d| PyTextDiagnostic { message: d.message })
                .collect(),
        }
    }
}
#[pyclass(name = "ValidationReport", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyValidationReport {
    findings: Vec<PyValidationFinding>,
}
impl From<rdocx::ValidationReport> for PyValidationReport {
    fn from(report: rdocx::ValidationReport) -> Self {
        Self {
            findings: report
                .findings
                .into_iter()
                .map(|finding| PyValidationFinding {
                    severity: match finding.severity {
                        rdocx::ValidationSeverity::Error => "error",
                        rdocx::ValidationSeverity::Warning => "warning",
                    }
                    .to_owned(),
                    message: finding.message,
                })
                .collect(),
        }
    }
}
#[pymethods]
impl PyValidationReport {
    #[getter]
    fn findings<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.findings.iter().cloned())
    }
    #[getter]
    fn is_valid(&self) -> bool {
        self.findings
            .iter()
            .all(|finding| finding.severity != "error")
    }
}
#[pyfunction]
pub fn validate_bytes(data: &[u8], py: Python<'_>) -> PyResult<PyValidationReport> {
    py.detach(|| rdocx::validate_document_bytes(data))
        .map(Into::into)
        .map_err(|error| rdocx_to_pyerr(py, error))
}
#[pyfunction]
pub fn validate_file(path: PathBuf, py: Python<'_>) -> PyResult<PyValidationReport> {
    py.detach(|| rdocx::validate_document_file(path))
        .map(Into::into)
        .map_err(|error| rdocx_to_pyerr(py, error))
}

fn type_name(value: &Bound<'_, PyAny>) -> String {
    value
        .get_type()
        .name()
        .map(|name| name.to_string())
        .unwrap_or_else(|_| "object".to_owned())
}

fn is_mapping(value: &Bound<'_, PyAny>) -> PyResult<bool> {
    Ok(value.is_instance_of::<PyDict>())
}

/// The deepest nesting of dicts and lists template data may use.
const MAX_TEMPLATE_DATA_DEPTH: usize = 256;

/// Caller-facing template refusals are `ValueError`. Other native failures
/// keep the shared error class.
fn template_render_pyerr(py: Python<'_>, error: rdocx::Error) -> PyErr {
    if let rdocx::Error::Other(message) = &error
        && (message.starts_with("invalid template:")
            || message.starts_with("unsupported template grammar version "))
    {
        return PyValueError::new_err(message.clone());
    }
    rdocx_to_pyerr(py, error)
}

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
        for (key, item) in value.cast::<PyDict>()?.iter() {
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
    let items = if let Ok(list) = value.cast::<PyList>() {
        list.iter().collect::<Vec<_>>()
    } else {
        value.cast::<PyTuple>()?.iter().collect::<Vec<_>>()
    };
    for (index, item) in items.iter().enumerate() {
        array.push(template_value(item, &format!("{path}[{index}]"), open)?);
    }
    Ok(Value::Array(array))
}

/// Caller authorization is required independently of the font's own fsType bits.
#[pyclass(
    name = "FontEmbeddingLicense",
    frozen,
    get_all,
    eq,
    skip_from_py_object
)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyFontEmbeddingLicense {
    pub authorized: bool,
    pub identity: String,
}
#[pymethods]
impl PyFontEmbeddingLicense {
    #[new]
    #[pyo3(signature = (*, authorized, identity))]
    fn new(authorized: bool, identity: String) -> Self {
        Self {
            authorized,
            identity,
        }
    }
}

fn checked_font_input(font: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    use std::io::Read;
    const LIMIT: usize = 64 * 1024 * 1024;
    let copy = |source: &[u8]| -> PyResult<Vec<u8>> {
        if source.len() > LIMIT {
            return Err(PyValueError::new_err("font exceeds 64 MiB input bound"));
        }
        let mut data = Vec::new();
        data.try_reserve_exact(source.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("font input allocation failed")
        })?;
        data.extend_from_slice(source);
        Ok(data)
    };
    if let Ok(bytes) = font.cast::<PyBytes>() {
        return copy(bytes.as_bytes());
    }
    if let Ok(path) = font.extract::<PathBuf>() {
        let file = std::fs::File::open(path)
            .map_err(|e| pyo3::exceptions::PyOSError::new_err(e.to_string()))?;
        if file
            .metadata()
            .map_err(|e| pyo3::exceptions::PyOSError::new_err(e.to_string()))?
            .len()
            > LIMIT as u64
        {
            return Err(PyValueError::new_err("font exceeds 64 MiB input bound"));
        }
        let mut data = Vec::new();
        let mut input = file.take((LIMIT + 1) as u64);
        let mut chunk = [0u8; 8192];
        loop {
            let count = input
                .read(&mut chunk)
                .map_err(|e| pyo3::exceptions::PyOSError::new_err(e.to_string()))?;
            if count == 0 {
                break;
            }
            if count > LIMIT - data.len() {
                return Err(PyValueError::new_err("font exceeds 64 MiB input bound"));
            }
            data.try_reserve_exact(count).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("font input allocation failed")
            })?;
            data.extend_from_slice(&chunk[..count]);
        }
        return Ok(data);
    }
    let read = font.getattr("read").map_err(|_| {
        PyTypeError::new_err("font must be bytes, a filesystem path or a binary reader")
    })?;
    let mut data = Vec::new();
    loop {
        let remaining = LIMIT - data.len();
        let value = read.call1(((remaining + 1).min(8192),))?;
        let bytes = value
            .cast::<PyBytes>()
            .map_err(|_| PyTypeError::new_err("font reader must return bytes"))?;
        let chunk = bytes.as_bytes();
        if chunk.is_empty() {
            return Ok(data);
        }
        if chunk.len() > remaining {
            return Err(PyValueError::new_err("font exceeds 64 MiB input bound"));
        }
        let needed = data.len() + chunk.len();
        if needed > data.capacity() {
            let capacity = data
                .capacity()
                .saturating_mul(2)
                .clamp(8192, LIMIT)
                .max(needed);
            data.try_reserve_exact(capacity - data.len()).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("font input allocation failed")
            })?;
        }
        data.extend_from_slice(chunk);
    }
}

fn font_kind_name(kind: rdocx::EmbeddedFontKind) -> &'static str {
    match kind {
        rdocx::EmbeddedFontKind::Regular => "regular",
        rdocx::EmbeddedFontKind::Bold => "bold",
        rdocx::EmbeddedFontKind::Italic => "italic",
        rdocx::EmbeddedFontKind::BoldItalic => "bold_italic",
    }
}
fn checked_font_kind(kind: &str) -> PyResult<rdocx::EmbeddedFontKind> {
    match kind {
        "regular" => Ok(rdocx::EmbeddedFontKind::Regular),
        "bold" => Ok(rdocx::EmbeddedFontKind::Bold),
        "italic" => Ok(rdocx::EmbeddedFontKind::Italic),
        "bold_italic" => Ok(rdocx::EmbeddedFontKind::BoldItalic),
        _ => Err(PyValueError::new_err(
            "font kind must be regular, bold, italic or bold_italic",
        )),
    }
}
fn checked_font_license(license: &PyFontEmbeddingLicense) -> PyResult<rdocx::FontEmbeddingLicense> {
    if !license.authorized || license.identity.trim().is_empty() {
        return Err(PyValueError::new_err(
            "font embedding requires explicit authorization and nonempty license identity",
        ));
    }
    Ok(rdocx::FontEmbeddingLicense::new(
        license.authorized,
        license.identity.clone(),
    ))
}

#[pyclass(name = "EmbeddedFont", frozen, get_all, skip_from_py_object)]
pub struct PyEmbeddedFont {
    pub kind: &'static str,
    pub data: Py<PyBytes>,
    pub font_key: String,
    pub subsetted: bool,
    pub license: Py<PyFontEmbeddingLicense>,
}
impl PyEmbeddedFont {
    fn snapshot(py: Python<'_>, font: rdocx::EmbeddedFont) -> PyResult<Py<Self>> {
        let data = PyBytes::new_with(py, font.data.len(), |out| {
            out.copy_from_slice(&font.data);
            Ok(())
        })?
        .unbind();
        let license = Py::new(
            py,
            PyFontEmbeddingLicense {
                authorized: font.license.authorized,
                identity: font.license.identity,
            },
        )?;
        Py::new(
            py,
            Self {
                kind: font_kind_name(font.kind),
                data,
                font_key: font.font_key,
                subsetted: font.subsetted,
                license,
            },
        )
    }
}
#[pyclass(name = "FontDefinition", frozen, get_all, skip_from_py_object)]
pub struct PyFontDefinition {
    pub name: String,
    pub alternate_name: Option<String>,
    pub family: Option<String>,
    pub pitch: Option<String>,
    pub embedded_fonts: Py<PyTuple>,
}
#[pymethods]
impl PyFontDefinition {
    #[new]
    #[pyo3(signature = (name, *, alternate_name=None, family=None, pitch=None))]
    fn new(
        py: Python<'_>,
        name: String,
        alternate_name: Option<String>,
        family: Option<String>,
        pitch: Option<String>,
    ) -> PyResult<Self> {
        Ok(Self {
            name,
            alternate_name,
            family,
            pitch,
            embedded_fonts: PyTuple::new(py, Vec::<Py<PyEmbeddedFont>>::new())?.unbind(),
        })
    }
}
impl PyFontDefinition {
    fn snapshot(py: Python<'_>, font: rdocx::FontDefinition) -> PyResult<Py<Self>> {
        let mut faces = Vec::new();
        faces
            .try_reserve_exact(font.embedded_fonts.len())
            .map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("font snapshot allocation failed")
            })?;
        for font in font.embedded_fonts {
            faces.push(PyEmbeddedFont::snapshot(py, font)?);
        }
        let embedded_fonts = PyTuple::new(py, faces)?.unbind();
        Py::new(
            py,
            Self {
                name: font.name,
                alternate_name: font.alternate_name,
                family: font.family,
                pitch: font.pitch,
                embedded_fonts,
            },
        )
    }
}

#[cfg(test)]
fn fx237_legacy_word_owner_locations(
    document: &rdocx::Document,
) -> rdocx::Result<Vec<WordOwnerAddress>> {
    let mut result = Vec::new();
    let mut table_ordinals = HashMap::new();
    let table_facts = word_table_read_facts(document);
    let snapshots = document.story_item_snapshots()?;
    let mut body_slots = HashMap::new();
    let mut cell_parents = HashMap::new();
    for snapshot in &snapshots {
        let location = snapshot.location();
        if location.story().kind() == rdocx::StoryKind::Body
            && snapshot.is_direct_child()
            && matches!(
                location.item_kind(),
                rdocx::StoryItemKind::Paragraph | rdocx::StoryItemKind::ContentControl
            )
        {
            body_slots.insert(location.clone(), snapshot.direct_body_index());
        }
    }
    for snapshot in &snapshots {
        if snapshot.location().item_kind() != rdocx::StoryItemKind::Table {
            continue;
        }
        let location = snapshot.location().clone();
        let ordinal = table_ordinals
            .entry(location.story().clone())
            .or_insert(0usize);
        let projection_ordinal = Some(*ordinal);
        *ordinal += 1;
        let table_index = result.len();
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Table(location.clone()),
            parent: None,
            projection_ordinal,
            read_facts: table_facts.get(&location).cloned().unwrap_or_default(),
            body_slot: if snapshot.is_direct_child() {
                snapshot.direct_body_index()
            } else {
                None
            },
        });
        let cells = document.read_story_table(&location, |table| {
            (0..table.row_count())
                .flat_map(|row| {
                    let count = table.row(row).map_or(0, |row| row.cell_count());
                    (0..count).map(move |cell| (row, cell))
                })
                .collect::<Vec<_>>()
        })?;
        for (row, cell) in cells {
            let story = document.story_table_cell(&location, row, cell)?;
            result.push(WordOwnerAddress {
                location: WordOwnerLocation::Cell {
                    table: location.clone(),
                    row,
                    cell,
                    story,
                },
                parent: Some(table_index),
                body_slot: None,
                projection_ordinal: None,
                read_facts: WordReadFacts::None,
            });
        }
    }
    for (index, address) in result.iter().enumerate() {
        if let WordOwnerLocation::Cell { story, .. } = &address.location {
            cell_parents
                .entry((
                    story.kind(),
                    story.part_name().to_owned(),
                    story.owner_index(),
                ))
                .or_insert(index);
        }
    }
    // Join nested table routes through their actual physical cell owner. An
    // unresolved route stays conservative rather than acquiring an ordinal parent.
    for address in &mut result {
        if let WordOwnerLocation::Table(location) = &address.location {
            let story = location.story();
            let parent = cell_parents
                .get(&(
                    story.kind(),
                    story.part_name().to_owned(),
                    story.owner_index(),
                ))
                .copied();
            address.parent = parent;
        }
    }
    let mut paragraph_ordinals = HashMap::new();
    for (location, xml) in document.story_range_paragraphs()? {
        let story = location.story();
        let parent = cell_parents
            .get(&(
                story.kind(),
                story.part_name().to_owned(),
                story.owner_index(),
            ))
            .copied();
        let ordinal = paragraph_ordinals
            .entry(location.story().clone())
            .or_insert(0usize);
        let projection_ordinal = Some(*ordinal);
        *ordinal += 1;
        let index = result.len();
        let paragraph = rdocx_oxml::CT_P::from_xml_fragment(&xml)?;
        // Reuse the containing item's fact from this exact immutable source
        // inventory, matching the native paragraph snapshot's containing route.
        let body_slot = if location.story().kind() == rdocx::StoryKind::Body {
            let containing = match location.index_path() {
                [_] => Some(location.clone()),
                [control, _] => Some(rdocx::ContentLocation::new(
                    location.story().clone(),
                    rdocx::StoryItemKind::ContentControl,
                    vec![*control],
                )),
                _ => None,
            };
            containing.and_then(|containing| body_slots.get(&containing).copied().flatten())
        } else {
            None
        };
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Paragraph(location.clone()),
            parent,
            body_slot,
            projection_ordinal,
            read_facts: WordReadFacts::Paragraph(paragraph.accepted_text()),
        });
        for path in paragraph.accepted_run_paths() {
            result.push(WordOwnerAddress {
                location: WordOwnerLocation::Run(location.clone(), path),
                parent: Some(index),
                body_slot: None,
                projection_ordinal: None,
                read_facts: WordReadFacts::None,
            });
        }
    }
    for index in 0..document.section_count() {
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Section(index),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        });
    }
    let mut local_indexes = HashMap::<rdocx::ContentLocation, usize>::new();
    let mut story_indexes = HashMap::<rdocx::StoryId, usize>::new();
    for (location, _) in document.story_link_snapshots()? {
        let local = local_indexes.entry(location.clone()).or_default();
        let position = story_indexes.entry(location.story().clone()).or_default();
        result.push(WordOwnerAddress {
            location: WordOwnerLocation::Hyperlink(location.clone(), *local),
            parent: None,
            body_slot: None,
            projection_ordinal: Some(*position),
            read_facts: WordReadFacts::None,
        });
        *local += 1;
        *position += 1;
    }
    Ok(result)
}

#[cfg(test)]
fn fx237_legacy_matches(
    old: &[WordOwner],
    new: &[WordOwnerAddress],
    publication: &WordPublication,
) -> Vec<Option<usize>> {
    new.iter()
        .enumerate()
        .map(|(after, address)| {
            old.iter().enumerate().find_map(|(before, owner)| {
                if let (
                    WordPublication::HyperlinkRemove(WordOwnerLocation::Hyperlink(
                        removed,
                        position,
                    )),
                    WordOwnerLocation::Hyperlink(location, candidate),
                    WordOwnerLocation::Hyperlink(previous, index),
                ) = (publication, &address.location, &owner.location)
                    && same_content_owner(removed, location)
                {
                    let expected = if candidate >= position {
                        candidate.checked_add(1)
                    } else {
                        Some(*candidate)
                    };
                    return (same_content_owner(location, previous) && expected == Some(*index))
                        .then_some(before);
                }
                same_inventory_route(old, before, new, after).then_some(before)
            })
        })
        .collect()
}

#[cfg(test)]
fn fx237_test_owners(locations: Vec<WordOwnerAddress>) -> Vec<WordOwner> {
    let scopes = locations
        .iter()
        .flat_map(|address| address.location.scopes().iter().copied())
        .collect::<Vec<_>>();
    let states = RevisionCounter::new().initialize_scopes(&scopes).unwrap();
    word_owners_with_states(locations, &states)
}

#[cfg(test)]
fn fx237_control_table() -> rdocx_oxml::content_control::CT_Sdt {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>same</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt></w:body></w:document>"#;
    let mut parsed = rdocx_oxml::document::CT_Document::from_xml(xml).unwrap();
    let rdocx_oxml::document::BodyContent::ContentControl(control) = parsed.body.content.remove(0)
    else {
        unreachable!()
    };
    *control
}

#[test]
fn fx237_combined_registry_matches_legacy_success_and_refusal() {
    Python::initialize();
    Python::attach(|py| {
        for content in [
            "<w:p><w:hyperlink r:id='missing'><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>",
            "<w:p><w:hyperlink w:anchor='inside'><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>",
            "<w:p><f:hyperlink xmlns:f='urn:foreign'><w:r><w:t>opaque</w:t></w:r></f:hyperlink></w:p>",
            "<w:tbl><w:tblGrid><w:gridCol w:w='1000'/></w:tblGrid><w:tr><w:tc><w:tcPr><f:keep xmlns:f='urn:opaque'/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ] {
            let seed = rdocx::Document::new().to_bytes().unwrap();
            let xml = format!(
                "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><w:body>{content}<w:sectPr/></w:body></w:document>"
            );
            let scope = PyDict::new(py);
            scope.set_item("seed", PyBytes::new(py, &seed)).unwrap();
            scope.set_item("xml", xml).unwrap();
            // Author raw package input independently of checked native insertion,
            // so a dangling relationship actually reaches both inventory paths.
            py.run(c"import io, zipfile\noriginal = zipfile.ZipFile(io.BytesIO(seed))\noutput = io.BytesIO()\nwith zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:\n    for name in original.namelist():\n        archive.writestr(name, xml.encode() if name == 'word/document.xml' else original.read(name))\nresult = output.getvalue()", Some(&scope), Some(&scope)).unwrap();
            let bytes = scope
                .get_item("result")
                .unwrap()
                .unwrap()
                .extract::<Vec<u8>>()
                .unwrap();
            let document = rdocx::Document::from_bytes(&bytes).unwrap();
            let legacy = fx237_legacy_word_owner_locations(&document);
            let combined = word_owner_locations(&document, None);
            if content.contains("r:id='missing'") {
                assert!(legacy.is_err());
                assert!(combined.is_err());
            }
            match (legacy, combined) {
                (Ok(legacy), Ok(combined)) => assert_eq!(legacy, combined),
                (Err(legacy), Err(combined)) => {
                    assert_eq!(legacy.to_string(), combined.to_string())
                }
                (legacy, combined) => panic!(
                    "complete registry admission changed: legacy={legacy:?} combined={combined:?}"
                ),
            }
        }
    });
}

#[test]
fn fx237_revision_route_work_is_linear() {
    if let Ok(path) = std::env::var("FX237_PROFILE_INPUT") {
        let bytes = std::fs::read(path).unwrap();
        let started = std::time::Instant::now();
        let document = rdocx::Document::from_bytes(&bytes).unwrap();
        eprintln!("FX237 filled native_open={:?}", started.elapsed());
        for _ in 0..3 {
            let started = std::time::Instant::now();
            let items = document.story_item_snapshots().unwrap();
            let item_time = started.elapsed();
            let table = items
                .iter()
                .find(|item| item.location().item_kind() == rdocx::StoryItemKind::Table)
                .unwrap()
                .location();
            let started = std::time::Instant::now();
            document.story_table_cell(table, 0, 0).unwrap();
            let scalar_time = started.elapsed();
            let started = std::time::Instant::now();
            document.story_owner_snapshots().unwrap();
            let cell_time = started.elapsed();
            let started = std::time::Instant::now();
            document.story_range_paragraphs().unwrap();
            let paragraph_time = started.elapsed();
            let started = std::time::Instant::now();
            document.story_link_snapshots().unwrap();
            let link_time = started.elapsed();
            let started = std::time::Instant::now();
            let old = fx237_test_owners(word_owner_locations(&document, None).unwrap());
            let inventory_time = started.elapsed();
            let started = std::time::Instant::now();
            let mut candidate = document.clone_for_staging();
            candidate.try_set_cell_text(0, 0, 0, "edited").unwrap();
            let mutation_time = started.elapsed();
            let started = std::time::Instant::now();
            let new = word_owner_locations(&candidate, None).unwrap();
            let candidate_inventory_time = started.elapsed();
            FX237_ROUTE_WORK.set([0; 8]);
            let started = std::time::Instant::now();
            indexed_inventory_matches(&old, &new, &WordPublication::Conservative);
            eprintln!(
                "FX237 filled items={item_time:?} scalar={scalar_time:?} combined={cell_time:?} paragraphs={paragraph_time:?} links={link_time:?} registry={inventory_time:?} mutation={mutation_time:?} candidate_registry={candidate_inventory_time:?} index={:?} work={:?}",
                started.elapsed(),
                FX237_ROUTE_WORK.get()
            );
        }
    }
    for cells in [60, 120, 240] {
        let mut document = rdocx::Document::new();
        document.add_paragraph("same");
        document.add_table(cells / 6, 6);
        document.add_table(1, 1).cell(0, 0).unwrap().add_table(1, 1);
        let control = fx237_control_table();
        let body = document
            .stories()
            .unwrap()
            .into_iter()
            .find(|story| story.kind() == rdocx::StoryKind::Body)
            .unwrap();
        document
            .insert_content(
                &rdocx::ContentLocation::end(body),
                rdocx::ContentFragment::content_control(control).unwrap(),
            )
            .unwrap();
        assert_eq!(
            word_owner_locations(&document, None).unwrap(),
            fx237_legacy_word_owner_locations(&document).unwrap()
        );
        let old = fx237_test_owners(word_owner_locations(&document, None).unwrap());
        document.try_set_cell_text(0, 0, 0, "edited").unwrap();
        let new = word_owner_locations(&document, None).unwrap();
        FX237_ROUTE_WORK.set([0; 8]);
        let actual = indexed_inventory_matches(&old, &new, &WordPublication::Conservative);
        let work = FX237_ROUTE_WORK.get();
        assert_eq!(
            actual,
            fx237_legacy_matches(&old, &new, &WordPublication::Conservative)
        );
        eprintln!(
            "FX237 cells={cells} old={} new={} work={work:?}",
            old.len(),
            new.len()
        );
        assert!(
            work[4] <= 12 * new.len(),
            "actual predicate ancestry visits: {work:?}"
        );
        assert!(
            work[5] <= 2 * new.len(),
            "actual old-parent probes: {work:?}"
        );
        assert!(
            work[6] <= 3 * old.len(),
            "actual old-key insertions: {work:?}"
        );
        assert!(
            work[7] <= 1,
            "generated physical index keys must be unique: {work:?}"
        );
        assert_eq!(work[0], old.len());
        assert_eq!(work[1], new.len());
        assert!(
            work[2] <= 2 * new.len(),
            "bounded candidate checks: {work:?}"
        );
        assert!(work[3] <= 2, "physical generated parent matches: {work:?}");
    }
}

#[test]
fn fx237_indexed_routes_match_legacy_first_match() {
    let mut previous = rdocx::Document::new();
    previous.add_paragraph("before");
    let body = previous
        .stories()
        .unwrap()
        .into_iter()
        .find(|story| story.kind() == rdocx::StoryKind::Body)
        .unwrap();
    previous
        .insert_content(
            &rdocx::ContentLocation::end(body),
            rdocx::ContentFragment::content_control(fx237_control_table()).unwrap(),
        )
        .unwrap();
    previous.add_table(1, 1);
    let previous = fx237_test_owners(word_owner_locations(&previous, None).unwrap());
    let mut candidate = rdocx::Document::new();
    candidate.add_paragraph("before");
    candidate.add_paragraph("between");
    candidate.add_table(1, 1);
    let candidate = word_owner_locations(&candidate, None).unwrap();
    FX237_ROUTE_WORK.set([0; 8]);
    let indexed = indexed_inventory_matches(&previous, &candidate, &WordPublication::Conservative);
    let generated = FX237_ROUTE_WORK.get();
    assert_eq!(
        indexed,
        fx237_legacy_matches(&previous, &candidate, &WordPublication::Conservative)
    );
    assert!(
        generated[3] <= 2,
        "actual root class transitions remain bounded: {generated:?}"
    );
    assert_eq!(generated[7], 1, "each actual old index key remains unique");
    let mut document = rdocx::Document::new();
    document.add_paragraph("same");
    document.add_table(1, 2).cell(0, 0).unwrap().add_table(1, 1);
    let base = word_owner_locations(&document, None).unwrap();
    let mut old_addresses = Vec::new();
    // Mix slot presence independently across otherwise identical physical
    // addresses. Duplicate roots test first-match, not merely set equality.
    for address in &base {
        old_addresses.push(WordOwnerAddress {
            location: address.location.clone(),
            parent: address.parent,
            body_slot: address.body_slot,
            projection_ordinal: address.projection_ordinal,
            read_facts: address.read_facts.clone(),
        });
    }
    let paragraph = base
        .iter()
        .find(|address| {
            matches!(address.location, WordOwnerLocation::Paragraph(_)) && address.parent.is_none()
        })
        .unwrap();
    for slot in [None, Some(0), Some(1), None] {
        old_addresses.push(WordOwnerAddress {
            location: paragraph.location.clone(),
            parent: None,
            body_slot: slot,
            projection_ordinal: paragraph.projection_ordinal,
            read_facts: paragraph.read_facts.clone(),
        });
    }
    let location = match &paragraph.location {
        WordOwnerLocation::Paragraph(location) => location.clone(),
        _ => unreachable!(),
    };
    for index in [0, 1, 2, usize::MAX] {
        old_addresses.push(WordOwnerAddress {
            location: WordOwnerLocation::Hyperlink(location.clone(), index),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        });
    }
    old_addresses.push(WordOwnerAddress {
        location: WordOwnerLocation::Paragraph(rdocx::ContentLocation::new(
            location.story().clone(),
            rdocx::StoryItemKind::Paragraph,
            vec![0, 1, 2],
        )),
        parent: None,
        body_slot: Some(0),
        projection_ordinal: Some(0),
        read_facts: WordReadFacts::None,
    });
    // The earlier child belongs to the later duplicate parent. Choosing only
    // the first parent before querying descendants would return the wrong owner.
    let duplicate = fx237_test_owners(vec![
        WordOwnerAddress {
            location: WordOwnerLocation::Section(0),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        },
        WordOwnerAddress {
            location: WordOwnerLocation::Section(0),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        },
        WordOwnerAddress {
            location: paragraph.location.clone(),
            parent: Some(1),
            body_slot: None,
            projection_ordinal: Some(0),
            read_facts: WordReadFacts::None,
        },
        WordOwnerAddress {
            location: paragraph.location.clone(),
            parent: Some(0),
            body_slot: None,
            projection_ordinal: Some(0),
            read_facts: WordReadFacts::None,
        },
    ]);
    let duplicate_new = vec![
        WordOwnerAddress {
            location: WordOwnerLocation::Section(0),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        },
        WordOwnerAddress {
            location: paragraph.location.clone(),
            parent: Some(0),
            body_slot: None,
            projection_ordinal: Some(0),
            read_facts: WordReadFacts::None,
        },
    ];
    assert_eq!(
        indexed_inventory_matches(&duplicate, &duplicate_new, &WordPublication::Conservative),
        vec![Some(0), Some(2)]
    );
    assert_eq!(
        indexed_inventory_matches(&duplicate, &duplicate_new, &WordPublication::Conservative),
        fx237_legacy_matches(&duplicate, &duplicate_new, &WordPublication::Conservative)
    );
    let old = fx237_test_owners(old_addresses);
    for slot in [None, Some(0), Some(1), Some(2)] {
        let mut new = word_owner_locations(&document, None).unwrap();
        for address in &mut new {
            if matches!(address.location, WordOwnerLocation::Paragraph(_))
                && address.parent.is_none()
            {
                address.body_slot = slot;
            }
        }
        for index in [0, 1, 2, usize::MAX] {
            new.push(WordOwnerAddress {
                location: WordOwnerLocation::Hyperlink(location.clone(), index),
                parent: None,
                body_slot: None,
                projection_ordinal: None,
                read_facts: WordReadFacts::None,
            });
        }
        new.push(WordOwnerAddress {
            location: WordOwnerLocation::Paragraph(rdocx::ContentLocation::new(
                location.story().clone(),
                rdocx::StoryItemKind::Paragraph,
                vec![0, 1, 2],
            )),
            parent: None,
            body_slot: slot,
            projection_ordinal: Some(0),
            read_facts: WordReadFacts::None,
        });
        for publication in [
            WordPublication::Conservative,
            WordPublication::AppendBody,
            WordPublication::HyperlinkRemove(WordOwnerLocation::Hyperlink(location.clone(), 1)),
            WordPublication::CellText(location.clone(), 0, 0),
            WordPublication::TableGrid(location.clone()),
            WordPublication::ElementXml(paragraph.location.clone()),
            WordPublication::RunText(paragraph.location.clone()),
            WordPublication::HeaderFooterAppend(0),
            WordPublication::ParagraphAppend(location.clone()),
            WordPublication::ParagraphText(match &paragraph.location {
                WordOwnerLocation::Paragraph(location) => location.clone(),
                _ => unreachable!(),
            }),
        ] {
            assert_eq!(
                indexed_inventory_matches(&old, &new, &publication),
                fx237_legacy_matches(&old, &new, &publication)
            );
        }
    }
}

#[test]
fn fx241_form_certificates_require_unique_routes_and_unchanged_source() {
    let mut document = rdocx::Document::new();
    document.add_paragraph("equal");
    document.add_paragraph("equal");
    let mut candidate = document.clone_for_staging();
    candidate
        .paragraph_mut(1)
        .unwrap()
        .set_text("changed")
        .unwrap();
    let mut before = FormSources::default();
    let mut after = FormSources::default();
    let old = fx237_test_owners(word_owner_locations(&document, Some(&mut before)).unwrap());
    let new = word_owner_locations(&candidate, Some(&mut after)).unwrap();
    let matched = indexed_inventory_matches(&old, &new, &WordPublication::FormValues);
    let certificates = form_owner_certificates(&old, &new, &matched, &before, &after);
    let paragraph_indexes = new
        .iter()
        .enumerate()
        .filter_map(|(index, owner)| {
            matches!(owner.location, WordOwnerLocation::Paragraph(_)).then_some(index)
        })
        .collect::<Vec<_>>();
    let unchanged = paragraph_indexes[0];
    let changed = paragraph_indexes[1];
    assert!(certificates.contains(&(matched[unchanged].unwrap(), unchanged)));
    assert!(!certificates.contains(&(matched[changed].unwrap(), changed)));
    let route = form_owner_route(&new[unchanged].location).unwrap();
    before.xml.insert(route.clone(), None);
    assert!(
        !form_owner_certificates(&old, &new, &matched, &before, &after)
            .contains(&(matched[unchanged].unwrap(), unchanged))
    );
    // A duplicated actual owner route must not inherit the matcher's first index.
    let mut before = FormSources::default();
    let mut locations = word_owner_locations(&document, Some(&mut before)).unwrap();
    let source = &locations[unchanged];
    locations.push(WordOwnerAddress {
        location: source.location.clone(),
        parent: source.parent,
        body_slot: source.body_slot,
        projection_ordinal: source.projection_ordinal,
        read_facts: source.read_facts.clone(),
    });
    let duplicate = fx237_test_owners(locations);
    let matched = indexed_inventory_matches(&duplicate, &new, &WordPublication::FormValues);
    assert!(matched[unchanged].is_some());
    assert!(
        !form_owner_certificates(&duplicate, &new, &matched, &before, &after)
            .contains(&(matched[unchanged].unwrap(), unchanged))
    );
}

impl PyDocument {
    /// A body insertion that no save could write refuses before it changes
    /// anything, with the error the save would give.
    fn refuse_unsavable_body_edit(&self, py: Python<'_>) -> PyResult<()> {
        match self.inner.unsavable_body_edit() {
            Some(error) => Err(rdocx_to_pyerr(py, rdocx::Error::Other(error.to_owned()))),
            None => Ok(()),
        }
    }

    fn append_paragraph_candidate(
        slf: Py<Self>,
        py: Python<'_>,
        text: &str,
        style: Option<&str>,
        page_break: bool,
    ) -> PyResult<Py<PyParagraph>> {
        let mut document = slf.borrow_mut(py);
        document.refuse_unsavable_body_edit(py)?;
        let index = document.inner.paragraph_count();
        let style = style
            .map(|value| style_id_of_type(&document.inner, value, rdocx::StyleType::Paragraph))
            .transpose()?;
        // The same edit builds the paragraph read for the inventory and the
        // paragraph placed.
        let edit = |paragraph: &mut rdocx::Paragraph<'_>| {
            if let Some(style) = &style {
                paragraph.set_style(style);
            }
            if page_break {
                paragraph.add_run("").add_break(rdocx::BreakKind::Page);
            }
        };
        let mut appended = rdocx::Document::detached_text_paragraph(text);
        edit(&mut rdocx::Paragraph::detached(&mut appended));
        let slot = document.inner.content_count();
        let splice = document.inner.direct_body_paragraph_facts(&appended);
        if let Ok(Some(facts)) = splice
            && let Some(plan) = document.plan_body_splice(BodySplice {
                slot,
                removed: 0,
                inserted: Some(SplicedBodyItem::Paragraph(facts.0, facts.1)),
            })?
        {
            // Every fallible step precedes the in-place native append.
            let publication = crate::document::WordPublication::AppendBody;
            let prepared =
                document.prepare_body_splice(plan, "Document.add_paragraph", &publication)?;
            let path = document
                .revisions
                .capture_prepared(
                    &prepared.revision,
                    smallvec![PathSeg::Body(0), PathSeg::Para(index)],
                )
                .map_err(revision_to_pyerr)?;
            let result = Py::new(
                py,
                PyParagraph::new(slf.clone_ref(py), path)
                    .bind_prepared_splice(&document, &prepared)?,
            )?;
            edit(&mut document.inner.add_paragraph(text));
            debug_assert_eq!(prepared.plan.paragraph, index);
            document.commit_body_splice(prepared, &publication)?;
            return Ok(result);
        }
        let mut candidate = document.inner.clone_for_staging();
        edit(&mut candidate.add_paragraph(text));
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Document.add_paragraph",
            crate::document::WordPublication::AppendBody,
        )?;
        let path = document
            .revisions
            .capture_prepared(
                &prepared.revision,
                smallvec![PathSeg::Body(0), PathSeg::Para(index)],
            )
            .map_err(revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyParagraph::new(slf.clone_ref(py), path)
                .bind_prepared(py, &document, &prepared, &candidate, None)?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }
}

#[test]
fn issue389_local_edits_do_not_scale_with_unrelated_package() {
    Python::initialize();
    Python::attach(|py| -> PyResult<()> {
        for count in [50usize, 800] {
            let mut native = rdocx::Document::new();
            for index in 0..count {
                native.add_paragraph(&format!("p{index}"));
            }
            let document = Py::new(py, PyDocument::from_document(native, py)?)?;
            let path = {
                let borrowed = document.borrow(py);
                borrowed
                    .revisions
                    .capture(smallvec![PathSeg::Body(0), PathSeg::Para(count / 2)])
            };
            let paragraph = Py::new(
                py,
                crate::paragraph::PyParagraph::new(document.clone_ref(py), path)
                    .bind_current(py, None)?,
            )?;
            let xml = paragraph.borrow(py).xml(py)?;
            let changed = xml.replace(&format!(">p{}<", count / 2), ">changed<");
            assert_ne!(xml, changed);
            reset_fx260_word_visits();
            paragraph.borrow(py).replace_xml(py, &changed)?;
            assert!(
                fx260_word_visits() <= 8,
                "replace count={count}, visits={}",
                fx260_word_visits()
            );
            reset_fx260_word_visits();
            let removed = document.borrow_mut(py).remove_content(py, count - 1)?;
            assert!(removed);
            assert!(
                fx260_word_visits() <= 8,
                "remove count={count}, visits={}",
                fx260_word_visits()
            );
        }
        Ok(())
    })
    .unwrap();
}

#[test]
fn issue389_first_run_replacement_does_not_scan_unrelated_owners() {
    Python::initialize();
    Python::attach(|py| -> PyResult<()> {
        for count in [50usize, 800] {
            let mut native = rdocx::Document::new();
            for index in 0..count {
                native.add_paragraph(&format!("p{index}"));
            }
            native.insert_paragraph(count / 2, "");
            let document = Py::new(py, PyDocument::from_document(native, py)?)?;
            let path = document.borrow(py).revisions.capture(
                smallvec![PathSeg::Body(0), PathSeg::Para(count / 2)]);
            let paragraph = crate::paragraph::PyParagraph::new(document.clone_ref(py), path)
                .bind_current(py, None)?;
            let changed = r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>x</w:t></w:r></w:p>"#;
            reset_fx260_word_visits();
            paragraph.replace_xml(py, changed)?;
            assert!(fx260_word_visits() <= 8, "first run count={count}, visits={}", fx260_word_visits());
            reset_fx260_word_visits();
            paragraph.replace_xml(py, &changed.replace(">x<", ">y<"))?;
            assert!(fx260_word_visits() <= 8, "follow-up count={count}, visits={}", fx260_word_visits());
        }
        Ok(())
    }).unwrap();
}

#[test]
fn issue389_non_tail_preparation_refusal_preserves_native_and_inventory() {
    Python::initialize();
    Python::attach(|py| -> PyResult<()> {
        let mut native = rdocx::Document::new();
        native.add_paragraph("first");
        native.add_paragraph("second");
        let mut document = PyDocument::from_document(native, py)?;
        // Inject a retired inventory token to force the same preparation-error
        // boundary as a checked allocation or revision-budget refusal. This is
        // internal fault injection, not a normal authored document state.
        let scope = document.owners[0].owner.clone();
        let fault = document
            .revisions
            .prepare_local_scopes("test fault", &[scope], &[])
            .map_err(revision_to_pyerr)?;
        document
            .revisions
            .commit_scopes(fault)
            .map_err(revision_to_pyerr)?;
        let bytes = document.inner.to_bytes().unwrap();
        let owners = document.owners.len();
        let revision = document.revisions.current();
        assert!(document.remove_content(py, 0).is_err());
        assert_eq!(document.inner.to_bytes().unwrap(), bytes);
        assert_eq!(document.owners.len(), owners);
        assert_eq!(document.revisions.current(), revision);
        assert_eq!(document.indexes.by_generation.len(), owners);
        Ok(())
    })
    .unwrap();
}

#[test]
fn issue398_read_handles_visit_only_the_selected_owner_after_open() {
    Python::initialize();
    Python::attach(|py| -> PyResult<()> {
        let mut samples = Vec::new();
        let mut walks = Vec::new();
        for (paragraph_count, table_count) in [(50usize, 2usize), (400, 16)] {
            let mut native = rdocx::Document::new();
            for _ in 0..paragraph_count {
                native.add_paragraph("body text");
            }
            for _ in 0..table_count {
                let mut table = native.add_table(4, 3);
                table.cell(2, 1).unwrap().set_text("cell text");
            }
            let document = Py::new(py, PyDocument::from_document(native, py)?)?;
            let object = document.bind(py);
            let paragraphs = object.getattr("paragraphs")?;
            let tables = object.getattr("tables")?;
            let last = tables.get_item(table_count - 1)?;
            let row = last.getattr("rows")?.get_item(2)?;
            let cells = row.getattr("cells")?;
            let cell = cells.get_item(1)?;
            let cell_paragraphs = cell.getattr("paragraphs")?;
            let mut work = Vec::new();
            reset_fx260_word_visits();
            assert_eq!(
                paragraphs
                    .get_item(paragraph_count - 1)?
                    .getattr("text")?
                    .extract::<String>()?,
                "body text"
            );
            work.push(fx260_word_visits());
            reset_fx260_word_visits();
            cells.get_item(1)?;
            work.push(fx260_word_visits());
            reset_fx260_word_visits();
            assert_eq!(
                cell_paragraphs
                    .get_item(0)?
                    .getattr("text")?
                    .extract::<String>()?,
                "cell text"
            );
            work.push(fx260_word_visits());
            assert!(
                work.iter().all(|value| *value > 0 && *value <= 32),
                "{paragraph_count}: {work:?}"
            );
            samples.push(work);
            reset_fx260_word_visits();
            let mut read = 0;
            for paragraph in paragraphs.try_iter()? {
                paragraph?.getattr("text")?.extract::<String>()?;
                read += 1;
            }
            for table in tables.try_iter()? {
                for row in table?.getattr("rows")?.try_iter()? {
                    for cell in row?.getattr("cells")?.try_iter()? {
                        for paragraph in cell?.getattr("paragraphs")?.try_iter()? {
                            paragraph?.getattr("text")?.extract::<String>()?;
                            read += 1;
                        }
                    }
                }
            }
            assert_eq!(read, paragraph_count + table_count * 12);
            assert!(fx260_word_visits() <= read * 32);
            walks.push(fx260_word_visits());
            if paragraph_count == 400 {
                // Deliberately disable the projection, exercising the original
                // native whole-source fallback as a discriminating work control.
                document.borrow_mut(py).indexes.body_paragraphs = None;
                reset_fx260_word_visits();
                paragraphs.get_item(paragraph_count - 1)?;
                assert!(fx260_word_visits() >= paragraph_count);
            }
        }
        assert_eq!(samples[0], samples[1]);
        assert_eq!(walks[1], walks[0] * 8);
        Ok(())
    })
    .unwrap();
}

#[cfg(debug_assertions)]
#[test]
fn word_owner_pieces_read_only_the_blocks_an_edit_changes() {
    Python::initialize();
    Python::attach(|py| -> PyResult<()> {
        let mut samples = Vec::new();
        for count in [50usize, 200] {
            let mut native = rdocx::Document::new();
            for index in 0..count {
                native.add_paragraph(&format!("p{index}"));
                if index % 10 == 9 {
                    native.add_table(2, 2);
                }
            }
            let document = Py::new(py, PyDocument::from_document(native, py)?)?;
            let object = document.bind(py);
            let paragraphs = object.getattr("paragraphs")?;
            let table = object.getattr("tables")?.get_item(2)?;
            let cell = table.call_method1("cell", (0, 0))?;
            let mut sample = Vec::new();
            let reads = |edit: &dyn Fn() -> PyResult<()>| -> PyResult<usize> {
                let before = word_piece_work();
                edit()?;
                Ok(word_piece_work() - before)
            };
            sample.push(reads(&|| table.setattr("indent", 182_880))?);
            sample.push(reads(&|| cell.setattr("width", 914_400))?);
            sample.push(reads(&|| {
                cell.call_method1("add_paragraph", ("x",)).map(|_| ())
            })?);
            sample.push(reads(&|| {
                paragraphs.get_item(count / 2)?.setattr("text", "x")
            })?);
            sample.push(reads(&|| {
                paragraphs
                    .get_item(count / 3)?
                    .call_method1("add_run", ("x",))
                    .map(|_| ())
            })?);
            assert!(
                sample.iter().all(|reads| *reads == 1),
                "{count}: {sample:?}"
            );
            // Without its pieces, the next inventory reads every block again.
            document.borrow_mut(py).pieces = WordPieces::default();
            let all = reads(&|| paragraphs.get_item(1)?.setattr("text", "y"))?;
            assert!(all > count, "{count}: {all}");
            samples.push(sample);
        }
        assert_eq!(samples[0], samples[1]);
        Ok(())
    })
    .unwrap();
}

#[test]
fn issue398_route_hash_collisions_preserve_exact_first_match() {
    let addresses = (0..2)
        .map(|_| WordOwnerAddress {
            location: WordOwnerLocation::Section(0),
            parent: None,
            body_slot: None,
            projection_ordinal: None,
            read_facts: WordReadFacts::None,
        })
        .collect::<Vec<_>>();
    let mut revisions = RevisionCounter::new();
    let states = revisions
        .initialize_scopes(&[
            RevisionScope::Owner,
            RevisionScope::Text,
            RevisionScope::Owner,
            RevisionScope::Text,
        ])
        .unwrap();
    let owners = word_owners_with_states(addresses, &states);
    let mut indexes =
        WordOwnerIndexes::prepare(&owners, WordIndexProjections::of(&rdocx::Document::new()))
            .unwrap();
    assert_eq!(
        indexes.find(&owners, &WordOwnerLocation::Section(0)),
        Some(0)
    );
    let unrelated = WordOwnerLocation::Section(1);
    indexes.by_route.insert(unrelated.route_hash(), vec![0, 1]);
    assert_eq!(indexes.find(&owners, &unrelated), None);
}
