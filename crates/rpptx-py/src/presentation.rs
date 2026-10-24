use oxml_py_support::PathSeg;
use std::collections::{HashMap, HashSet};
use std::path::PathBuf;
use std::sync::Arc;

use oxml_py_support::RevisionCounter;
#[cfg(debug_assertions)]
use std::cell::Cell;

// Debug builds, which the binding tests use, count the owner inventory work.
#[cfg(debug_assertions)]
thread_local! {
    static FX260_PPT_VISITS: Cell<usize> = const { Cell::new(0) };
}

pub(crate) fn note_fx260_ppt(count: usize) {
    #[cfg(debug_assertions)]
    FX260_PPT_VISITS.with(|cell| cell.set(cell.get().saturating_add(count)));
    #[cfg(not(debug_assertions))]
    let _ = count;
}

#[cfg(debug_assertions)]
fn fx260_ppt_visits() -> usize {
    FX260_PPT_VISITS.with(|cell| cell.get())
}

#[cfg(test)]
fn reset_fx260_ppt_visits() {
    FX260_PPT_VISITS.with(|cell| cell.set(0));
}

/// Debug builds only: the owner inventory work counted since the previous
/// call, so that the Python tests can compare decks of N and 4N slides.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_inventory_work")]
pub(crate) fn inventory_work() -> usize {
    FX260_PPT_VISITS.with(|cell| cell.replace(0))
}

/// Debug builds only: the work staging and reopening packages did since the
/// previous call, as parts parsed, slide and notes parts serialized, media
/// hashed and packages written as a ZIP.
#[cfg(debug_assertions)]
#[pyfunction]
#[pyo3(name = "_commit_work")]
pub(crate) fn commit_work() -> [usize; 4] {
    rpptx::take_commit_work()
}
use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBytes, PyList, PyString, PyTuple};
use smallvec::smallvec;

use crate::layout::PyTextFrameLayout;
use crate::shape::length;
use crate::slide::{
    PySlideCollection, PySlideLayoutCollection, PySlideMaster, PySlideMasterCollection,
};
use crate::{replacement_count_to_pyerr, rpptx_to_pyerr, rpptx_value_to_pyerr};

/// The bundled 16:9 slide size, paired with the first dimension set on a deck
/// that has no `p:sldSz`.
const DEFAULT_SLIDE_SIZE: (i64, i64) = (12_192_000, 6_858_000);

#[pyclass(name = "CommentAuthor", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyCommentAuthor {
    pub id: String,
    pub name: String,
    pub initials: Option<String>,
    pub user_id: String,
    pub provider_id: String,
}

impl From<&rpptx::CommentAuthor> for PyCommentAuthor {
    fn from(author: &rpptx::CommentAuthor) -> Self {
        Self {
            id: author.id.clone(),
            name: author.name.clone(),
            initials: author.initials.clone(),
            user_id: author.user_id.clone(),
            provider_id: author.provider_id.clone(),
        }
    }
}

#[pyclass(name = "CommentReply", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyCommentReply {
    pub id: String,
    pub author_id: String,
    pub status: Option<String>,
    pub created: String,
    pub text: String,
}

impl From<&rpptx::CommentReply> for PyCommentReply {
    fn from(reply: &rpptx::CommentReply) -> Self {
        Self {
            id: reply.id.clone(),
            author_id: reply.author_id.clone(),
            status: reply.status.clone(),
            created: reply.created.clone(),
            text: reply.text(),
        }
    }
}

#[pyclass(name = "Comment", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyComment {
    id: String,
    author_id: String,
    status: Option<String>,
    created: String,
    text: String,
    replies: Vec<PyCommentReply>,
}

impl From<&rpptx::Comment> for PyComment {
    fn from(comment: &rpptx::Comment) -> Self {
        Self {
            id: comment.id.clone(),
            author_id: comment.author_id.clone(),
            status: comment.status.clone(),
            created: comment.created.clone(),
            text: comment.text(),
            replies: comment.replies().iter().map(Into::into).collect(),
        }
    }
}

#[pymethods]
impl PyComment {
    #[getter]
    fn id(&self) -> &str {
        &self.id
    }

    #[getter]
    fn author_id(&self) -> &str {
        &self.author_id
    }

    #[getter]
    fn status(&self) -> Option<&str> {
        self.status.as_deref()
    }

    #[getter]
    fn created(&self) -> &str {
        &self.created
    }

    #[getter]
    fn text(&self) -> &str {
        &self.text
    }

    #[getter]
    fn replies<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.replies.iter().cloned())
    }
}

/// One issue `Presentation::validate` reports, with the snake_case variant
/// name and the line `rpptx validate` prints for it.
#[pyclass(name = "ValidationIssue", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyValidationIssue {
    pub kind: &'static str,
    pub message: String,
}

impl From<&rpptx::ValidationIssue> for PyValidationIssue {
    fn from(issue: &rpptx::ValidationIssue) -> Self {
        use rpptx::ValidationIssue as Issue;
        let kind = match issue {
            Issue::DuplicateShapeId { .. } => "duplicate_shape_id",
            Issue::SlideIdOutOfRange { .. } => "slide_id_out_of_range",
            Issue::DuplicateSlideId { .. } => "duplicate_slide_id",
            Issue::MissingContentTypeOverride { .. } => "missing_content_type_override",
            Issue::DanglingRelationship { .. } => "dangling_relationship",
            Issue::UnreachableRelationshipTarget { .. } => "unreachable_relationship_target",
            Issue::EmptyTextBody { .. } => "empty_text_body",
            Issue::DuplicatePlaceholderIdx { .. } => "duplicate_placeholder_idx",
            Issue::OrphanMedia { .. } => "orphan_media",
            Issue::CustomShowReference { .. } => "custom_show_reference",
            Issue::MissingLayoutRel { .. } => "missing_layout_rel",
            Issue::MissingThemeRel { .. } => "missing_theme_rel",
            Issue::InvalidXmlChild { .. } => "invalid_xml_child",
        };
        Self {
            kind,
            message: format!("{issue:?}"),
        }
    }
}

/// One presentation section: a named, contiguous run of slides.
#[pyclass(name = "Section", frozen, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PySection {
    /// The braced GUID PowerPoint identifies the section by.
    pub id: String,
    pub name: String,
    /// The zero-based indices of the section's slides, in order.
    pub slide_indices: Vec<usize>,
    pub slide_ids: Vec<u32>,
}

#[pymethods]
impl PySection {
    #[getter]
    fn id(&self) -> &str {
        &self.id
    }
    #[getter]
    fn name(&self) -> &str {
        &self.name
    }
    #[getter]
    fn slide_indices<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.slide_indices)
    }
    #[getter]
    fn slide_ids<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.slide_ids)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!(
            "Section(id={}, name={}, slide_indices={:?})",
            pyo3::types::PyString::new(py, &self.id).repr()?,
            pyo3::types::PyString::new(py, &self.name).repr()?,
            self.slide_indices
        ))
    }
}

fn section_guid(name: &str, position: usize, attempt: u32) -> String {
    let hash = |seed: u64| {
        let mut hash = seed;
        for byte in name
            .bytes()
            .chain(position.to_le_bytes())
            .chain(attempt.to_le_bytes())
        {
            hash ^= u64::from(byte);
            hash = hash.wrapping_mul(0x0000_0100_0000_01B3);
        }
        hash
    };
    let (high, low) = (hash(0xCBF2_9CE4_8422_2325), hash(0x6C62_272E_07BB_0142));
    format!(
        "{{{:08X}-{:04X}-4{:03X}-8{:03X}-{:012X}}}",
        high >> 32,
        (high >> 16) & 0xFFFF,
        high & 0x0FFF,
        (low >> 48) & 0x0FFF,
        low & 0xFFFF_FFFF_FFFF
    )
}

const DEFAULT_JPEG_QUALITY: u8 = 90;

/// Reads a raster format, refusing an option the format would ignore.
fn raster_format(
    format: &str,
    quality: Option<i64>,
    transparent: bool,
) -> PyResult<rpptx::RasterFormat> {
    if quality.is_some() && !matches!(format, "jpg" | "jpeg") {
        return Err(PyValueError::new_err(format!(
            "quality applies to JPEG only, drop it or pass format=\"jpeg\" (format is {format:?})"
        )));
    }
    if transparent && format != "png" {
        return Err(PyValueError::new_err(format!(
            "transparent applies to PNG only, drop it or pass format=\"png\" (format is {format:?})"
        )));
    }
    match format {
        "png" => Ok(rpptx::RasterFormat::Png {
            transparent_background: transparent,
        }),
        "jpg" | "jpeg" => match quality.unwrap_or(i64::from(DEFAULT_JPEG_QUALITY)) {
            quality @ 1..=100 => Ok(rpptx::RasterFormat::Jpeg {
                quality: quality as u8,
            }),
            quality => Err(PyValueError::new_err(format!(
                "JPEG quality must be from 1 to 100, got {quality}"
            ))),
        },
        "tif" | "tiff" => Ok(rpptx::RasterFormat::Tiff),
        other => Err(PyValueError::new_err(format!(
            "unknown raster format {other:?}, expected png, jpeg, or tiff"
        ))),
    }
}

fn handout_layout(slides_per_page: u8) -> PyResult<rpptx::HandoutLayout> {
    Ok(match slides_per_page {
        1 => rpptx::HandoutLayout::One,
        2 => rpptx::HandoutLayout::Two,
        3 => rpptx::HandoutLayout::Three,
        4 => rpptx::HandoutLayout::Four,
        6 => rpptx::HandoutLayout::Six,
        9 => rpptx::HandoutLayout::Nine,
        _ => {
            return Err(PyValueError::new_err(format!(
                "slides_per_page must be 1, 2, 3, 4, 6, or 9, got {slides_per_page}"
            )));
        }
    })
}

/// One ODP conversion diagnostic as a `(path, message)` pair.
type OdpDiagnostics = Vec<(String, String)>;

fn odp_diagnostics(diagnostics: &[rpptx::OdpDiagnostic]) -> OdpDiagnostics {
    diagnostics
        .iter()
        .map(|diagnostic| (diagnostic.path.clone(), diagnostic.message.clone()))
        .collect()
}

#[pyclass(name = "AppProperties")]
pub struct PyAppProperties {
    presentation: Py<PyPresentation>,
}
impl PyAppProperties {
    fn read<T: Clone>(
        &self,
        py: Python<'_>,
        field: fn(&mut rpptx::AppProperties) -> &mut Option<T>,
    ) -> Option<T> {
        let mut properties = self
            .presentation
            .borrow(py)
            .inner
            .application_properties()?
            .clone();
        field(&mut properties).clone()
    }
    fn write<T: Clone + PartialEq>(
        &self,
        py: Python<'_>,
        field: fn(&mut rpptx::AppProperties) -> &mut Option<T>,
        value: Option<T>,
    ) -> PyResult<()> {
        let mut presentation = self.presentation.borrow_mut(py);
        let mut properties = presentation
            .inner
            .application_properties()
            .cloned()
            .unwrap_or_default();
        if *field(&mut properties) == value {
            return Ok(());
        }
        *field(&mut properties) = value;
        presentation
            .inner
            .set_application_properties(properties)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}
#[pymethods]
impl PyAppProperties {
    #[getter]
    fn template(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| &mut p.template)
    }
    #[setter]
    fn set_template(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |p| &mut p.template, value)
    }
    #[getter]
    fn manager(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| &mut p.manager)
    }
    #[setter]
    fn set_manager(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |p| &mut p.manager, value)
    }
    #[getter]
    fn company(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| &mut p.company)
    }
    #[setter]
    fn set_company(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |p| &mut p.company, value)
    }
    #[getter]
    fn pages(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.pages)
    }
    #[setter]
    fn set_pages(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.pages, value)
    }
    #[getter]
    fn words(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.words)
    }
    #[setter]
    fn set_words(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.words, value)
    }
    #[getter]
    fn characters(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.characters)
    }
    #[setter]
    fn set_characters(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.characters, value)
    }
    #[getter]
    fn presentation_format(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| &mut p.presentation_format)
    }
    #[setter]
    fn set_presentation_format(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |p| &mut p.presentation_format, value)
    }
    #[getter]
    fn lines(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.lines)
    }
    #[setter]
    fn set_lines(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.lines, value)
    }
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.paragraphs)
    }
    #[setter]
    fn set_paragraphs(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.paragraphs, value)
    }
    #[getter]
    fn slides(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.slides)
    }
    #[setter]
    fn set_slides(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.slides, value)
    }
    #[getter]
    fn notes(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.notes)
    }
    #[setter]
    fn set_notes(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.notes, value)
    }
    #[getter]
    fn total_time(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.total_time)
    }
    #[setter]
    fn set_total_time(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.total_time, value)
    }
    #[getter]
    fn hidden_slides(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.hidden_slides)
    }
    #[setter]
    fn set_hidden_slides(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.hidden_slides, value)
    }
    #[getter]
    fn multimedia_clips(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.multimedia_clips)
    }
    #[setter]
    fn set_multimedia_clips(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.multimedia_clips, value)
    }
    #[getter]
    fn scale_crop(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |p| &mut p.scale_crop)
    }
    #[setter]
    fn set_scale_crop(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.write(py, |p| &mut p.scale_crop, value)
    }
    #[getter]
    fn links_up_to_date(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |p| &mut p.links_up_to_date)
    }
    #[setter]
    fn set_links_up_to_date(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.write(py, |p| &mut p.links_up_to_date, value)
    }
    #[getter]
    fn characters_with_spaces(&self, py: Python<'_>) -> Option<i32> {
        self.read(py, |p| &mut p.characters_with_spaces)
    }
    #[setter]
    fn set_characters_with_spaces(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.write(py, |p| &mut p.characters_with_spaces, value)
    }
    #[getter]
    fn shared_document(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |p| &mut p.shared_document)
    }
    #[setter]
    fn set_shared_document(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.write(py, |p| &mut p.shared_document, value)
    }
    #[getter]
    fn hyperlinks_changed(&self, py: Python<'_>) -> Option<bool> {
        self.read(py, |p| &mut p.hyperlinks_changed)
    }
    #[setter]
    fn set_hyperlinks_changed(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.write(py, |p| &mut p.hyperlinks_changed, value)
    }
    #[getter]
    fn application(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| &mut p.application)
    }
    #[setter]
    fn set_application(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |p| &mut p.application, value)
    }
    #[getter]
    fn application_version(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |p| &mut p.application_version)
    }
    #[setter]
    fn set_application_version(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |p| &mut p.application_version, value)
    }
}

#[pyclass(name = "CustomFileTime", frozen, get_all)]
pub struct PyCustomFileTime {
    value: String,
}
#[pymethods]
impl PyCustomFileTime {
    #[new]
    fn new(value: String) -> Self {
        Self { value }
    }
}
#[pyclass(name = "CustomProperty", frozen, get_all)]
pub struct PyCustomProperty {
    name: Option<String>,
    pid: i32,
    fmtid: String,
    kind: &'static str,
    value: Py<PyAny>,
}
fn custom_property_result(
    py: Python<'_>,
    property: &rpptx::CustomProperty,
) -> PyResult<Py<PyCustomProperty>> {
    use rpptx::CustomPropertyValue as Value;
    let (kind, value) = match &property.value {
        Value::Lpstr(v) => ("lpstr", v.clone().into_pyobject(py)?.into_any().unbind()),
        Value::Lpwstr(v) => ("lpwstr", v.clone().into_pyobject(py)?.into_any().unbind()),
        Value::I4(v) => ("i4", v.into_pyobject(py)?.into_any().unbind()),
        Value::R8(v) => ("r8", v.into_pyobject(py)?.into_any().unbind()),
        Value::Bool(v) => ("bool", v.into_pyobject(py)?.to_owned().into_any().unbind()),
        Value::FileTime(v) => (
            "filetime",
            Py::new(py, PyCustomFileTime { value: v.clone() })?.into_any(),
        ),
        Value::Empty => ("empty", py.None()),
        Value::Raw(v) => ("raw", PyBytes::new(py, v).into_any().unbind()),
    };
    Py::new(
        py,
        PyCustomProperty {
            name: property.name.clone(),
            pid: property.pid,
            fmtid: property.fmtid.clone(),
            kind,
            value,
        },
    )
}
#[pyclass(name = "CustomProperties")]
pub struct PyCustomProperties {
    presentation: Py<PyPresentation>,
}
#[pymethods]
impl PyCustomProperties {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.presentation.borrow(py).inner.custom_properties().len()
    }
    fn get(&self, py: Python<'_>, name: &str) -> PyResult<Option<Py<PyCustomProperty>>> {
        self.presentation
            .borrow(py)
            .inner
            .custom_property(name)
            .map(|p| custom_property_result(py, p))
            .transpose()
    }
    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let p = self.presentation.borrow(py);
        let values = p
            .inner
            .custom_properties()
            .iter()
            .map(|p| custom_property_result(py, p))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, values)?
            .call_method0("__iter__")
            .map(Bound::unbind)
    }
    fn set(&self, py: Python<'_>, name: String, value: &Bound<'_, PyAny>) -> PyResult<()> {
        use rpptx::CustomPropertyValue as Value;
        let mut p = self.presentation.borrow_mut(py);
        let existing = p.inner.custom_property(&name);
        let converted = if value.is_none() {
            Value::Empty
        } else if value.is_instance_of::<pyo3::types::PyBool>() {
            Value::Bool(value.extract()?)
        } else if value.is_instance_of::<pyo3::types::PyInt>() {
            Value::I4(value.extract::<i32>()?)
        } else if value.is_instance_of::<pyo3::types::PyFloat>() {
            let number = value.extract::<f64>()?;
            if !number.is_finite() {
                return Err(PyValueError::new_err("custom float must be finite"));
            }
            Value::R8(number)
        } else if value.is_instance_of::<PyString>() {
            let text = value.extract::<String>()?;
            if existing.is_some_and(|p| matches!(p.value, Value::Lpstr(_))) {
                Value::Lpstr(text)
            } else {
                Value::Lpwstr(text)
            }
        } else if let Ok(filetime) = value.extract::<PyRef<'_, PyCustomFileTime>>() {
            Value::FileTime(filetime.value.clone())
        } else {
            return Err(PyTypeError::new_err(
                "custom value must be string, signed32 int, finite float, bool, CustomFileTime or None",
            ));
        };
        let (pid, fmtid) = if let Some(old) = existing {
            (old.pid, old.fmtid.clone())
        } else {
            let mut pid = 2i32;
            while p
                .inner
                .custom_properties()
                .iter()
                .any(|prop| prop.pid == pid)
            {
                pid = pid.checked_add(1).ok_or_else(|| {
                    pyo3::exceptions::PyOverflowError::new_err("custom property PID exhausted")
                })?;
            }
            (pid, "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}".to_owned())
        };
        p.inner
            .set_custom_property(rpptx::CustomProperty {
                fmtid,
                pid,
                name: Some(name),
                value: converted,
            })
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
    fn delete(&self, py: Python<'_>, name: &str) -> PyResult<()> {
        self.presentation
            .borrow_mut(py)
            .inner
            .remove_custom_property(name)
            .map(|_| ())
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

type CoreField = fn(&mut rpptx::CoreProperties) -> &mut Option<String>;

/// The longest text python-docx accepts for a core property.
const CORE_TEXT_LIMIT: usize = 255;

#[pyclass(name = "CoreProperties")]
pub struct PyCoreProperties {
    presentation: Py<PyPresentation>,
}

impl PyCoreProperties {
    fn read(&self, py: Python<'_>, field: CoreField) -> Option<String> {
        let presentation = self.presentation.borrow(py);
        let mut properties = presentation.inner.core_properties()?.clone();
        field(&mut properties).take()
    }

    fn write(&self, py: Python<'_>, field: CoreField, value: Option<String>) -> PyResult<()> {
        let mut presentation = self.presentation.borrow_mut(py);
        let mut properties = presentation
            .inner
            .core_properties()
            .cloned()
            .unwrap_or_default();
        let value = value.filter(|value| !value.is_empty());
        if *field(&mut properties) == value {
            return Ok(());
        }
        *field(&mut properties) = value;
        presentation
            .inner
            .set_core_properties(properties)
            .map_err(|error| rpptx_to_pyerr(py, error))
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
}

#[pymethods]
impl PyCoreProperties {
    #[getter]
    fn author(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |properties| &mut properties.creator)
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
    fn created(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |properties| &mut properties.created)
    }
    #[setter]
    fn set_created(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |properties| &mut properties.created, value)
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
    fn last_printed(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |properties| &mut properties.last_printed)
    }
    #[setter]
    fn set_last_printed(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |properties| &mut properties.last_printed, value)
    }
    #[getter]
    fn modified(&self, py: Python<'_>) -> Option<String> {
        self.read(py, |properties| &mut properties.modified)
    }
    #[setter]
    fn set_modified(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.write(py, |properties| &mut properties.modified, value)
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

/// Detached semantic text, owned image bytes, and loss diagnostics.
#[pyclass(name = "PresentationTextExport", frozen)]
pub struct PyPresentationTextExport {
    result: rpptx::PresentationTextExport,
}

#[pymethods]
impl PyPresentationTextExport {
    #[getter]
    fn text(&self) -> &str {
        &self.result.text
    }

    #[getter]
    fn assets<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.result
                .assets
                .iter()
                .map(|(name, bytes)| (name.as_str(), PyBytes::new(py, bytes))),
        )
    }

    #[getter]
    fn diagnostics<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, &self.result.diagnostics)
    }
}

#[pyclass(name = "Presentation")]
pub struct PyPresentation {
    pub(crate) inner: rpptx::Presentation,
    pub(crate) revisions: RevisionCounter,
    owners: OwnerInventory,
    pub(crate) gradient_stops: Vec<crate::dml::GradientStopTokens>,
    pub(crate) next_gradient_stop: u64,
    slides: oxml_py_support::RevisionState,
    layouts: oxml_py_support::RevisionState,
    masters: oxml_py_support::RevisionState,
}

impl PyPresentation {
    fn set_slide_size(
        &mut self,
        py: Python<'_>,
        width: Option<i64>,
        height: Option<i64>,
    ) -> PyResult<()> {
        let (current_width, current_height) = self
            .inner
            .slide_size()
            .map_or(DEFAULT_SLIDE_SIZE, |(width, height)| (width.0, height.0));
        for (name, selected, current) in [
            ("slide_width", width, current_width),
            ("slide_height", height, current_height),
        ] {
            if selected
                .is_some_and(|value| value != current && !(914_400..=51_206_400).contains(&value))
            {
                return Err(PyValueError::new_err(format!(
                    "{name} must be between 1 and 56 inches"
                )));
            }
        }
        self.inner
            .set_slide_size(
                rpptx::Emu(width.unwrap_or(current_width)),
                rpptx::Emu(height.unwrap_or(current_height)),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyPresentation {
    #[new]
    #[pyo3(signature = (path = None))]
    fn new(path: Option<PathBuf>, py: Python<'_>) -> PyResult<Self> {
        match path {
            Some(path) => rpptx::Presentation::open(path)
                .map_err(|error| rpptx_to_pyerr(py, error))
                .and_then(Self::from_presentation_owned),
            None => rpptx::Presentation::new()
                .map_err(|error| rpptx_to_pyerr(py, error))
                .and_then(Self::from_presentation_owned),
        }
    }

    #[staticmethod]
    fn from_bytes(bytes: &[u8], py: Python<'_>) -> PyResult<Self> {
        rpptx::Presentation::from_bytes(bytes)
            .map_err(|error| rpptx_to_pyerr(py, error))
            .and_then(Self::from_presentation_owned)
    }

    /// Sign current content with DER or one unencrypted PKCS#8/X.509 PEM block.
    fn sign(
        &mut self,
        private_key: &[u8],
        certificate: &[u8],
        py: Python<'_>,
    ) -> PyResult<Py<pyo3::types::PyDict>> {
        let mut candidate = self.inner.clone();
        let report = py
            .detach(|| candidate.sign(private_key, certificate))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let result = oxml_py_support::signature_report_to_py(py, &report)?;
        let survivors = self.unchanged_owners()?;
        let prepared = self.prepare_owners(
            &candidate,
            "Presentation.sign",
            &survivors,
            false,
            false,
            false,
        )?;
        self.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    /// Verify current content. Identity does not assert certificate-chain trust.
    fn verify_signatures(&self, py: Python<'_>) -> PyResult<Py<PyTuple>> {
        let reports = py
            .detach(|| self.inner.verify_signatures())
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let reports = reports
            .iter()
            .map(|report| oxml_py_support::signature_report_to_py(py, report))
            .collect::<PyResult<Vec<_>>>()?;
        Ok(PyTuple::new(py, reports)?.unbind())
    }

    fn save(&self, path: PathBuf, py: Python<'_>) -> PyResult<()> {
        if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| value.eq_ignore_ascii_case("pdf"))
        {
            return Err(PyValueError::new_err(
                "save writes PPTX, use to_pdf for PDF output",
            ));
        }
        self.inner
            .save(path)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(name = "to_bytes")]
    fn serialize<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        self.inner
            .to_bytes()
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[getter]
    fn custom_properties(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyCustomProperties>> {
        Py::new(py, PyCustomProperties { presentation: slf })
    }

    #[getter]
    fn app_properties(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyAppProperties>> {
        Py::new(py, PyAppProperties { presentation: slf })
    }

    #[getter]
    fn core_properties(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PyCoreProperties>> {
        Py::new(py, PyCoreProperties { presentation: slf })
    }

    /// The presentation sections in order, empty without sections.
    #[getter]
    fn sections<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let indices = self
            .inner
            .slides()
            .enumerate()
            .map(|(index, slide)| (slide.id(), index))
            .collect::<std::collections::HashMap<_, _>>();
        PyTuple::new(
            py,
            self.inner.sections().iter().map(|section| PySection {
                id: section.id.clone().unwrap_or_default(),
                name: section.name.clone().unwrap_or_default(),
                slide_ids: section.slide_ids.clone(),
                slide_indices: section
                    .slide_ids
                    .iter()
                    .filter_map(|id| indices.get(id).copied())
                    .collect(),
            }),
        )
    }

    /// Replaces the sections with `(name, slide_indices)` pairs. The
    /// sections must take every slide once, in slide order, as PowerPoint
    /// requires, and an empty list removes them. A section keeps the id of
    /// a current section with the same name.
    fn set_sections(
        &mut self,
        py: Python<'_>,
        sections: Vec<(String, Vec<usize>)>,
    ) -> PyResult<()> {
        let slide_ids = self
            .inner
            .slides()
            .map(|slide| slide.id())
            .collect::<Vec<_>>();
        let assigned = sections
            .iter()
            .flat_map(|(_, indices)| indices.iter().copied())
            .collect::<Vec<_>>();
        if !sections.is_empty() && assigned != (0..slide_ids.len()).collect::<Vec<_>>() {
            return Err(PyValueError::new_err(
                "sections must take every slide once, in slide order",
            ));
        }
        let mut current = self
            .inner
            .sections()
            .iter()
            .filter_map(|section| Some((section.name.clone()?, section.id.clone()?)))
            .collect::<Vec<_>>();
        let mut used = std::collections::HashSet::new();
        let mut native = Vec::with_capacity(sections.len());
        for (position, (name, indices)) in sections.into_iter().enumerate() {
            let reused = current
                .iter()
                .position(|(current_name, _)| *current_name == name)
                .map(|index| current.remove(index).1);
            let id = match reused {
                Some(id) if !used.contains(&id) => id,
                _ => (0..)
                    .map(|attempt| section_guid(&name, position, attempt))
                    .find(|id| !used.contains(id))
                    .expect("an unused section id exists"),
            };
            used.insert(id.clone());
            let ids = indices.iter().map(|index| slide_ids[*index]).collect();
            native.push(
                rpptx::Section::new(id, name, ids)
                    .map_err(|error| rpptx_value_to_pyerr(py, error.to_string()))?,
            );
        }
        self.inner
            .set_sections(native)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None))]
    fn to_pdf<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });

        py.detach(|| match &font_refs {
            None => self.inner.to_pdf_deterministic(),
            Some(fonts) => self.inner.to_pdf_deterministic_with_fonts(fonts),
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None))]
    fn to_pdf_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<(Bound<'py, PyBytes>, Bound<'py, PyTuple>)> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (payload, diagnostics) = py
            .detach(|| self.inner.to_pdf_with_diagnostics(font_refs))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((PyBytes::new(py, &payload), messages))
    }
    #[pyo3(signature = (slide_index, dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    #[allow(clippy::too_many_arguments)]
    fn render_slide_to_png<'py>(
        &self,
        py: Python<'py>,
        slide_index: usize,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Option<Bound<'py, PyBytes>>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });

        py.detach(|| match &font_refs {
            None => self.inner.slide_png_deterministic(slide_index, dpi),
            Some(fonts) => self
                .inner
                .slide_png_deterministic_with_fonts(slide_index, dpi, fonts),
        })
        .map(|bytes| bytes.map(|bytes| PyBytes::new(py, &bytes)))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (slide_index, dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    #[allow(clippy::too_many_arguments)]
    fn render_slide_to_png_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        slide_index: usize,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<(Option<Bound<'py, PyBytes>>, Bound<'py, PyTuple>)> {
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
                    .render_slide_to_png_with_diagnostics(slide_index, dpi, font_refs)
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((payload.map(|bytes| PyBytes::new(py, &bytes)), messages))
    }
    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    fn render_all_slides<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyList>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });

        let slides = py
            .detach(|| match &font_refs {
                None => self.inner.slide_pngs_deterministic(dpi),
                Some(fonts) => self.inner.slide_pngs_deterministic_with_fonts(dpi, fonts),
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, slides.iter().map(|slide| PyBytes::new(py, slide)))
    }

    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    fn render_all_slides_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<(Bound<'py, PyList>, Bound<'py, PyTuple>)> {
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
                    .render_all_slides_with_diagnostics(dpi, font_refs)
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((
            PyList::new(py, payload.iter().map(|bytes| PyBytes::new(py, bytes)))?,
            messages,
        ))
    }
    #[pyo3(signature = (*, width_factor = 1.0, fonts = None, font_dir = None, font_dirs = None))]
    fn text_layout<'py>(
        &self,
        py: Python<'py>,
        width_factor: f64,
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

        let frames = py
            .detach(|| match &font_refs {
                None => self.inner.text_layout_deterministic(width_factor),
                Some(fonts) => self
                    .inner
                    .text_layout_deterministic_with_fonts(width_factor, fonts),
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyTuple::new(py, frames.iter().map(PyTextFrameLayout::from))
    }

    /// Renders slides as PNG or JPEG images, one per slide, or as one
    /// multi-page TIFF. `slides` picks zero-based slides, all by default.
    /// `quality` (JPEG, 90 by default) and `transparent` (PNG) raise for
    /// another format rather than being ignored.
    #[pyo3(signature = (*, dpi = 150.0, format = "png", quality = None, transparent = false, slides = None, fonts = None, font_dir = None, font_dirs = None))]
    #[allow(clippy::too_many_arguments)] // Preserve the eight public keyword options at this PyO3 boundary.
    fn render_slides<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        format: &str,
        quality: Option<i64>,
        transparent: bool,
        slides: Option<Vec<usize>>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Py<PyAny>> {
        let format = raster_format(format, quality, transparent)?;
        let slides = slides.unwrap_or_else(|| (0..self.inner.len()).collect());
        let files = collect_caller_fonts(py, fonts, font_dir, font_dirs)?.unwrap_or_default();
        let refs = files
            .iter()
            .map(|font| (font.family.as_str(), font.data.as_slice()))
            .collect::<Vec<_>>();
        let (rendered, _) = py
            .detach(|| {
                self.inner.render_slides_with_diagnostics(
                    &slides,
                    rpptx::RasterOptions { dpi, format },
                    &refs,
                )
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        match rendered {
            rpptx::RasterOutput::SeparatePages(pages) => Ok(PyList::new(
                py,
                pages.iter().map(|page| PyBytes::new(py, page)),
            )?
            .into_any()
            .unbind()),
            rpptx::RasterOutput::MultiPageTiff(tiff) => {
                Ok(PyBytes::new(py, &tiff).into_any().unbind())
            }
        }
    }

    /// Read-only fit facts from the shared native solver, with caller fonts.
    #[pyo3(signature=(*,fonts=None,font_dir=None,font_dirs=None))]
    fn fit_report<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<crate::layout::PyFitReport> {
        let files = collect_caller_fonts(py, fonts, font_dir, font_dirs)?.unwrap_or_default();
        let refs = files
            .iter()
            .map(|f| (f.family.as_str(), f.data.as_slice()))
            .collect::<Vec<_>>();
        let result = py
            .detach(|| self.inner.fit_report(&refs))
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        Ok(crate::layout::PyFitReport::from(&result))
    }

    /// Explicit authoring transaction. Existing content handles keep their paths.
    #[pyo3(signature=(*,fonts=None,font_dir=None,font_dirs=None))]
    fn refresh_autofit<'py>(
        &mut self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<crate::layout::PyAutofitRefreshResult> {
        let files = collect_caller_fonts(py, fonts, font_dir, font_dirs)?.unwrap_or_default();
        let refs = files
            .iter()
            .map(|f| (f.family.as_str(), f.data.as_slice()))
            .collect::<Vec<_>>();
        let result = py
            .detach(|| self.inner.refresh_autofit(&refs))
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        Ok(crate::layout::PyAutofitRefreshResult::from(&result))
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None))]
    fn to_notes_pdf<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });

        py.detach(|| match &font_refs {
            None => self.inner.to_notes_pdf_deterministic(),
            Some(fonts) => self.inner.to_notes_pdf_deterministic_with_fonts(fonts),
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Archival PDF from one caller-font bundle with the selected exact profile.
    #[pyo3(signature = (profile = "pdfa-2b", *, fonts = None, font_dir = None, font_dirs = None))]
    fn to_pdfa<'py>(
        &self,
        py: Python<'py>,
        profile: &str,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let profile = match profile {
            "pdfa-2b" => rpptx::PdfConformance::PdfA2b,
            "pdfa-3b" => rpptx::PdfConformance::PdfA3b,
            _ => return Err(PyValueError::new_err("profile must be pdfa-2b or pdfa-3b")),
        };
        let files = collect_caller_fonts(py, fonts, font_dir, font_dirs)?.unwrap_or_default();
        let refs = files
            .iter()
            .map(|font| (font.family.as_str(), font.data.as_slice()))
            .collect::<Vec<_>>();
        py.detach(|| self.inner.to_pdfa_deterministic_with_fonts(profile, &refs))
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Audience handout PDF. Legal absent masters use read-only neutral defaults.
    #[pyo3(signature = (slides_per_page = 6, *, fonts = None, font_dir = None, font_dirs = None))]
    fn to_handout_pdf<'py>(
        &self,
        py: Python<'py>,
        slides_per_page: u8,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let arrangement = handout_layout(slides_per_page)?;
        let files = collect_caller_fonts(py, fonts, font_dir, font_dirs)?.unwrap_or_default();
        let refs = files
            .iter()
            .map(|font| (font.family.as_str(), font.data.as_slice()))
            .collect::<Vec<_>>();
        py.detach(|| {
            self.inner
                .to_handout_pdf_deterministic_with_fonts(arrangement, &refs)
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (slides_per_page = 6, dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    fn render_all_handouts<'py>(
        &self,
        py: Python<'py>,
        slides_per_page: u8,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyList>> {
        let arrangement = handout_layout(slides_per_page)?;
        let files = collect_caller_fonts(py, fonts, font_dir, font_dirs)?.unwrap_or_default();
        let refs = files
            .iter()
            .map(|font| (font.family.as_str(), font.data.as_slice()))
            .collect::<Vec<_>>();
        let pages = py
            .detach(|| {
                self.inner
                    .handout_page_pngs_deterministic_with_fonts(arrangement, dpi, &refs)
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, pages.iter().map(|bytes| PyBytes::new(py, bytes)))
    }

    /// Returns semantic html text. Use export_html for diagnostics and assets.
    #[pyo3(signature = (*, include_notes = false, skip_hidden = false, inline_images = true))]
    fn to_html(
        &self,
        py: Python<'_>,
        include_notes: bool,
        skip_hidden: bool,
        inline_images: bool,
    ) -> PyResult<String> {
        if !inline_images {
            return Err(PyValueError::new_err(
                "inline_images=False requires export_html or save_html to retain image assets",
            ));
        }
        let options = rpptx::PresentationTextExportOptions {
            include_notes,
            skip_hidden,
            inline_images,
        };
        py.detach(|| self.inner.to_html_with_options(&options))
            .map(|result| result.text)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Returns detached semantic text, image assets, and diagnostics.
    #[pyo3(signature = (*, include_notes = false, skip_hidden = false, inline_images = true))]
    fn export_html(
        &self,
        py: Python<'_>,
        include_notes: bool,
        skip_hidden: bool,
        inline_images: bool,
    ) -> PyResult<PyPresentationTextExport> {
        let options = rpptx::PresentationTextExportOptions {
            include_notes,
            skip_hidden,
            inline_images,
        };
        py.detach(|| self.inner.to_html_with_options(&options))
            .map(|result| PyPresentationTextExport { result })
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Atomically saves semantic text and adjacent assets, returning diagnostics.
    #[pyo3(signature = (path, *, include_notes = false, skip_hidden = false, inline_images = false))]
    fn save_html<'py>(
        &self,
        py: Python<'py>,
        path: PathBuf,
        include_notes: bool,
        skip_hidden: bool,
        inline_images: bool,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let options = rpptx::PresentationTextExportOptions {
            include_notes,
            skip_hidden,
            inline_images,
        };
        let diagnostics = py
            .detach(|| self.inner.save_html(&path, &options))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyTuple::new(py, diagnostics)
    }

    /// Returns semantic markdown text. Use export_markdown for diagnostics and assets.
    #[pyo3(signature = (*, include_notes = false, skip_hidden = false, inline_images = true))]
    fn to_markdown(
        &self,
        py: Python<'_>,
        include_notes: bool,
        skip_hidden: bool,
        inline_images: bool,
    ) -> PyResult<String> {
        if !inline_images {
            return Err(PyValueError::new_err(
                "inline_images=False requires export_markdown or save_markdown to retain image assets",
            ));
        }
        let options = rpptx::PresentationTextExportOptions {
            include_notes,
            skip_hidden,
            inline_images,
        };
        py.detach(|| self.inner.to_markdown_with_options(&options))
            .map(|result| result.text)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Returns detached semantic text, image assets, and diagnostics.
    #[pyo3(signature = (*, include_notes = false, skip_hidden = false, inline_images = true))]
    fn export_markdown(
        &self,
        py: Python<'_>,
        include_notes: bool,
        skip_hidden: bool,
        inline_images: bool,
    ) -> PyResult<PyPresentationTextExport> {
        let options = rpptx::PresentationTextExportOptions {
            include_notes,
            skip_hidden,
            inline_images,
        };
        py.detach(|| self.inner.to_markdown_with_options(&options))
            .map(|result| PyPresentationTextExport { result })
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Atomically saves semantic text and adjacent assets, returning diagnostics.
    #[pyo3(signature = (path, *, include_notes = false, skip_hidden = false, inline_images = false))]
    fn save_markdown<'py>(
        &self,
        py: Python<'py>,
        path: PathBuf,
        include_notes: bool,
        skip_hidden: bool,
        inline_images: bool,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let options = rpptx::PresentationTextExportOptions {
            include_notes,
            skip_hidden,
            inline_images,
        };
        let diagnostics = py
            .detach(|| self.inner.save_markdown(&path, &options))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyTuple::new(py, diagnostics)
    }

    /// Converts an OpenDocument presentation, from a path or bytes, and
    /// returns the deck with the `(path, message)` diagnostics of what the
    /// conversion changed or left out.
    #[staticmethod]
    fn from_odp(py: Python<'_>, file: &Bound<'_, PyAny>) -> PyResult<(Self, OdpDiagnostics)> {
        let result = if let Ok(bytes) = file.extract::<Vec<u8>>()
            && !file.is_instance_of::<pyo3::types::PyString>()
        {
            py.detach(|| rpptx::Presentation::from_odp_bytes(&bytes))
        } else {
            let path = file.extract::<PathBuf>()?;
            py.detach(|| rpptx::Presentation::open_odp(path))
        }
        .map_err(|error| rpptx_to_pyerr(py, error))?;
        let diagnostics = odp_diagnostics(&result.diagnostics);
        Ok((
            Self::from_presentation_owned(result.presentation)?,
            diagnostics,
        ))
    }

    /// Serializes the deck as an OpenDocument presentation and returns the
    /// bytes with the `(path, message)` diagnostics of what was left out.
    fn to_odp<'py>(&self, py: Python<'py>) -> PyResult<(Bound<'py, PyBytes>, OdpDiagnostics)> {
        let result = py
            .detach(|| self.inner.to_odp_bytes())
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        Ok((
            PyBytes::new(py, &result.bytes),
            odp_diagnostics(&result.diagnostics),
        ))
    }

    /// Saves the deck as an OpenDocument presentation and returns the
    /// `(path, message)` diagnostics of what was left out.
    fn save_odp(&self, py: Python<'_>, path: PathBuf) -> PyResult<OdpDiagnostics> {
        py.detach(|| self.inner.save_odp(path))
            .map(|diagnostics| odp_diagnostics(&diagnostics))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    fn render_all_notes<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<Bound<'py, PyList>> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });

        let notes = py
            .detach(|| match &font_refs {
                None => self.inner.notes_page_pngs_deterministic(dpi),
                Some(fonts) => self
                    .inner
                    .notes_page_pngs_deterministic_with_fonts(dpi, fonts),
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, notes.iter().map(|page| PyBytes::new(py, page)))
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None, font_dirs = None))]
    fn to_notes_pdf_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<(Bound<'py, PyBytes>, Bound<'py, PyTuple>)> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (payload, diagnostics) = py
            .detach(|| self.inner.to_notes_pdf_with_diagnostics(font_refs))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((PyBytes::new(py, &payload), messages))
    }
    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None, font_dirs = None))]
    fn render_all_notes_with_diagnostics<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
        font_dirs: Option<Bound<'py, PyAny>>,
    ) -> PyResult<(Bound<'py, PyList>, Bound<'py, PyTuple>)> {
        let caller_fonts = collect_caller_fonts(py, fonts, font_dir, font_dirs)?;
        let font_refs = caller_fonts.as_ref().map(|files| {
            files
                .iter()
                .map(|font| (font.family.as_str(), font.data.as_slice()))
                .collect::<Vec<_>>()
        });
        let font_refs = font_refs.as_deref().unwrap_or(&[]);
        let (payload, diagnostics) = py
            .detach(|| self.inner.render_all_notes_with_diagnostics(dpi, font_refs))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let messages = PyTuple::new(py, diagnostics.iter().map(|d| d.message.as_str()))?;
        Ok((
            PyList::new(py, payload.iter().map(|bytes| PyBytes::new(py, bytes)))?,
            messages,
        ))
    }

    /// Replaces literal text in slides and speaker notes and returns the count.
    ///
    /// With `expect`, the replacement runs on a clone, so a count that
    /// differs raises and leaves the presentation and its revision as they
    /// were. Both paths prepare owner retirement before publication. The revision
    /// advances once only when something was replaced.
    #[pyo3(signature = (placeholder, replacement, *, expect = None))]
    fn try_replace_text(
        &mut self,
        py: Python<'_>,
        placeholder: &str,
        replacement: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        self.replace_text_candidate(
            py,
            placeholder,
            replacement,
            expect,
            "Presentation.try_replace_text",
        )
    }

    /// Applies an ordered counted batch and publishes one owner revision.
    fn replace_all<'py>(
        &mut self,
        py: Python<'py>,
        pairs: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let mut owned = Vec::new();
        for (index, pair) in pairs.try_iter()?.enumerate() {
            let pair = pair?;
            let parsed = match pair.extract::<(String, String)>() {
                Ok((old, new)) => (old, new, None),
                Err(_) => pair.extract::<(String, String, Option<usize>)>().map_err(|_| {
                    PyTypeError::new_err(format!("pair {index} must be an (old, new) or (old, new, expected) tuple, where expected is a nonnegative int or None"))
                })?,
            };
            owned.push(parsed);
        }
        let pairs = owned
            .iter()
            .map(|(old, new, expected)| (old.as_str(), new.as_str(), *expected))
            .collect::<Vec<_>>();
        let mut candidate = self.inner.clone();
        let counts = py
            .detach(|| candidate.try_replace_all_expected(&pairs))
            .map_err(|error| rpptx_to_pyerr(py, error))?
            .map_err(|mismatch| {
                let error = replacement_count_to_pyerr(
                    py,
                    &mismatch.placeholder,
                    mismatch.expected,
                    mismatch.found,
                );
                if let Err(attribute_error) = error
                    .value(py)
                    .setattr("index", mismatch.index)
                    .and_then(|()| error.value(py).setattr("placeholder", mismatch.placeholder))
                {
                    return attribute_error;
                }
                error
            })?;
        let result = PyTuple::new(py, counts)?;
        let changed = py
            .detach(|| -> rpptx::Result<bool> { Ok(!self.inner.saves_as(&candidate)?) })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        if changed {
            let survivors = self.replaced_text_owners(&candidate)?;
            let prepared = self.prepare_owners(
                &candidate,
                "Presentation.replace_all",
                &survivors,
                false,
                false,
                false,
            )?;
            self.publish_owners(candidate, prepared)?;
        }
        Ok(result)
    }

    /// Alias for the checked literal replacement operation.
    #[pyo3(signature = (placeholder, replacement, *, expect = None))]
    fn replace_text(
        &mut self,
        py: Python<'_>,
        placeholder: &str,
        replacement: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        self.replace_text_candidate(
            py,
            placeholder,
            replacement,
            expect,
            "Presentation.replace_text",
        )
    }

    /// Returns every package and PresentationML invariant violation.
    fn validate<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let issues = py.detach(|| self.inner.validate());
        PyTuple::new(py, issues.iter().map(PyValidationIssue::from))
    }

    #[getter]
    fn slide_width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        length(py, self.inner.slide_size().map(|(width, _)| width))
    }

    #[setter(slide_width)]
    fn set_slide_width(&mut self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.set_slide_size(py, Some(value), None)
    }

    #[getter]
    fn slide_height(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        length(py, self.inner.slide_size().map(|(_, height)| height))
    }

    #[setter(slide_height)]
    fn set_slide_height(&mut self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.set_slide_size(py, None, Some(value))
    }

    #[getter]
    fn comment_authors<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.inner
                .comment_authors()
                .iter()
                .map(PyCommentAuthor::from),
        )
    }

    /// Applies selected branding schemes, optionally importing a fresh master graph.
    #[pyo3(signature = (source, *, source_master=None, import_masters=false, layout_map=None))]
    fn apply_theme(
        &mut self,
        py: Python<'_>,
        source: &Bound<'_, PyAny>,
        source_master: Option<usize>,
        import_masters: bool,
        layout_map: Option<Vec<(usize, usize)>>,
    ) -> PyResult<()> {
        let bytes = if source.is_instance_of::<PyBytes>() {
            source.extract::<Vec<u8>>()?
        } else {
            let path = source.extract::<PathBuf>()?;
            std::fs::read(path)
                .map_err(|error| pyo3::exceptions::PyOSError::new_err(error.to_string()))?
        };
        let package = rpptx::OpcPackage::from_reader(std::io::Cursor::new(bytes))
            .map_err(|error| rpptx_to_pyerr(py, error.into()))?;
        let mut candidate = self.inner.clone();
        candidate
            .apply_theme(
                &package,
                source_master,
                import_masters,
                layout_map.as_deref().unwrap_or(&[]),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let proofs = self.unchanged_owners()?;
        let prepared = self.prepare_owners(
            &candidate,
            "Presentation.apply_theme",
            &proofs,
            false,
            import_masters,
            import_masters,
        )?;
        self.publish_owners(candidate, prepared)
    }

    #[pyo3(signature = (*, id, name, user_id, provider_id, initials = None))]
    fn add_comment_author(
        &mut self,
        id: String,
        name: String,
        user_id: String,
        provider_id: String,
        initials: Option<&str>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let author = rpptx::CommentAuthor::new(id, name, initials, user_id, provider_id)
            .map_err(|error| rpptx_value_to_pyerr(py, error.to_string()))?;
        let mut candidate = self.inner.clone();
        candidate
            .add_comment_author(author)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = self.unchanged_owners()?;
        let prepared = self.prepare_owners(
            &candidate,
            "Presentation.add_comment_author",
            &survivors,
            false,
            false,
            false,
        )?;
        self.publish_owners(candidate, prepared)?;
        Ok(())
    }

    #[pyo3(signature = (slide_number=true, footer=None, date=None, hide_on_title=true, date_format="datetime1", *, date_cache=None))]
    fn set_header_footer(
        mut slf: PyRefMut<'_, Self>,
        slide_number: bool,
        footer: Option<String>,
        date: Option<String>,
        hide_on_title: bool,
        date_format: &str,
        date_cache: Option<String>,
    ) -> PyResult<()> {
        let py = slf.py();
        let presentation = &mut *slf;
        let settings = crate::slide::header_footer_settings(
            slide_number,
            footer,
            date,
            date_format,
            date_cache,
        );
        let mut candidate = presentation.inner.clone();
        if candidate
            .set_header_footer(&settings, hide_on_title)
            .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            let proofs = presentation.footer_owners(&candidate)?;
            let prepared = presentation.prepare_owners(
                &candidate,
                "Presentation.set_header_footer",
                &proofs,
                false,
                false,
                false,
            )?;
            presentation.publish_owners(candidate, prepared)?;
        }
        Ok(())
    }

    #[getter]
    fn slide_number_start(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        self.inner
            .slide_number_start()
            .map_err(|e| rpptx_to_pyerr(py, e))
    }

    #[setter]
    fn set_slide_number_start(&mut self, py: Python<'_>, value: Option<u32>) -> PyResult<()> {
        if self
            .inner
            .slide_number_start()
            .map_err(|e| rpptx_to_pyerr(py, e))?
            == value
        {
            return Ok(());
        }
        let mut candidate = self.inner.clone();
        candidate
            .set_slide_number_start(value)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        let proofs = self.unchanged_owners()?;
        let prepared = self.prepare_owners(
            &candidate,
            "Presentation.slide_number_start",
            &proofs,
            false,
            false,
            false,
        )?;
        self.publish_owners(candidate, prepared)
    }

    #[getter]
    fn slide_masters(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideMasterCollection>> {
        let path = slf
            .borrow(py)
            .revisions
            .capture_scopes(smallvec![], &[slf.borrow(py).masters.clone()])
            .map_err(oxml_py_support::revision_to_pyerr)?;
        Py::new(py, PySlideMasterCollection::new(slf, path))
    }

    #[getter]
    fn slide_master(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideMaster>> {
        let path = slf
            .borrow(py)
            .revisions
            .capture_scopes(smallvec![], &[slf.borrow(py).masters.clone()])
            .map_err(oxml_py_support::revision_to_pyerr)?;
        PySlideMasterCollection::new(slf, path).item(py, 0)
    }

    #[getter]
    fn slide_layouts(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideLayoutCollection>> {
        let path = slf
            .borrow(py)
            .revisions
            .capture_scopes(smallvec![], &[slf.borrow(py).layouts.clone()])
            .map_err(oxml_py_support::revision_to_pyerr)?;
        Py::new(py, PySlideLayoutCollection::new(slf, path))
    }

    #[getter]
    fn slides(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideCollection>> {
        let path = slf.borrow(py).capture_path(smallvec![], None)?;
        Py::new(py, PySlideCollection::new(slf, path))
    }

    #[getter]
    fn embedded_fonts(&self, py: Python<'_>) -> PyResult<Py<PyTuple>> {
        let fonts = self
            .inner
            .embedded_fonts()
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        let mut snapshots = Vec::new();
        snapshots.try_reserve_exact(fonts.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("font listing allocation failed")
        })?;
        let mut bodies = std::collections::BTreeMap::new();
        for font in fonts {
            snapshots.push(PyEmbeddedTypeface::snapshot(py, font, &mut bodies)?);
        }
        Ok(PyTuple::new(py, snapshots)?.unbind())
    }
    #[pyo3(signature = (name, font, kind="regular", *, license))]
    fn embed_font(
        &mut self,
        py: Python<'_>,
        name: &str,
        font: &Bound<'_, PyAny>,
        kind: &str,
        license: PyRef<'_, PyFontEmbeddingLicense>,
    ) -> PyResult<()> {
        let kind = checked_font_kind(kind)?;
        let license = checked_font_license(&license)?;
        let data = if let Ok(value) = font.extract::<PyRef<'_, PyEmbeddedFont>>() {
            let original = value.license.borrow(py);
            if original.authorized != license.authorized
                || original.identity != license.identity
                || checked_font_kind(value.kind)? != kind
            {
                return Err(PyValueError::new_err(
                    "authored font kind and license must match the embedding request",
                ));
            }
            checked_font_input(value.data.bind(py).as_any())?
        } else {
            checked_font_input(font)?
        };
        let mut candidate = self.inner.clone();
        candidate
            .embed_font(name, rpptx::EmbeddedFont::new(kind, data, license))
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        self.publish_font_candidate(py, candidate, "Presentation.embed_font")
    }
    #[pyo3(signature = (name, kind="regular"))]
    fn remove_embedded_font(
        &mut self,
        py: Python<'_>,
        name: &str,
        kind: &str,
    ) -> PyResult<Option<Py<PyEmbeddedFontFace>>> {
        let kind = checked_font_kind(kind)?;
        let mut candidate = self.inner.clone();
        let removed = candidate
            .remove_embedded_font(name, kind)
            .map_err(|e| rpptx_to_pyerr(py, e))?
            .map(|font| {
                PyEmbeddedFontFace::snapshot(py, font, &mut std::collections::BTreeMap::new())
            })
            .transpose()?;
        self.publish_font_candidate(py, candidate, "Presentation.remove_embedded_font")?;
        Ok(removed)
    }
    fn remove_font(
        &mut self,
        py: Python<'_>,
        name: &str,
    ) -> PyResult<Option<Py<PyEmbeddedTypeface>>> {
        let mut candidate = self.inner.clone();
        let removed = candidate
            .remove_font(name)
            .map_err(|e| rpptx_to_pyerr(py, e))?
            .map(|font| {
                PyEmbeddedTypeface::snapshot(py, font, &mut std::collections::BTreeMap::new())
            })
            .transpose()?;
        self.publish_font_candidate(py, candidate, "Presentation.remove_font")?;
        Ok(removed)
    }

    #[getter]
    fn embed_true_type_fonts(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.inner
            .embed_true_type_fonts()
            .map_err(|e| rpptx_to_pyerr(py, e))
    }
    #[setter]
    fn set_embed_true_type_fonts(&mut self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let mut candidate = self.inner.clone();
        candidate
            .set_embed_true_type_fonts(value)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        self.publish_font_candidate(py, candidate, "Presentation.set_embed_true_type_fonts")
    }
    #[getter]
    fn save_subset_fonts(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.inner
            .save_subset_fonts()
            .map_err(|e| rpptx_to_pyerr(py, e))
    }
    #[setter]
    fn set_save_subset_fonts(&mut self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let mut candidate = self.inner.clone();
        candidate
            .set_save_subset_fonts(value)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        self.publish_font_candidate(py, candidate, "Presentation.set_save_subset_fonts")
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum PptElementKind {
    Shape,
    Table,
    Cell,
    TextBody,
    Paragraph,
    Run,
    Slide,
    Layout,
    Master,
}

/// A checked physical XML owner. Snapshots are detached strings, while writes
/// stage the native package and all owner tokens before one publication.
#[pyclass(name = "Element", frozen)]
pub struct PyPptElement {
    presentation: Py<PyPresentation>,
    path: oxml_py_support::ContentPath,
    kind: PptElementKind,
}

impl PyPptElement {
    pub(crate) fn new_checked(
        py: Python<'_>,
        presentation: Py<PyPresentation>,
        path: oxml_py_support::ContentPath,
        kind: PptElementKind,
    ) -> PyResult<Py<Self>> {
        let element = Self {
            presentation,
            path,
            kind,
        };
        let owner = element.presentation.borrow(py);
        let path = element.checked_path(py, &owner)?;
        element.read_xml(py, &owner.inner, &path)?;
        drop(owner);
        Py::new(py, element)
    }

    fn checked_path(
        &self,
        py: Python<'_>,
        owner: &PyPresentation,
    ) -> PyResult<oxml_py_support::ContentPath> {
        crate::validate_path(py, owner, &self.path, "Element", ".element")?;
        let generation = self.path.owner_generation().ok_or_else(|| {
            pyo3::exceptions::PyNotImplementedError::new_err(
                "Element requires a physical XML owner",
            )
        })?;
        let actual = owner
            .owners
            .owner(generation)
            .ok_or_else(|| PyValueError::new_err("Element owner is absent"))?;
        let valid = matches!(
            (self.kind, &actual.address.physical),
            (
                PptElementKind::Shape | PptElementKind::Table,
                PhysicalOwner::Shape { .. }
            ) | (PptElementKind::Cell, PhysicalOwner::Cell { .. })
                | (PptElementKind::TextBody, PhysicalOwner::TextBody { .. })
                | (PptElementKind::Paragraph, PhysicalOwner::Paragraph { .. })
                | (PptElementKind::Run, PhysicalOwner::Run { .. })
                | (PptElementKind::Slide, PhysicalOwner::Slide(_))
                | (PptElementKind::Layout, PhysicalOwner::Layout(_))
                | (PptElementKind::Master, PhysicalOwner::Master(_))
        );
        if !valid {
            return Err(pyo3::exceptions::PyNotImplementedError::new_err(
                "Element is not backed by the requested physical XML owner",
            ));
        }
        owner
            .current_path(&self.path)
            .ok_or_else(|| PyValueError::new_err("Element owner is absent"))
    }

    fn read_xml(
        &self,
        py: Python<'_>,
        inner: &rpptx::Presentation,
        path: &oxml_py_support::ContentPath,
    ) -> PyResult<Vec<u8>> {
        use oxml_py_support::PathSeg;
        let part = crate::shape::shape_owner(path)?;
        let shapes = path
            .segs
            .iter()
            .filter_map(|seg| {
                if let PathSeg::Shape(i) = seg {
                    Some(*i)
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        let row = path.segs.iter().find_map(|seg| {
            if let PathSeg::Row(i) = seg {
                Some(*i)
            } else {
                None
            }
        });
        let col = path.segs.iter().find_map(|seg| {
            if let PathSeg::Cell(i) = seg {
                Some(*i)
            } else {
                None
            }
        });
        let cell = row.zip(col);
        let paragraph = path
            .segs
            .iter()
            .find_map(|seg| {
                if let PathSeg::Para(i) = seg {
                    Some(*i)
                } else {
                    None
                }
            })
            .unwrap_or(0);
        let run = path
            .segs
            .iter()
            .find_map(|seg| {
                if let PathSeg::Run(i) = seg {
                    Some(*i)
                } else {
                    None
                }
            })
            .unwrap_or(0);
        let result = match self.kind {
            PptElementKind::Shape => inner.shape_xml_at(part, &shapes),
            PptElementKind::Table => inner.table_xml_at(part, &shapes),
            PptElementKind::Cell => {
                let (r, c) = cell.ok_or_else(|| PyValueError::new_err("cell address is absent"))?;
                inner.cell_xml_at(part, &shapes, r, c)
            }
            PptElementKind::TextBody => inner.text_body_xml_at(part, &shapes, cell),
            PptElementKind::Paragraph => inner.paragraph_xml_at(part, &shapes, cell, paragraph),
            PptElementKind::Run => inner.run_xml_at(part, &shapes, cell, paragraph, run),
            PptElementKind::Slide => {
                let rpptx::ShapeOwner::Slide(i) = part else {
                    return Err(PyValueError::new_err("slide address is absent"));
                };
                inner.slide_xml(i)
            }
            PptElementKind::Layout => {
                let rpptx::ShapeOwner::Layout(i) = part else {
                    return Err(PyValueError::new_err("layout address is absent"));
                };
                inner.layout_xml(i)
            }
            PptElementKind::Master => {
                let rpptx::ShapeOwner::Master(i) = part else {
                    return Err(PyValueError::new_err("master address is absent"));
                };
                inner.master_xml(i)
            }
        };
        result.map_err(|error| rpptx_to_pyerr(py, error))
    }
    fn write_xml(
        &self,
        py: Python<'_>,
        inner: &mut rpptx::Presentation,
        path: &oxml_py_support::ContentPath,
        xml: &str,
    ) -> PyResult<()> {
        let part = crate::shape::shape_owner(path)?;
        let shapes = path
            .segs
            .iter()
            .filter_map(|seg| {
                if let PathSeg::Shape(i) = seg {
                    Some(*i)
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        let row = path.segs.iter().find_map(|seg| {
            if let PathSeg::Row(i) = seg {
                Some(*i)
            } else {
                None
            }
        });
        let col = path.segs.iter().find_map(|seg| {
            if let PathSeg::Cell(i) = seg {
                Some(*i)
            } else {
                None
            }
        });
        let cell = row.zip(col);
        let paragraph = path
            .segs
            .iter()
            .find_map(|seg| {
                if let PathSeg::Para(i) = seg {
                    Some(*i)
                } else {
                    None
                }
            })
            .unwrap_or(0);
        let run = path
            .segs
            .iter()
            .find_map(|seg| {
                if let PathSeg::Run(i) = seg {
                    Some(*i)
                } else {
                    None
                }
            })
            .unwrap_or(0);
        let result = match self.kind {
            PptElementKind::Shape => inner.replace_shape_xml_at(part, &shapes, xml),
            PptElementKind::Table => inner.replace_table_xml_at(part, &shapes, xml),
            PptElementKind::Cell => {
                let (r, c) = cell.ok_or_else(|| PyValueError::new_err("cell address is absent"))?;
                inner.replace_cell_xml_at(part, &shapes, r, c, xml)
            }
            PptElementKind::TextBody => inner.replace_text_body_xml_at(part, &shapes, cell, xml),
            PptElementKind::Paragraph => {
                inner.replace_paragraph_xml_at(part, &shapes, cell, paragraph, xml)
            }
            PptElementKind::Run => {
                inner.replace_run_xml_at(part, &shapes, cell, paragraph, run, xml)
            }
            PptElementKind::Slide => {
                let rpptx::ShapeOwner::Slide(i) = part else {
                    return Err(PyValueError::new_err("slide address is absent"));
                };
                inner.replace_slide_xml(i, xml)
            }
            PptElementKind::Layout => {
                let rpptx::ShapeOwner::Layout(i) = part else {
                    return Err(PyValueError::new_err("layout address is absent"));
                };
                inner.replace_layout_xml(i, xml)
            }
            PptElementKind::Master => {
                let rpptx::ShapeOwner::Master(i) = part else {
                    return Err(PyValueError::new_err("master address is absent"));
                };
                inner.replace_master_xml(i, xml)
            }
        };
        result.map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyPptElement {
    #[getter]
    fn xml(&self, py: Python<'_>) -> PyResult<String> {
        let owner = self.presentation.borrow(py);
        let path = self.checked_path(py, &owner)?;
        String::from_utf8(self.read_xml(py, &owner.inner, &path)?)
            .map_err(|_| PyValueError::new_err("physical XML snapshot is not UTF-8"))
    }

    fn replace_xml(&self, py: Python<'_>, xml: &str) -> PyResult<()> {
        let mut owner = self.presentation.borrow_mut(py);
        let path = self.checked_path(py, &owner)?;
        let before_xml = self.read_xml(py, &owner.inner, &path)?;
        // Native checked replacement has the same exact-snapshot no-op rule.
        if before_xml == xml.as_bytes() {
            return Ok(());
        }
        if self.kind == PptElementKind::Shape
            && prove_local_shape_xml(&owner.inner, &path, &before_xml, xml.as_bytes())?
        {
            let prepared = owner.prepare_shape_xml(&path)?;
            self.write_xml(py, &mut owner.inner, &path, xml)?;
            // The pre-native scalar witness establishes an actual change.
            // No fallible read, inventory discovery or fallback follows it.
            return owner.commit_shape_xml(prepared);
        }
        let before_structure = (self.kind == PptElementKind::Shape)
            .then(|| shape_structure(&owner.inner, &path))
            .transpose()?;
        let mut candidate = owner.inner.clone();
        self.write_xml(py, &mut candidate, &path, xml)?;
        let after_xml = self.read_xml(py, &candidate, &path)?;
        if before_xml == after_xml {
            return Ok(());
        }
        if let Some(before) = before_structure
            && before.has_table != shape_structure(&candidate, &path)?.has_table
        {
            return Err(pyo3::exceptions::PyNotImplementedError::new_err(
                "Shape.element.replace_xml does not support changing table payload presence",
            ));
        }
        let generation = self
            .path
            .owner_generation()
            .ok_or_else(|| PyValueError::new_err("Element owner is absent"))?;
        let survivors = owner.element_xml_survivors(&path, self.kind, generation)?;
        let prepared = owner.prepare_owners(
            &candidate,
            "Element.replace_xml",
            &survivors,
            false,
            false,
            false,
        )?;
        owner.publish_owners(candidate, prepared)
    }
}

// Physical addresses are current lookup locations, not lifetime identities.
#[derive(Clone, Copy)]
pub(crate) enum SlideOwnerEdit {
    Insert(usize),
    InsertMany { at: usize, count: usize },
    Remove(usize),
    Move { from: usize, to: usize },
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
enum PartIdentity {
    Slide(u32),
    Layout(String),
    Master(String),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
enum PhysicalOwner {
    Master(String),
    Slide(u32),
    Shape {
        slide: PartIdentity,
        id: Option<u32>,
        kind: rpptx::ShapeKind,
    },
    Cell {
        slide: PartIdentity,
        table: Option<u32>,
        row: usize,
        column: usize,
    },
    Layout(String),
    TextBody {
        slide: PartIdentity,
        shape: Option<u32>,
        cell: Option<(usize, usize)>,
    },
    Paragraph {
        slide: PartIdentity,
        shape: Option<u32>,
        cell: Option<(usize, usize)>,
        index: usize,
    },
    Run {
        slide: PartIdentity,
        shape: Option<u32>,
        cell: Option<(usize, usize)>,
        paragraph: usize,
        index: usize,
    },
}

#[derive(Clone, Debug, PartialEq)]
struct OwnerAddress {
    physical: PhysicalOwner,
    path: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
    layout_index: Option<usize>,
    table: bool,
}

#[derive(Clone, Debug)]
struct LiveOwner {
    address: OwnerAddress,
    owner: oxml_py_support::RevisionState,
    text: oxml_py_support::RevisionState,
    xml_children: oxml_py_support::RevisionState,
    grid: Option<oxml_py_support::RevisionState>,
}

/// One owner's explicit survival proof.
/// A matching native ID never constructs this proof automatically.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct OwnerSurvival {
    pub(crate) generation: u64,
    pub(crate) destination: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
    pub(crate) layout_index: Option<usize>,
    pub(crate) text_changed: bool,
    pub(crate) xml_children_changed: bool,
    pub(crate) grid_changed: bool,
}

/// An operation's explicit survival proofs. Omitted old generations are deleted.
pub(crate) struct OwnerProofs {
    /// How slide or layout ordinals move. Every owner of a part absent from
    /// `parts` survives unchanged at its path under this edit.
    edit: PartEdit,
    /// Parts whose surviving owners are exactly these proofs, in inventory order.
    parts: Vec<(OwnerPart, Vec<OwnerSurvival>)>,
}

#[derive(Clone, Copy)]
enum PartEdit {
    Unchanged,
    Slides(SlideOwnerEdit),
    RemoveLayout(usize),
}

impl PartEdit {
    /// Where a live slide ends up, or `None` when the edit removes it.
    fn slide_index(self, index: usize) -> Option<usize> {
        let Self::Slides(edit) = self else {
            return Some(index);
        };
        Some(match edit {
            SlideOwnerEdit::Insert(at) if index >= at => index + 1,
            SlideOwnerEdit::InsertMany { at, count } if index >= at => index
                .checked_add(count)
                .expect("native slide count was checked"),
            SlideOwnerEdit::Remove(at) if index == at => return None,
            SlideOwnerEdit::Remove(at) if index > at => index - 1,
            SlideOwnerEdit::Move { from, to } if index == from => to,
            SlideOwnerEdit::Move { from, to } if from < to && index > from && index <= to => {
                index - 1
            }
            SlideOwnerEdit::Move { from, to } if to < from && index >= to && index < from => {
                index + 1
            }
            _ => index,
        })
    }

    /// The proof of an owner this edit keeps unchanged, or `None` when it
    /// removes the owner's slide or layout.
    fn proof(self, owner: &LiveOwner) -> Option<OwnerSurvival> {
        let mut proof = OwnerSurvival {
            generation: owner.owner.generation(),
            destination: owner.address.path.clone(),
            layout_index: owner.address.layout_index,
            text_changed: false,
            xml_children_changed: false,
            grid_changed: false,
        };
        match self {
            Self::Unchanged => {}
            Self::Slides(_) => {
                if let Some(PathSeg::Slide(index)) = proof.destination.first_mut() {
                    *index = self.slide_index(*index)?;
                }
            }
            Self::RemoveLayout(removed) => {
                if proof.layout_index == Some(removed)
                    || matches!(proof.destination.first(), Some(PathSeg::Layout(index)) if *index == removed)
                {
                    return None;
                }
                if let Some(index) = &mut proof.layout_index
                    && *index > removed
                {
                    *index -= 1;
                }
                if let Some(PathSeg::Layout(index)) = proof.destination.first_mut()
                    && *index > removed
                {
                    *index -= 1;
                }
            }
        }
        Some(proof)
    }
}

impl OwnerProofs {
    fn edit(edit: PartEdit) -> Self {
        Self {
            edit,
            parts: Vec::new(),
        }
    }

    /// The proofs of one live part's owners, in inventory order.
    fn for_part(
        &self,
        part: OwnerPart,
        owners: &PartOwners,
    ) -> PyResult<std::borrow::Cow<'_, [OwnerSurvival]>> {
        if let Some((_, proofs)) = self.parts.iter().find(|(listed, _)| *listed == part) {
            return Ok(std::borrow::Cow::Borrowed(proofs));
        }
        let mut proofs = Vec::new();
        proofs.try_reserve_exact(owners.owners.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("owner proofs allocation failed")
        })?;
        proofs.extend(
            owners
                .owners
                .iter()
                .filter_map(|owner| self.edit.proof(owner)),
        );
        Ok(std::borrow::Cow::Owned(proofs))
    }
}

type OwnerPathKey = (
    smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
    Option<usize>,
    bool,
);

struct PreparedShapeXml {
    revision: oxml_py_support::PreparedRevision,
    part: OwnerPart,
    updates: Vec<(usize, LiveOwner)>,
}

/// One part of the owner inventory: a slide, or every layout and master.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum OwnerPart {
    Slide(usize),
    Masters,
}

impl OwnerPart {
    /// The part that holds the owners under a path. Layout owners have an
    /// empty path and belong with the layout and master shapes.
    fn of(path: &[PathSeg]) -> Self {
        match path.first() {
            Some(PathSeg::Slide(index)) => Self::Slide(*index),
            _ => Self::Masters,
        }
    }
}

/// The owners of one part in inventory order, with their local lookups.
#[derive(Clone, Debug, Default)]
struct PartOwners {
    owners: Vec<LiveOwner>,
    by_path: HashMap<OwnerPathKey, usize>,
    shape_ids: HashMap<(PartIdentity, u32), usize>,
    /// The structural proof's refusal when the whole part survives unchanged:
    /// its first shape without an ID or with one the part repeats.
    refusal: Option<&'static str>,
}

impl PartOwners {
    fn new(owners: Vec<LiveOwner>) -> PyResult<Self> {
        let mut part = Self {
            owners,
            ..Self::default()
        };
        let fail = |_| {
            pyo3::exceptions::PyMemoryError::new_err("PowerPoint owner index allocation failed")
        };
        part.by_path.try_reserve(part.owners.len()).map_err(fail)?;
        part.shape_ids
            .try_reserve(part.owners.len())
            .map_err(fail)?;
        for index in 0..part.owners.len() {
            part.index(index);
        }
        part.refusal = part.structural_refusal();
        Ok(part)
    }

    fn reserve(&mut self, additional: usize) -> PyResult<()> {
        let fail = |_| {
            pyo3::exceptions::PyMemoryError::new_err("PowerPoint owner index allocation failed")
        };
        self.owners.try_reserve(additional).map_err(fail)?;
        self.by_path.try_reserve(additional).map_err(fail)?;
        self.shape_ids.try_reserve(additional).map_err(fail)?;
        Ok(())
    }

    fn index(&mut self, index: usize) {
        note_fx260_ppt(1);
        let owner = &self.owners[index];
        let body = matches!(owner.address.physical, PhysicalOwner::TextBody { .. });
        self.by_path
            .entry((owner.address.path.clone(), owner.address.layout_index, body))
            .or_insert(index);
        if let PhysicalOwner::Shape {
            slide,
            id: Some(id),
            ..
        } = &owner.address.physical
        {
            *self.shape_ids.entry((slide.clone(), *id)).or_default() += 1;
        }
    }

    fn structural_refusal(&self) -> Option<&'static str> {
        self.owners
            .iter()
            .find_map(|owner| match &owner.address.physical {
                PhysicalOwner::Shape {
                    kind: rpptx::ShapeKind::AlternateContent,
                    ..
                } => None,
                PhysicalOwner::Shape { id: None, .. } => {
                    Some("structural owner proof requires a shape id")
                }
                PhysicalOwner::Shape {
                    slide,
                    id: Some(id),
                    ..
                } => (self.shape_ids.get(&(slide.clone(), *id)) != Some(&1))
                    .then_some("structural owner proof has an ambiguous shape id"),
                _ => None,
            })
    }

    /// The same owners after their slide moved to another ordinal.
    fn moved_to(&self, slide: usize) -> PyResult<Self> {
        let mut owners = Vec::new();
        owners.try_reserve_exact(self.owners.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("moved owner inventory allocation failed")
        })?;
        owners.extend(self.owners.iter().cloned());
        for owner in &mut owners {
            if let Some(PathSeg::Slide(index)) = owner.address.path.first_mut() {
                *index = slide;
            }
        }
        Self::new(owners)
    }
}

/// The owner inventory in parts, one per slide in slide order, then the
/// layouts and masters. Untouched parts are shared with a prepared edit.
#[derive(Clone, Debug, Default)]
struct OwnerParts {
    slides: Vec<Arc<PartOwners>>,
    masters: Arc<PartOwners>,
}

impl OwnerParts {
    fn part(&self, part: OwnerPart) -> Option<&PartOwners> {
        match part {
            OwnerPart::Slide(index) => self.slides.get(index).map(|owners| &**owners),
            OwnerPart::Masters => Some(&self.masters),
        }
    }

    fn part_mut(&mut self, part: OwnerPart) -> Option<&mut PartOwners> {
        match part {
            OwnerPart::Slide(index) => self.slides.get_mut(index).map(Arc::make_mut),
            OwnerPart::Masters => Some(Arc::make_mut(&mut self.masters)),
        }
    }

    /// Every part in inventory order.
    fn iter(&self) -> impl Iterator<Item = (OwnerPart, &PartOwners)> {
        self.slides
            .iter()
            .enumerate()
            .map(|(index, owners)| (OwnerPart::Slide(index), &**owners))
            .chain(std::iter::once((OwnerPart::Masters, &*self.masters)))
    }

    fn len(&self) -> usize {
        self.iter().map(|(_, part)| part.owners.len()).sum()
    }

    fn owner_at(
        &self,
        path: &[PathSeg],
        layout_index: Option<usize>,
        body: bool,
    ) -> Option<&LiveOwner> {
        note_fx260_ppt(1);
        let part = self.part(OwnerPart::of(path))?;
        let key = (smallvec::SmallVec::from_slice(path), layout_index, body);
        part.owners.get(*part.by_path.get(&key)?)
    }
}

/// The live owner inventory with its generation and shape ID indexes.
#[derive(Debug, Default)]
struct OwnerInventory {
    parts: OwnerParts,
    by_generation: HashMap<u64, (OwnerPart, usize)>,
    shape_ids: HashMap<(PartIdentity, u32), usize>,
}

impl OwnerInventory {
    fn new(parts: OwnerParts) -> PyResult<Self> {
        let mut inventory = Self {
            parts,
            ..Self::default()
        };
        let fail = |_| {
            pyo3::exceptions::PyMemoryError::new_err("PowerPoint owner index allocation failed")
        };
        let parts = &inventory.parts;
        inventory
            .by_generation
            .try_reserve(parts.len())
            .map_err(fail)?;
        for (part, owners) in parts.iter() {
            for (index, owner) in owners.owners.iter().enumerate() {
                inventory
                    .by_generation
                    .insert(owner.owner.generation(), (part, index));
            }
            for (key, count) in &owners.shape_ids {
                *inventory.shape_ids.entry(key.clone()).or_default() += count;
            }
        }
        Ok(inventory)
    }

    fn owner(&self, generation: u64) -> Option<&LiveOwner> {
        let (part, index) = self.by_generation.get(&generation)?;
        self.parts.part(*part)?.owners.get(*index)
    }

    /// Reserves every index entry a prepared change adds, before its commit.
    fn reserve(&mut self, change: &InventoryChange) -> PyResult<()> {
        let fail = |_| {
            pyo3::exceptions::PyMemoryError::new_err("PowerPoint owner index allocation failed")
        };
        let mut owners = 0;
        let mut ids = 0;
        for part in change.rebuilt.iter().chain(&change.moved) {
            let part = change.parts.part(*part).expect("prepared inventory part");
            owners += part.owners.len();
            ids += part.shape_ids.len();
        }
        self.by_generation.try_reserve(owners).map_err(fail)?;
        self.shape_ids.try_reserve(ids).map_err(fail)?;
        Ok(())
    }

    /// Installs a prepared change. Nothing here can fail.
    fn apply(&mut self, change: InventoryChange) {
        for part in &change.dropped {
            let owners = self.parts.part(*part).expect("dropped live part");
            for owner in &owners.owners {
                self.by_generation.remove(&owner.owner.generation());
            }
            for (key, count) in &owners.shape_ids {
                match self.shape_ids.get_mut(key) {
                    Some(total) if *total > *count => *total -= count,
                    _ => {
                        self.shape_ids.remove(key);
                    }
                }
            }
        }
        self.parts = change.parts;
        for part in change.rebuilt.iter().chain(&change.moved) {
            let owners = self.parts.part(*part).expect("prepared inventory part");
            for (index, owner) in owners.owners.iter().enumerate() {
                self.by_generation
                    .insert(owner.owner.generation(), (*part, index));
            }
        }
        for part in &change.rebuilt {
            let owners = self.parts.part(*part).expect("rebuilt inventory part");
            for (key, count) in &owners.shape_ids {
                *self.shape_ids.entry(key.clone()).or_default() += count;
            }
        }
        #[cfg(debug_assertions)]
        self.assert_indexes();
    }

    /// Debug builds check that the indexes kept by deltas match a rebuild.
    #[cfg(debug_assertions)]
    fn assert_indexes(&self) {
        let rebuilt = Self::new(self.parts.clone()).expect("debug index rebuild");
        assert_eq!(
            self.by_generation, rebuilt.by_generation,
            "owner generation index"
        );
        assert_eq!(self.shape_ids, rebuilt.shape_ids, "shape ID index");
    }
}

/// A prepared inventory and the live parts it replaces.
struct InventoryChange {
    parts: OwnerParts,
    /// Live parts whose owners leave the indexes.
    dropped: Vec<OwnerPart>,
    /// Prepared parts walked again, whose owners enter the indexes.
    rebuilt: Vec<OwnerPart>,
    /// Prepared parts carried from another slide ordinal.
    moved: Vec<OwnerPart>,
}

/// The live parts a prepared edit keeps.
struct CarriedParts {
    /// For each candidate slide, the live slide whose part it keeps.
    slides: Vec<Option<usize>>,
    /// Whether the layout and master part is kept.
    masters: bool,
}

impl CarriedParts {
    fn none(candidate: &rpptx::Presentation) -> Self {
        Self {
            slides: vec![None; candidate.len()],
            masters: false,
        }
    }

    /// For each live slide, whether a candidate part keeps it.
    fn kept_live_slides(&self, live_slides: usize) -> Vec<bool> {
        let mut kept = vec![false; live_slides];
        for from in self.slides.iter().flatten() {
            kept[*from] = true;
        }
        kept
    }

    /// Whether a candidate part is kept from the live inventory.
    fn keeps(&self, part: OwnerPart) -> bool {
        match part {
            OwnerPart::Slide(at) => self.slides.get(at).is_some_and(Option::is_some),
            OwnerPart::Masters => self.masters,
        }
    }
}

enum OwnerChanges {
    Replace(InventoryChange),
    Append(Vec<LiveOwner>),
}

pub(crate) struct PreparedOwners {
    revision: oxml_py_support::PreparedRevision,
    owners: OwnerChanges,
    slides: oxml_py_support::RevisionState,
    layouts: oxml_py_support::RevisionState,
    masters: oxml_py_support::RevisionState,
    /// The generations this edit retires or replaces, in preparation order.
    #[cfg(debug_assertions)]
    retired: Vec<u64>,
}

/// Runs a debug-only cross-check without counting its work.
#[cfg(debug_assertions)]
fn without_work_count<T>(run: impl FnOnce() -> T) -> T {
    let visits = fx260_ppt_visits();
    let result = run();
    FX260_PPT_VISITS.with(|cell| cell.set(visits));
    result
}

/// Debug builds check that carrying parts over prepares what walking every
/// part prepares: the same error, or the same retired generations, owners,
/// generations and inventory scopes.
#[cfg(debug_assertions)]
fn assert_same_preparation(carried: &PyResult<PreparedOwners>, walked: &PyResult<PreparedOwners>) {
    let summary = |prepared: &PreparedOwners| {
        let OwnerChanges::Replace(change) = &prepared.owners else {
            unreachable!("a full preparation replaces the inventory");
        };
        let owners = change
            .parts
            .iter()
            .flat_map(|(_, part)| &part.owners)
            .map(|owner| {
                (
                    owner.address.clone(),
                    owner.owner.generation(),
                    owner.text.generation(),
                    owner.xml_children.generation(),
                    owner
                        .grid
                        .as_ref()
                        .map(oxml_py_support::RevisionState::generation),
                )
            })
            .collect::<Vec<_>>();
        let scopes = [&prepared.slides, &prepared.layouts, &prepared.masters]
            .map(oxml_py_support::RevisionState::generation);
        (prepared.retired.clone(), owners, scopes)
    };
    match (carried, walked) {
        (Ok(carried), Ok(walked)) => {
            assert!(
                summary(carried) == summary(walked),
                "carried owner inventory differs from a full walk"
            );
        }
        (Err(carried), Err(walked)) => assert_eq!(carried.to_string(), walked.to_string()),
        _ => panic!(
            "carried and walked owner inventories disagree: {:?} {:?}",
            carried.as_ref().err(),
            walked.as_ref().err()
        ),
    }
}
fn text_owner_addresses(
    frame: rpptx::TextFrameRef<'_>,
    slide: PartIdentity,
    shape: Option<u32>,
    cell: Option<(usize, usize)>,
    path: smallvec::SmallVec<[PathSeg; 5]>,
    out: &mut Vec<OwnerAddress>,
) -> PyResult<()> {
    out.try_reserve(1).map_err(|_| {
        pyo3::exceptions::PyMemoryError::new_err("text body owner allocation failed")
    })?;
    out.push(OwnerAddress {
        physical: PhysicalOwner::TextBody {
            slide: slide.clone(),
            shape,
            cell,
        },
        path: path.clone(),
        layout_index: None,
        table: false,
    });
    for index in 0..frame.paragraph_count() {
        let paragraph = frame
            .paragraph(index)
            .ok_or_else(|| PyValueError::new_err("physical paragraph is absent"))?;
        let mut paragraph_path = path.clone();
        paragraph_path.try_reserve(1).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("paragraph owner path allocation failed")
        })?;
        paragraph_path.push(PathSeg::Para(index));
        out.try_reserve(1).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("paragraph owner allocation failed")
        })?;
        out.push(OwnerAddress {
            physical: PhysicalOwner::Paragraph {
                slide: slide.clone(),
                shape,
                cell,
                index,
            },
            path: paragraph_path.clone(),
            layout_index: None,
            table: false,
        });
        for run in 0..paragraph.run_count() {
            let mut run_path = paragraph_path.clone();
            run_path.try_reserve(1).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("run owner path allocation failed")
            })?;
            run_path.push(PathSeg::Run(run));
            out.try_reserve(1).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("run owner allocation failed")
            })?;
            out.push(OwnerAddress {
                physical: PhysicalOwner::Run {
                    slide: slide.clone(),
                    shape,
                    cell,
                    paragraph: index,
                    index: run,
                },
                path: run_path,
                layout_index: None,
                table: false,
            });
        }
    }
    Ok(())
}
fn shape_owner_addresses(
    shape: rpptx::ShapeRef<'_>,
    slide: PartIdentity,
    path: smallvec::SmallVec<[PathSeg; 5]>,
    out: &mut Vec<OwnerAddress>,
) -> PyResult<()> {
    // AlternateContent remains an opaque physical owner. Its fallback
    // children cannot become mutable addresses by traversing this inventory.
    out.try_reserve(1).map_err(|_| {
        pyo3::exceptions::PyMemoryError::new_err("owner inventory allocation failed")
    })?;
    out.push(OwnerAddress {
        physical: PhysicalOwner::Shape {
            slide: slide.clone(),
            id: shape.non_visual_id(),
            kind: shape.kind(),
        },
        path: path.clone(),
        layout_index: None,
        table: shape.table().is_some(),
    });
    if let Some(frame) = shape.text_frame() {
        text_owner_addresses(
            frame,
            slide.clone(),
            shape.non_visual_id(),
            None,
            path.clone(),
            out,
        )?;
    }
    if let Some(table) = shape.table() {
        for row in 0..table.row_count() {
            for column in 0..table.column_count() {
                let Some(cell) = table.cell(row, column) else {
                    continue;
                };
                let mut cell_path = path.clone();
                cell_path.try_reserve(2).map_err(|_| {
                    pyo3::exceptions::PyMemoryError::new_err("cell owner path allocation failed")
                })?;
                cell_path.extend([PathSeg::Row(row), PathSeg::Cell(column)]);
                out.try_reserve(1).map_err(|_| {
                    pyo3::exceptions::PyMemoryError::new_err("cell inventory allocation failed")
                })?;
                out.push(OwnerAddress {
                    physical: PhysicalOwner::Cell {
                        slide: slide.clone(),
                        table: shape.non_visual_id(),
                        row,
                        column,
                    },
                    path: cell_path.clone(),
                    layout_index: None,
                    table: false,
                });
                if let Some(frame) = cell.text_frame() {
                    text_owner_addresses(
                        frame,
                        slide.clone(),
                        shape.non_visual_id(),
                        Some((row, column)),
                        cell_path,
                        out,
                    )?;
                }
            }
        }
    }
    if shape.kind() == rpptx::ShapeKind::Group {
        for (index, child) in shape.children().enumerate() {
            let mut nested = path.clone();
            nested.try_reserve(1).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("owner path allocation failed")
            })?;
            nested.push(PathSeg::Shape(index));
            shape_owner_addresses(child, slide.clone(), nested, out)?;
        }
    }
    Ok(())
}

/// The owners of one slide part: the slide and its shapes.
fn slide_owner_addresses(
    inner: &rpptx::Presentation,
    index: usize,
    out: &mut Vec<OwnerAddress>,
) -> PyResult<()> {
    let start = out.len();
    let slide = inner
        .slide(index)
        .ok_or_else(|| PyValueError::new_err("slide owner is absent"))?;
    out.try_reserve(1).map_err(|_| {
        pyo3::exceptions::PyMemoryError::new_err("owner inventory allocation failed")
    })?;
    out.push(OwnerAddress {
        physical: PhysicalOwner::Slide(slide.id()),
        path: smallvec![PathSeg::Slide(index)],
        layout_index: None,
        table: false,
    });
    for (shape_index, shape) in slide.shapes().enumerate() {
        shape_owner_addresses(
            shape,
            PartIdentity::Slide(slide.id()),
            smallvec![PathSeg::Slide(index), PathSeg::Shape(shape_index)],
            out,
        )?;
    }
    note_fx260_ppt(out.len() - start);
    Ok(())
}

/// The owners of the layout and master part: every layout, then the layout
/// shapes, then each master with its shapes.
fn master_owner_addresses(
    inner: &rpptx::Presentation,
    out: &mut Vec<OwnerAddress>,
) -> PyResult<()> {
    let start = out.len();
    for index in 0..inner.layout_count() {
        // Originating layout parts provide physical ownership independently of order.
        let part = inner
            .layout_part_name(index)
            .ok_or_else(|| pyo3::exceptions::PyValueError::new_err("layout owner is absent"))?;
        out.try_reserve(1).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("owner inventory allocation failed")
        })?;
        out.push(OwnerAddress {
            physical: PhysicalOwner::Layout(part.to_owned()),
            path: smallvec![],
            layout_index: Some(index),
            table: false,
        });
    }
    for index in 0..inner.layout_count() {
        let part = inner
            .layout_part_name(index)
            .expect("layout inventory")
            .to_owned();
        for (shape_index, _) in inner
            .layout(index)
            .expect("layout inventory")
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .enumerate()
        {
            shape_owner_addresses(
                inner
                    .part_shape_at(rpptx::ShapeOwner::Layout(index), &[shape_index])
                    .map_err(|e| PyValueError::new_err(e.to_string()))?,
                PartIdentity::Layout(part.clone()),
                smallvec![PathSeg::Layout(index), PathSeg::Shape(shape_index)],
                out,
            )?;
        }
    }
    for index in 0..inner.master_count() {
        let part = inner
            .slide_master_part_name(index)
            .map_err(|e| PyValueError::new_err(e.to_string()))?
            .ok_or_else(|| PyValueError::new_err("master owner absent"))?;
        out.try_reserve(1).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("master owner allocation failed")
        })?;
        out.push(OwnerAddress {
            physical: PhysicalOwner::Master(part.clone()),
            path: smallvec![PathSeg::Master(index)],
            layout_index: None,
            table: false,
        });
        if let Some(master) = inner.master(index) {
            for shape_index in 0..master.common_slide_data.shape_tree.children.len() {
                shape_owner_addresses(
                    inner
                        .part_shape_at(rpptx::ShapeOwner::Master(index), &[shape_index])
                        .map_err(|e| PyValueError::new_err(e.to_string()))?,
                    PartIdentity::Master(part.clone()),
                    smallvec![PathSeg::Master(index), PathSeg::Shape(shape_index)],
                    out,
                )?;
            }
        }
    }
    note_fx260_ppt(out.len() - start);
    Ok(())
}

impl PyPresentation {
    /// Fallible construction before any Python handles can escape.
    fn from_presentation_owned(inner: rpptx::Presentation) -> PyResult<Self> {
        use oxml_py_support::RevisionScope as S;
        let mut addresses = Vec::new();
        let mut slide_ends = Vec::new();
        for index in 0..inner.len() {
            slide_owner_addresses(&inner, index, &mut addresses)?;
            slide_ends.push(addresses.len());
        }
        master_owner_addresses(&inner, &mut addresses)?;
        let mut kinds = vec![S::SlideInventory, S::LayoutInventory, S::MasterInventory];
        for address in &addresses {
            kinds.try_reserve(4).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("owner scopes allocation failed")
            })?;
            kinds.extend([S::Owner, S::Text, S::Text]);
            if address.table {
                kinds.push(S::Grid);
            }
        }
        let mut revisions = oxml_py_support::RevisionCounter::new();
        let mut states = revisions
            .initialize_scopes(&kinds)
            .map_err(oxml_py_support::revision_to_pyerr)?
            .into_iter();
        let slides = states
            .next()
            .expect("three inventory scopes were initialized");
        let layouts = states
            .next()
            .expect("three inventory scopes were initialized");
        let masters = states
            .next()
            .expect("three inventory scopes were initialized");
        let mut parts = OwnerParts::default();
        let mut owners = Vec::new();
        for (at, address) in addresses.into_iter().enumerate() {
            if parts.slides.len() < slide_ends.len() && at == slide_ends[parts.slides.len()] {
                parts
                    .slides
                    .push(Arc::new(PartOwners::new(std::mem::take(&mut owners))?));
            }
            let owner = states.next().expect("owner scope was initialized");
            let text = states.next().expect("text scope was initialized");
            let xml_children = states.next().expect("XML child scope was initialized");
            let grid = address
                .table
                .then(|| states.next().expect("table grid scope was initialized"));
            owners.try_reserve(1).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("owner inventory allocation failed")
            })?;
            owners.push(LiveOwner {
                address,
                owner,
                text,
                xml_children,
                grid,
            });
        }
        while parts.slides.len() < slide_ends.len() {
            parts
                .slides
                .push(Arc::new(PartOwners::new(std::mem::take(&mut owners))?));
        }
        parts.masters = Arc::new(PartOwners::new(owners)?);
        Ok(Self {
            inner,
            revisions,
            owners: OwnerInventory::new(parts)?,
            gradient_stops: Vec::new(),
            next_gradient_stop: 0,
            slides,
            layouts,
            masters,
        })
    }

    fn indexed_owner_at(
        &self,
        path: &[oxml_py_support::PathSeg],
        layout_index: Option<usize>,
        body: bool,
    ) -> Option<&LiveOwner> {
        self.owners.parts.owner_at(path, layout_index, body)
    }

    fn owner_at(
        &self,
        path: &[oxml_py_support::PathSeg],
        layout_index: Option<usize>,
    ) -> Option<&LiveOwner> {
        self.indexed_owner_at(path, layout_index, false)
    }

    fn body_owner_at(&self, path: &[oxml_py_support::PathSeg]) -> Option<&LiveOwner> {
        self.indexed_owner_at(path, None, true)
    }

    pub(crate) fn current_path(
        &self,
        captured: &oxml_py_support::ContentPath,
    ) -> Option<oxml_py_support::ContentPath> {
        use oxml_py_support::PathSeg;
        if !captured.belongs_to(&self.revisions) {
            return None;
        }
        let Some(generation) = captured.owner_generation() else {
            return Some(captured.clone());
        };
        note_fx260_ppt(1);
        let owner = self.owners.owner(generation)?;
        let mut path = captured.clone();
        let end = match &owner.address.physical {
            PhysicalOwner::TextBody { cell, .. } => {
                let native_path = oxml_py_support::ContentPath::new(
                    owner.address.path.clone(),
                    self.revisions.current(),
                );
                let shape = crate::shape::native_shape_ref_at(&self.inner, &native_path)?;
                let exists = match cell {
                    Some((row, column)) => {
                        shape.table()?.cell(*row, *column)?.text_frame().is_some()
                    }
                    None => shape.text_frame().is_some(),
                };
                if !exists {
                    return None;
                }
                captured
                    .segs
                    .iter()
                    .rposition(|seg| matches!(seg, PathSeg::Shape(_) | PathSeg::Cell(_)))?
            }
            PhysicalOwner::Paragraph { .. } => captured
                .segs
                .iter()
                .rposition(|seg| matches!(seg, PathSeg::Para(_)))?,
            PhysicalOwner::Run { .. } => captured
                .segs
                .iter()
                .rposition(|seg| matches!(seg, PathSeg::Run(_)))?,
            PhysicalOwner::Cell { .. } => captured
                .segs
                .iter()
                .rposition(|seg| matches!(seg, PathSeg::Cell(_)))?,
            PhysicalOwner::Shape { .. } => captured
                .segs
                .iter()
                .rposition(|seg| matches!(seg, PathSeg::Shape(_)))?,
            PhysicalOwner::Slide(_) => captured
                .segs
                .iter()
                .position(|seg| matches!(seg, PathSeg::Slide(_)))?,
            PhysicalOwner::Layout(_) => {
                if let Some(index) = owner.address.layout_index {
                    for segment in &mut path.segs {
                        if let PathSeg::Layout(at) = segment {
                            *at = index;
                        }
                    }
                }
                return Some(path);
            }
            PhysicalOwner::Master(_) => captured
                .segs
                .iter()
                .position(|seg| matches!(seg, PathSeg::Master(_)))?,
        };
        let suffix = captured.segs.iter().skip(end + 1).copied();
        path.segs = owner.address.path.clone();
        path.segs.extend(suffix);
        Some(path)
    }

    /// ID-based native operations require one qualified current physical owner.
    /// A parsed duplicate ID remains inspectable but cannot select another shape.
    pub(crate) fn checked_shape_id(&self, path: &oxml_py_support::ContentPath) -> PyResult<u32> {
        let current = self
            .current_path(path)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("shape owner is absent"))?;
        let owner = self.owner_at(&current.segs, None).ok_or_else(|| {
            pyo3::exceptions::PyValueError::new_err("shape is not a modeled mutable owner")
        })?;
        let PhysicalOwner::Shape {
            slide,
            id: Some(id),
            kind,
        } = &owner.address.physical
        else {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "shape has no qualified id",
            ));
        };
        if *kind == rpptx::ShapeKind::AlternateContent
            || self.owners.shape_ids.get(&(slide.clone(), *id)) != Some(&1)
        {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "shape id is ambiguous or read-only",
            ));
        }
        Ok(*id)
    }

    pub(crate) fn capture_path(
        &self,
        segs: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
        descendants: Option<oxml_py_support::RevisionScope>,
    ) -> PyResult<oxml_py_support::ContentPath> {
        self.capture_owner_path(
            &self.owners.parts,
            &self.slides,
            self.revisions.current(),
            segs,
            (descendants, false, false),
        )
    }

    pub(crate) fn capture_prepared_path(
        &self,
        prepared: &PreparedOwners,
        segs: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
        descendants: Option<oxml_py_support::RevisionScope>,
    ) -> PyResult<oxml_py_support::ContentPath> {
        match &prepared.owners {
            OwnerChanges::Replace(change) => self.capture_owner_path(
                &change.parts,
                &prepared.slides,
                prepared.revision.committed_serial(),
                segs,
                (descendants, false, false),
            ),
            OwnerChanges::Append(owners) => {
                use oxml_py_support::PathSeg;
                let [PathSeg::Slide(slide), PathSeg::Shape(_)] = segs.as_slice() else {
                    return Err(PyValueError::new_err("append requires a root shape owner"));
                };
                if descendants.is_some() {
                    return Err(PyValueError::new_err(
                        "append does not capture descendant views",
                    ));
                }
                let part = self
                    .owner_at(&[PathSeg::Slide(*slide)], None)
                    .ok_or_else(|| PyValueError::new_err("append slide owner is absent"))?;
                let shape = owners
                    .iter()
                    .find(|owner| {
                        note_fx260_ppt(1);
                        owner.address.path == segs
                            && matches!(owner.address.physical, PhysicalOwner::Shape { .. })
                    })
                    .ok_or_else(|| PyValueError::new_err("append shape owner is absent"))?;
                let mut path = self
                    .revisions
                    .capture_scopes(segs, &[part.owner.clone(), shape.owner.clone()])
                    .map_err(oxml_py_support::revision_to_pyerr)?;
                path.revision = prepared.revision.committed_serial();
                Ok(path)
            }
        }
    }

    pub(crate) fn capture_text_body_path(
        &self,
        segs: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
        descendants: Option<oxml_py_support::RevisionScope>,
    ) -> PyResult<oxml_py_support::ContentPath> {
        self.capture_owner_path(
            &self.owners.parts,
            &self.slides,
            self.revisions.current(),
            segs,
            (descendants, true, false),
        )
    }

    pub(crate) fn capture_xml_children_path(
        &self,
        segs: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
    ) -> PyResult<oxml_py_support::ContentPath> {
        self.capture_owner_path(
            &self.owners.parts,
            &self.slides,
            self.revisions.current(),
            segs,
            (None, false, true),
        )
    }

    fn capture_owner_path(
        &self,
        parts: &OwnerParts,
        slides: &oxml_py_support::RevisionState,
        serial: u64,
        segs: smallvec::SmallVec<[oxml_py_support::PathSeg; 5]>,
        selection: (Option<oxml_py_support::RevisionScope>, bool, bool),
    ) -> PyResult<oxml_py_support::ContentPath> {
        use oxml_py_support::PathSeg;
        let (descendants, select_body, xml_children) = selection;
        let owner_at = |path: &[PathSeg], layout_index: Option<usize>, body: bool| {
            parts.owner_at(path, layout_index, body)
        };
        if segs.is_empty() {
            let mut path = self
                .revisions
                .capture_scopes(segs, std::slice::from_ref(slides))
                .map_err(oxml_py_support::revision_to_pyerr)?;
            path.revision = serial;
            return Ok(path);
        }
        let physical = segs
            .iter()
            .filter(|seg| {
                matches!(
                    seg,
                    PathSeg::Slide(_) | PathSeg::Layout(_) | PathSeg::Master(_) | PathSeg::Shape(_)
                )
            })
            .copied()
            .collect::<smallvec::SmallVec<[PathSeg; 5]>>();
        let cell = segs.iter().any(|seg| matches!(seg, PathSeg::Row(_)))
            && segs.iter().any(|seg| matches!(seg, PathSeg::Cell(_)));
        let styled = segs.iter().any(|seg| matches!(seg, PathSeg::TextStyle(_)));
        let node = !styled
            && segs
                .iter()
                .any(|seg| matches!(seg, PathSeg::Para(_) | PathSeg::Run(_)));
        let base = if cell {
            segs.iter()
                .filter(|seg| !matches!(seg, PathSeg::Para(_) | PathSeg::Run(_)))
                .copied()
                .collect::<smallvec::SmallVec<[PathSeg; 5]>>()
        } else {
            physical.clone()
        };
        let owner_path = if node { segs.clone() } else { base.clone() };
        // An absent cell body has one established virtual paragraph view. It
        // remains bound to the cell, never inventing a physical paragraph.
        let virtual_cell = cell
            && node
            && !select_body
            && segs.last() == Some(&PathSeg::Para(0))
            && owner_at(&base, None, true).is_none();
        let direct = if virtual_cell {
            owner_at(&base, None, false)
        } else {
            match owner_path.as_slice() {
                [PathSeg::Layout(index)] if !select_body => owner_at(&[], Some(*index), false),
                _ => owner_at(&owner_path, None, select_body),
            }
        };
        let owner = match direct {
            Some(owner) => owner,
            None => {
                // An opaque ancestor shares the first segment, so its part.
                let opaque_parent = !select_body
                    && parts.part(OwnerPart::of(&owner_path)).is_some_and(|part| {
                        part.owners.iter().any(|owner| {
                            matches!(
                                &owner.address.physical,
                                PhysicalOwner::Shape {
                                    kind: rpptx::ShapeKind::AlternateContent,
                                    ..
                                }
                            ) && owner_path.starts_with(&owner.address.path)
                                && owner_path.len() > owner.address.path.len()
                        })
                    });
                if opaque_parent && serial == self.revisions.current() {
                    return Ok(self.revisions.capture(segs));
                }
                if select_body {
                    return Err(pyo3::exceptions::PyNotImplementedError::new_err(
                        "TextFrame.element requires an existing physical text body, not a virtual cell or inherited text-style view",
                    ));
                }
                return Err(pyo3::exceptions::PyIndexError::new_err(
                    "physical XML owner is absent",
                ));
            }
        };
        let part_path = physical
            .first()
            .copied()
            .into_iter()
            .collect::<smallvec::SmallVec<[PathSeg; 5]>>();
        let part = match part_path.as_slice() {
            [PathSeg::Layout(index)] => owner_at(&[], Some(*index), false),
            _ => owner_at(&part_path, None, false),
        }
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("part owner is absent"))?;
        let mut scopes = smallvec::SmallVec::<[oxml_py_support::RevisionState; 4]>::new();
        scopes.try_reserve(segs.len() + 4).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("captured owner scopes allocation failed")
        })?;
        scopes.push(part.owner.clone());
        let mut prefix = smallvec::SmallVec::<[PathSeg; 5]>::new();
        prefix.try_reserve(segs.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("captured owner path allocation failed")
        })?;
        for segment in &segs {
            prefix.push(*segment);
            if matches!(segment, PathSeg::Para(_)) && !styled && !virtual_cell {
                let body = owner_at(&base, None, true).ok_or_else(|| {
                    pyo3::exceptions::PyIndexError::new_err("text body owner is absent")
                })?;
                if !scopes.iter().any(|state| state.same_state(&body.owner)) {
                    scopes.push(body.owner.clone());
                }
            }
            if matches!(
                segment,
                PathSeg::Shape(_) | PathSeg::Cell(_) | PathSeg::Para(_) | PathSeg::Run(_)
            ) && !styled
                && (!matches!(segment, PathSeg::Shape(_)) || prefix == physical)
                && let Some(ancestor) = owner_at(&prefix, None, false)
                && !scopes.iter().any(|state| state.same_state(&ancestor.owner))
            {
                scopes.push(ancestor.owner.clone());
            }
        }
        if !scopes.iter().any(|state| state.same_state(&owner.owner)) {
            scopes.push(owner.owner.clone());
        }
        if descendants == Some(oxml_py_support::RevisionScope::Grid)
            || cell
            || segs.iter().any(|seg| matches!(seg, PathSeg::Row(_)))
        {
            let table = owner_at(&physical, None, false)
                .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("table owner is absent"))?;
            scopes.push(
                table
                    .grid
                    .as_ref()
                    .ok_or_else(|| PyValueError::new_err("owner has no table grid"))?
                    .clone(),
            );
        }
        if descendants == Some(oxml_py_support::RevisionScope::Text) || virtual_cell || styled {
            scopes.push(owner.text.clone());
        }
        if xml_children || virtual_cell || styled {
            scopes.push(owner.xml_children.clone());
        }
        let mut path = self
            .revisions
            .capture_scopes(segs, &scopes)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        path.revision = serial;
        Ok(path)
    }

    pub(crate) fn capture_layout(&self, index: usize) -> PyResult<oxml_py_support::ContentPath> {
        let owner = self
            .owner_at(&[], Some(index))
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("layout owner is absent"))?;
        self.revisions
            .capture_scopes(smallvec![], std::slice::from_ref(&owner.owner))
            .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))
    }

    pub(crate) fn current_layout_index(
        &self,
        path: &oxml_py_support::ContentPath,
    ) -> Option<usize> {
        self.owners
            .owner(path.owner_generation()?)?
            .address
            .layout_index
    }

    /// Build the candidate inventory, checking all survival proofs.
    /// It never infers that an ID still names the same incarnation.
    ///
    /// A part whose native records the candidate still shares, and whose
    /// owners all survive unchanged, is kept. Every other part is walked
    /// again, so the work follows the parts an edit touches.
    pub(crate) fn prepare_owners(
        &self,
        candidate: &rpptx::Presentation,
        operation: &'static str,
        preserved: &OwnerProofs,
        slides_changed: bool,
        layouts_changed: bool,
        masters_changed: bool,
    ) -> PyResult<PreparedOwners> {
        let changed = [slides_changed, layouts_changed, masters_changed];
        let carried = self.carried_parts(candidate, preserved);
        let prepared = self.prepare_parts(candidate, operation, preserved, &carried, changed);
        // Debug builds, which run the binding tests, check that carrying
        // parts over prepares exactly what walking every part prepares.
        #[cfg(debug_assertions)]
        if carried.masters || carried.slides.iter().any(Option::is_some) {
            let walked = without_work_count(|| {
                let none = CarriedParts::none(candidate);
                self.prepare_parts(candidate, operation, preserved, &none, changed)
            });
            assert_same_preparation(&prepared, &walked);
        }
        prepared
    }

    /// The live parts a candidate keeps: those whose native records it
    /// shares, at the ordinal the proofs move them to, without an explicit
    /// proof. Any doubt walks every part.
    fn carried_parts(&self, candidate: &rpptx::Presentation, proofs: &OwnerProofs) -> CarriedParts {
        let walk_all = CarriedParts::none(candidate);
        // Shape ID counts stay local to a part while slide IDs are unique.
        let unique = |inner: &rpptx::Presentation| {
            let mut ids = HashSet::new();
            inner.slides().all(|slide| ids.insert(slide.id()))
        };
        if self.owners.parts.slides.len() != self.inner.len()
            || !unique(&self.inner)
            || !unique(candidate)
        {
            return walk_all;
        }
        let listed = proofs
            .parts
            .iter()
            .map(|(part, _)| *part)
            .collect::<HashSet<_>>();
        let slides = candidate
            .unchanged_slides_since(&self.inner)
            .into_iter()
            .enumerate()
            .map(|(at, origin)| {
                origin.filter(|from| {
                    proofs.edit.slide_index(*from) == Some(at)
                        && !listed.contains(&OwnerPart::Slide(*from))
                })
            })
            .collect::<Vec<_>>();
        // Master part names come from the presentation relationships, not
        // from the master records, so they are compared too.
        let masters = !matches!(proofs.edit, PartEdit::RemoveLayout(_))
            && !listed.contains(&OwnerPart::Masters)
            && candidate.unchanged_layouts_and_masters_since(&self.inner)
            && (0..candidate.master_count()).all(|index| {
                let live = self.owners.parts.owner_at(&[PathSeg::Master(index)], None, false);
                matches!(
                    (live.map(|owner| &owner.address.physical), candidate.slide_master_part_name(index)),
                    (Some(PhysicalOwner::Master(live)), Ok(Some(part))) if *live == part
                )
            });
        let carried = CarriedParts { slides, masters };
        // Explicit proofs must prove owners of their own part and land in a
        // part walked again, as every proof of a full walk does.
        for (part, list) in &proofs.parts {
            for proof in list {
                let own = self
                    .owners
                    .by_generation
                    .get(&proof.generation)
                    .is_some_and(|(at, _)| at == part);
                if !own || carried.keeps(OwnerPart::of(&proof.destination)) {
                    return walk_all;
                }
            }
        }
        carried
    }

    fn prepare_parts(
        &self,
        candidate: &rpptx::Presentation,
        operation: &'static str,
        proofs: &OwnerProofs,
        carried: &CarriedParts,
        [slides_changed, layouts_changed, masters_changed]: [bool; 3],
    ) -> PyResult<PreparedOwners> {
        use oxml_py_support::RevisionScope as S;
        let fail = |message: &str| pyo3::exceptions::PyValueError::new_err(message.to_owned());
        let live = &self.owners.parts;
        let kept = carried.kept_live_slides(live.slides.len());
        let kept_part = |part: OwnerPart| match part {
            OwnerPart::Slide(index) => kept[index],
            OwnerPart::Masters => carried.masters,
        };
        // Live parts nothing carries over, with the proofs of their owners.
        let mut old_parts = Vec::new();
        for (part, owners) in live.iter() {
            if !kept_part(part) {
                old_parts.push((part, owners, proofs.for_part(part, owners)?));
            }
        }
        // Candidate parts walked again, in inventory order.
        let mut new_parts = Vec::new();
        for (at, origin) in carried.slides.iter().enumerate() {
            if origin.is_none() {
                let mut addresses = Vec::new();
                slide_owner_addresses(candidate, at, &mut addresses)?;
                new_parts.push((OwnerPart::Slide(at), addresses));
            }
        }
        if !carried.masters {
            let mut addresses = Vec::new();
            master_owner_addresses(candidate, &mut addresses)?;
            new_parts.push((OwnerPart::Masters, addresses));
        }
        let mut destinations = HashMap::new();
        for (part_at, (_, addresses)) in new_parts.iter().enumerate() {
            destinations.try_reserve(addresses.len()).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("survival destinations allocation failed")
            })?;
            for (index, address) in addresses.iter().enumerate() {
                let key = (
                    address.path.clone(),
                    address.layout_index,
                    std::mem::discriminant(&address.physical),
                );
                destinations.entry(key).or_insert((part_at, index));
            }
        }
        let mut claims = new_parts
            .iter()
            .map(|(_, addresses)| vec![None; addresses.len()])
            .collect::<Vec<_>>();
        let mut seen = HashSet::new();
        let mut retained = Vec::new();
        let mut retained_ends = Vec::new();
        for (_, _, list) in &old_parts {
            for proof in list.iter() {
                note_fx260_ppt(1);
                if !seen.insert(proof.generation) {
                    return Err(fail("duplicate owner survival proof"));
                }
                let old = self
                    .owners
                    .owner(proof.generation)
                    .ok_or_else(|| fail("survival proof has no current owner"))?;
                let key = (
                    proof.destination.clone(),
                    proof.layout_index,
                    std::mem::discriminant(&old.address.physical),
                );
                let &(part_at, index) = destinations
                    .get(&key)
                    .ok_or_else(|| fail("survival destination is absent"))?;
                if claims[part_at][index].is_some() {
                    return Err(fail("two owners claim one destination"));
                }
                let next = &new_parts[part_at].1[index];
                if old.address.physical != next.physical || old.address.table != next.table {
                    return Err(fail("survival proof changes physical owner identity"));
                }
                claims[part_at][index] = Some(retained.len());
                retained.push((old, proof));
            }
            retained_ends.push(retained.len());
        }
        let next_len = carried
            .slides
            .iter()
            .flatten()
            .map(|from| live.slides[*from].owners.len())
            .chain(carried.masters.then(|| live.masters.owners.len()))
            .chain(new_parts.iter().map(|(_, addresses)| addresses.len()))
            .sum::<usize>();
        let structural = next_len != live.len()
            || carried
                .slides
                .iter()
                .enumerate()
                .any(|(at, from)| from.is_some_and(|from| from != at))
            || retained.iter().any(|(old, proof)| {
                old.address.path != proof.destination
                    || old.address.layout_index != proof.layout_index
            });
        if structural {
            // A mutation cannot disambiguate imported missing/duplicate IDs by
            // merely finding the same number again in a candidate tree. The
            // candidate counts are the live ones with the walked parts swapped.
            let mut walked_ids = HashMap::<(PartIdentity, u32), (usize, usize)>::new();
            for (_, owners, _) in &old_parts {
                for (key, count) in &owners.shape_ids {
                    walked_ids.entry(key.clone()).or_default().0 += count;
                }
            }
            for (_, addresses) in &new_parts {
                for address in addresses {
                    if let PhysicalOwner::Shape {
                        slide,
                        id: Some(id),
                        ..
                    } = &address.physical
                    {
                        walked_ids.entry((slide.clone(), *id)).or_default().1 += 1;
                    }
                }
            }
            let mut walked = 0_usize;
            for (part, owners) in live.iter() {
                if kept_part(part) {
                    if let Some(refusal) = owners.refusal {
                        return Err(fail(refusal));
                    }
                    continue;
                }
                let start = walked.checked_sub(1).map_or(0, |at| retained_ends[at]);
                for (old, _) in &retained[start..retained_ends[walked]] {
                    if let PhysicalOwner::Shape { slide, id, kind } = &old.address.physical {
                        if *kind == rpptx::ShapeKind::AlternateContent {
                            continue;
                        }
                        let Some(id) = id else {
                            return Err(fail("structural owner proof requires a shape id"));
                        };
                        let key = (slide.clone(), *id);
                        let original = self.owners.shape_ids.get(&key).copied().unwrap_or(0);
                        let (before, after) = walked_ids.get(&key).copied().unwrap_or_default();
                        if original != 1 || original - before + after != 1 {
                            return Err(fail("structural owner proof has an ambiguous shape id"));
                        }
                    }
                }
                walked += 1;
            }
        }
        let mut affected = Vec::new();
        let walked_owners = old_parts
            .iter()
            .map(|(_, owners, _)| owners.owners.len())
            .sum::<usize>();
        let affected_capacity = walked_owners
            .checked_mul(4)
            .and_then(|count| count.checked_add(3))
            .ok_or_else(|| {
                pyo3::exceptions::PyOverflowError::new_err("owner retirement count overflow")
            })?;
        affected.try_reserve_exact(affected_capacity).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("owner retirement allocation failed")
        })?;
        for (_, owners, _) in &old_parts {
            for old in &owners.owners {
                if !seen.contains(&old.owner.generation()) {
                    affected.push(old.owner.clone());
                    affected.push(old.text.clone());
                    affected.push(old.xml_children.clone());
                    if let Some(grid) = &old.grid {
                        affected.push(grid.clone());
                    }
                }
            }
        }
        let mut successors = Vec::new();
        successors.try_reserve_exact(retained.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("owner replacements allocation failed")
        })?;
        for (old, proof) in &retained {
            let mut push = |state: &oxml_py_support::RevisionState| {
                affected.push(state.clone());
                affected.len() - 1
            };
            let xml_children = proof.xml_children_changed.then(|| push(&old.xml_children));
            let text = proof.text_changed.then(|| push(&old.text));
            let grid = if proof.grid_changed {
                Some(push(
                    old.grid
                        .as_ref()
                        .ok_or_else(|| fail("grid change on non-table"))?,
                ))
            } else {
                None
            };
            successors.push((xml_children, text, grid));
        }
        let slides_at = slides_changed.then(|| {
            let index = affected.len();
            affected.push(self.slides.clone());
            index
        });
        let layouts_at = layouts_changed.then(|| {
            let index = affected.len();
            affected.push(self.layouts.clone());
            index
        });
        let masters_at = masters_changed.then(|| {
            let index = affected.len();
            affected.push(self.masters.clone());
            index
        });
        let mut births = Vec::new();
        let walked_addresses = new_parts
            .iter()
            .map(|(_, addresses)| addresses.len())
            .sum::<usize>();
        births
            .try_reserve_exact(walked_addresses.checked_mul(4).ok_or_else(|| {
                pyo3::exceptions::PyOverflowError::new_err("owner birth count overflow")
            })?)
            .map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("owner births allocation failed")
            })?;
        for (part_at, (_, addresses)) in new_parts.iter().enumerate() {
            for (index, address) in addresses.iter().enumerate() {
                if claims[part_at][index].is_none() {
                    births.extend([S::Owner, S::Text, S::Text]);
                    if address.table {
                        births.push(S::Grid);
                    }
                }
            }
        }
        let revision = self
            .revisions
            .prepare_scopes(operation, &affected, &births)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let mut born = 0;
        let mut rebuilt = Vec::new();
        for (part_at, (part, addresses)) in new_parts.into_iter().enumerate() {
            let mut owners = Vec::new();
            owners.try_reserve_exact(addresses.len()).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err(
                    "candidate owner inventory allocation failed",
                )
            })?;
            for (index, address) in addresses.into_iter().enumerate() {
                if let Some(at) = claims[part_at][index] {
                    let (old, _) = retained[at];
                    let (xml_children, text, grid) = successors[at];
                    let successor = |at: usize| {
                        revision
                            .successor(at)
                            .expect("affected scope successor exists")
                    };
                    let mut owner = old.clone();
                    owner.address = address;
                    if let Some(at) = text {
                        owner.text = successor(at);
                    }
                    if let Some(at) = grid {
                        owner.grid = Some(successor(at));
                    }
                    if let Some(at) = xml_children {
                        owner.xml_children = successor(at);
                    }
                    owners.push(owner);
                } else {
                    let start = born;
                    born += 3 + usize::from(address.table);
                    let grid = address
                        .table
                        .then(|| revision.born(start + 3).expect("table birth exists"));
                    owners.push(LiveOwner {
                        address,
                        owner: revision.born(start).expect("owner birth exists"),
                        text: revision.born(start + 1).expect("text birth exists"),
                        xml_children: revision.born(start + 2).expect("XML child birth exists"),
                        grid,
                    });
                }
            }
            rebuilt.push((part, Arc::new(PartOwners::new(owners)?)));
        }
        let mut parts = OwnerParts::default();
        let mut moved = Vec::new();
        let mut walked = rebuilt.iter();
        for (at, from) in carried.slides.iter().enumerate() {
            parts.slides.push(match from {
                Some(from) if *from == at => Arc::clone(&live.slides[*from]),
                Some(from) => {
                    moved.push(OwnerPart::Slide(at));
                    Arc::new(live.slides[*from].moved_to(at)?)
                }
                None => Arc::clone(&walked.next().expect("walked slide part").1),
            });
        }
        parts.masters = if carried.masters {
            Arc::clone(&live.masters)
        } else {
            Arc::clone(&walked.next().expect("walked master part").1)
        };
        let slides = slides_at.map_or_else(
            || self.slides.clone(),
            |at| revision.successor(at).expect("inventory successor exists"),
        );
        let layouts = layouts_at.map_or_else(
            || self.layouts.clone(),
            |at| revision.successor(at).expect("inventory successor exists"),
        );
        let masters = masters_at.map_or_else(
            || self.masters.clone(),
            |at| revision.successor(at).expect("inventory successor exists"),
        );
        Ok(PreparedOwners {
            revision,
            owners: OwnerChanges::Replace(InventoryChange {
                parts,
                dropped: old_parts.iter().map(|(part, _, _)| *part).collect(),
                rebuilt: rebuilt.into_iter().map(|(part, _)| part).collect(),
                moved,
            }),
            slides,
            layouts,
            masters,
            #[cfg(debug_assertions)]
            retired: affected.iter().map(|state| state.generation()).collect(),
        })
    }

    /// No fallible action or Python allocation follows successful token commit.
    pub(crate) fn publish_owners(
        &mut self,
        candidate: rpptx::Presentation,
        prepared: PreparedOwners,
    ) -> PyResult<()> {
        self.commit_prepared(prepared)?;
        self.inner = candidate;
        Ok(())
    }

    pub(crate) fn commit_prepared(&mut self, prepared: PreparedOwners) -> PyResult<()> {
        match &prepared.owners {
            OwnerChanges::Replace(change) => self.owners.reserve(change)?,
            OwnerChanges::Append(owners) => {
                self.owners
                    .by_generation
                    .try_reserve(owners.len())
                    .map_err(|_| {
                        pyo3::exceptions::PyMemoryError::new_err(
                            "PowerPoint owner index allocation failed",
                        )
                    })?;
                self.owners.shape_ids.try_reserve(1).map_err(|_| {
                    pyo3::exceptions::PyMemoryError::new_err(
                        "PowerPoint owner index allocation failed",
                    )
                })?;
            }
        }
        self.revisions
            .commit_scopes(prepared.revision)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        match prepared.owners {
            OwnerChanges::Replace(change) => self.owners.apply(change),
            OwnerChanges::Append(owners) => {
                let Some(&PathSeg::Slide(slide)) = owners[0].address.path.first() else {
                    unreachable!("appended owners belong to a slide");
                };
                let part = OwnerPart::Slide(slide);
                let owners_part = self
                    .owners
                    .parts
                    .part_mut(part)
                    .expect("appended slide part");
                for owner in owners {
                    let index = owners_part.owners.len();
                    if let PhysicalOwner::Shape {
                        slide,
                        id: Some(id),
                        ..
                    } = &owner.address.physical
                    {
                        *self
                            .owners
                            .shape_ids
                            .entry((slide.clone(), *id))
                            .or_default() += 1;
                    }
                    self.owners
                        .by_generation
                        .insert(owner.owner.generation(), (part, index));
                    owners_part.owners.push(owner);
                    owners_part.index(index);
                }
                owners_part.refusal = owners_part.structural_refusal();
            }
        }
        self.slides = prepared.slides;
        self.layouts = prepared.layouts;
        self.masters = prepared.masters;
        Ok(())
    }

    /// A property edit that keeps every physical owner identity.
    /// The document serial advances and no slide inventory is rebuilt.
    pub(crate) fn prepare_local_property(
        &self,
        operation: &'static str,
    ) -> PyResult<oxml_py_support::PreparedRevision> {
        note_fx260_ppt(1);
        self.revisions
            .prepare_scopes(operation, &[], &[])
            .map_err(oxml_py_support::revision_to_pyerr)
    }

    pub(crate) fn commit_local_property(
        &mut self,
        revision: oxml_py_support::PreparedRevision,
    ) -> PyResult<()> {
        self.revisions
            .commit_scopes(revision)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        Ok(())
    }

    /// Birth the shape, text body and empty paragraph of one appended text box.
    /// Any other new child rebuilds the owner inventory.
    pub(crate) fn prepare_appended_textbox(
        &mut self,
        operation: &'static str,
        slide_index: usize,
        shape_index: usize,
    ) -> PyResult<PreparedOwners> {
        use oxml_py_support::PathSeg;
        use oxml_py_support::RevisionScope as S;
        self.owners
            .parts
            .part_mut(OwnerPart::Slide(slide_index))
            .ok_or_else(|| PyValueError::new_err("slide owner is absent"))?
            .reserve(3)?;
        let slide_id = self
            .inner
            .slide(slide_index)
            .ok_or_else(|| PyValueError::new_err("slide owner is absent"))?
            .id();
        // The constructor has fixed shape/body/paragraph arity. Its native ID
        // is bound after successful append, before installing these owners.
        let shape_id = None;
        let kind = rpptx::ShapeKind::Shape;
        note_fx260_ppt(1);
        let slide_identity = PartIdentity::Slide(slide_id);
        let shape_path = smallvec![PathSeg::Slide(slide_index), PathSeg::Shape(shape_index)];
        let mut paragraph_path = shape_path.clone();
        paragraph_path.push(PathSeg::Para(0));
        let addresses = [
            OwnerAddress {
                physical: PhysicalOwner::Shape {
                    slide: slide_identity.clone(),
                    id: shape_id,
                    kind,
                },
                path: shape_path.clone(),
                layout_index: None,
                table: false,
            },
            OwnerAddress {
                physical: PhysicalOwner::TextBody {
                    slide: slide_identity.clone(),
                    shape: shape_id,
                    cell: None,
                },
                path: shape_path,
                layout_index: None,
                table: false,
            },
            OwnerAddress {
                physical: PhysicalOwner::Paragraph {
                    slide: slide_identity,
                    shape: shape_id,
                    cell: None,
                    index: 0,
                },
                path: paragraph_path,
                layout_index: None,
                table: false,
            },
        ];
        let births = [
            S::Owner,
            S::Text,
            S::Text,
            S::Owner,
            S::Text,
            S::Text,
            S::Owner,
            S::Text,
            S::Text,
        ];
        let revision = self
            .revisions
            .prepare_scopes(operation, &[], &births)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let born = addresses
            .into_iter()
            .enumerate()
            .map(|(index, address)| {
                let start = index * 3;
                LiveOwner {
                    address,
                    owner: revision.born(start).expect("textbox owner birth"),
                    text: revision.born(start + 1).expect("textbox text birth"),
                    xml_children: revision.born(start + 2).expect("textbox xml birth"),
                    grid: None,
                }
            })
            .collect::<Vec<_>>();
        Ok(PreparedOwners {
            revision,
            owners: OwnerChanges::Append(born),
            slides: self.slides.clone(),
            layouts: self.layouts.clone(),
            masters: self.masters.clone(),
            #[cfg(debug_assertions)]
            retired: Vec::new(),
        })
    }

    pub(crate) fn bind_appended_textbox(
        &self,
        prepared: &mut PreparedOwners,
        slide_index: usize,
        shape_index: usize,
    ) {
        let shape = self
            .inner
            .slide(slide_index)
            .expect("validated append slide")
            .shape(shape_index)
            .expect("successful textbox append");
        let id = shape.non_visual_id();
        let kind = shape.kind();
        let OwnerChanges::Append(owners) = &mut prepared.owners else {
            unreachable!("textbox append has prepared births");
        };
        for owner in owners {
            match &mut owner.address.physical {
                PhysicalOwner::Shape {
                    id: stored,
                    kind: stored_kind,
                    ..
                } => {
                    *stored = id;
                    *stored_kind = kind;
                }
                PhysicalOwner::TextBody { shape, .. } | PhysicalOwner::Paragraph { shape, .. } => {
                    *shape = id
                }
                _ => unreachable!("textbox births have fixed owner kinds"),
            }
        }
    }

    fn element_xml_survivors(
        &self,
        path: &oxml_py_support::ContentPath,
        kind: PptElementKind,
        generation: u64,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        let mut parts = self.prefix_parts(&path.segs);
        if let Some((part, _)) = self.owners.by_generation.get(&generation)
            && !parts.contains(part)
        {
            parts.push(*part);
        }
        self.part_owner_proofs(parts, |old, proof| {
            let old_path = if let Some(index) = old.address.layout_index {
                smallvec![PathSeg::Layout(index)]
            } else {
                old.address.path.clone()
            };
            if proof.generation == generation {
                proof.xml_children_changed = kind != PptElementKind::Table;
                proof.text_changed = kind != PptElementKind::Run;
                proof.grid_changed = old.address.table;
                return Ok(true);
            }
            if !old_path.starts_with(&path.segs) {
                return Ok(true);
            }
            Ok(match kind {
                PptElementKind::Run => true,
                PptElementKind::Paragraph => {
                    !matches!(old.address.physical, PhysicalOwner::Run { .. })
                }
                PptElementKind::TextBody => !matches!(
                    old.address.physical,
                    PhysicalOwner::Paragraph { .. } | PhysicalOwner::Run { .. }
                ),
                _ => false,
            })
        })
    }

    fn prepare_shape_xml(
        &mut self,
        path: &oxml_py_support::ContentPath,
    ) -> PyResult<PreparedShapeXml> {
        let (part, selected) = *self
            .owners
            .by_generation
            .get(
                &path
                    .owner_generation()
                    .ok_or_else(|| PyValueError::new_err("Element owner is absent"))?,
            )
            .ok_or_else(|| PyValueError::new_err("Element owner is absent"))?;
        let owners = self
            .owners
            .parts
            .part(part)
            .ok_or_else(|| PyValueError::new_err("Element owner is absent"))?;
        let PhysicalOwner::Shape { slide, .. } = &owners.owners[selected].address.physical else {
            return Err(PyValueError::new_err("Element is not a shape owner"));
        };
        let shape = crate::shape::native_shape_ref_at(&self.inner, path)
            .ok_or_else(|| PyValueError::new_err("shape owner is absent"))?;
        let mut addresses = Vec::new();
        shape_owner_addresses(shape, slide.clone(), path.segs.clone(), &mut addresses)?;
        let mut indices = Vec::with_capacity(addresses.len());
        for address in addresses {
            note_fx260_ppt(1);
            let body = matches!(address.physical, PhysicalOwner::TextBody { .. });
            let key = (address.path, address.layout_index, body);
            let index = *owners
                .by_path
                .get(&key)
                .ok_or_else(|| PyValueError::new_err("local shape descendant owner is absent"))?;
            if owners.owners[index].address.physical != address.physical {
                return Err(PyValueError::new_err(
                    "local shape descendant identity changed",
                ));
            }
            indices.push(index);
        }
        let mut retired = Vec::new();
        for &index in &indices {
            note_fx260_ppt(1);
            let owner = &owners.owners[index];
            if index != selected {
                retired.push(owner.owner.clone());
            }
            retired.push(owner.text.clone());
            retired.push(owner.xml_children.clone());
            if let Some(grid) = &owner.grid {
                retired.push(grid.clone());
            }
        }
        let revision = self
            .revisions
            .prepare_scopes("Element.replace_xml", &retired, &[])
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let mut next = 0;
        let updates = indices
            .into_iter()
            .map(|index| {
                note_fx260_ppt(1);
                let mut owner = owners.owners[index].clone();
                if index != selected {
                    owner.owner = revision
                        .successor(next)
                        .expect("shape child owner successor");
                    next += 1;
                }
                owner.text = revision.successor(next).expect("shape text successor");
                owner.xml_children = revision.successor(next + 1).expect("shape XML successor");
                next += 2;
                if owner.grid.is_some() {
                    owner.grid = Some(revision.successor(next).expect("shape grid successor"));
                    next += 1;
                }
                (index, owner)
            })
            .collect::<Vec<_>>();
        self.owners
            .by_generation
            .try_reserve(updates.len())
            .map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("shape successor index allocation failed")
            })?;
        // Unshare the part before the commit, after which nothing may fail.
        self.owners.parts.part_mut(part);
        Ok(PreparedShapeXml {
            revision,
            part,
            updates,
        })
    }

    fn commit_shape_xml(&mut self, prepared: PreparedShapeXml) -> PyResult<()> {
        self.revisions
            .commit_scopes(prepared.revision)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let owners = self
            .owners
            .parts
            .part_mut(prepared.part)
            .expect("prepared shape part");
        for (index, owner) in prepared.updates {
            note_fx260_ppt(1);
            self.owners
                .by_generation
                .remove(&owners.owners[index].owner.generation());
            self.owners
                .by_generation
                .insert(owner.owner.generation(), (prepared.part, index));
            owners.owners[index] = owner;
        }
        Ok(())
    }
}

#[derive(PartialEq, Eq)]
struct ShapeStructure {
    id: Option<u32>,
    kind: rpptx::ShapeKind,
    has_table: bool,
    paragraph_runs: Option<Vec<usize>>,
}

fn shape_structure(
    inner: &rpptx::Presentation,
    path: &oxml_py_support::ContentPath,
) -> PyResult<ShapeStructure> {
    let shape = crate::shape::native_shape_ref_at(inner, path)
        .ok_or_else(|| PyValueError::new_err("shape owner is absent"))?;
    let paragraph_runs = shape.text_frame().map(|frame| {
        (0..frame.paragraph_count())
            .map(|index| {
                frame
                    .paragraph(index)
                    .expect("current paragraph index")
                    .run_count()
            })
            .collect()
    });
    Ok(ShapeStructure {
        id: shape.non_visual_id(),
        kind: shape.kind(),
        has_table: shape.table().is_some(),
        paragraph_runs,
    })
}

fn projected_runs(body: Option<&oxml_drawing::text::CT_TextBody>) -> Option<Vec<usize>> {
    body.map(|body| {
        body.paragraphs()
            .iter()
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .filter(|run| matches!(run, oxml_drawing::text::TextRun::Run(_)))
                    .count()
            })
            .collect()
    })
}

fn geometry_witness(
    shape: &rpptx::ShapeRef<'_>,
    before: Option<&oxml_drawing::xfrm::CT_Transform2D>,
    after: Option<&oxml_drawing::xfrm::CT_Transform2D>,
) -> bool {
    let Some(before) = before else {
        return false;
    };
    let position = |transform: &oxml_drawing::xfrm::CT_Transform2D| {
        transform.offset.as_ref().map(|offset| (offset.x, offset.y))
    };
    let size = |transform: &oxml_drawing::xfrm::CT_Transform2D| {
        transform
            .extent
            .as_ref()
            .map(|extent| (extent.cx, extent.cy))
    };
    if shape.position() != position(before)
        || shape.size() != size(before)
        || shape.rotation() != Some(before.rotation)
        || shape.flips() != Some((before.flip_horizontal, before.flip_vertical))
    {
        return false;
    }
    let Some(after) = after else {
        return position(before).is_some()
            || size(before).is_some()
            || before.rotation != rpptx::Angle(0)
            || before.flip_horizontal
            || before.flip_vertical;
    };
    position(before) != position(after)
        || size(before) != size(after)
        || before.rotation != after.rotation
        || before.flip_horizontal != after.flip_horizontal
        || before.flip_vertical != after.flip_vertical
}

fn prove_local_shape_xml(
    inner: &rpptx::Presentation,
    path: &oxml_py_support::ContentPath,
    before: &[u8],
    after: &[u8],
) -> PyResult<bool> {
    use rpptx_oxml::shape_tree::{CT_Shape, ShapeTreeChild};
    // Template writers also carry dirty-part bookkeeping. Keep those uncertain
    // native contexts on a candidate, including native rejection paths.
    if !matches!(
        crate::shape::shape_owner(path)?,
        rpptx::ShapeOwner::Slide(_)
    ) {
        return Ok(false);
    }
    let shape = crate::shape::native_shape_ref_at(inner, path)
        .ok_or_else(|| PyValueError::new_err("shape owner is absent"))?;
    let actual = shape_structure(inner, path)?;
    match actual.kind {
        rpptx::ShapeKind::Shape => {
            let (Ok(old), Ok(new)) = (CT_Shape::from_xml(before), CT_Shape::from_xml(after)) else {
                return Ok(false);
            };
            let old_runs = projected_runs(old.text_body.as_ref());
            let new_runs = projected_runs(new.text_body.as_ref());
            let old_alt = old.alt_text();
            let new_alt = new.alt_text();
            let geometry = geometry_witness(
                &shape,
                old.shape_properties.transform.as_ref(),
                new.shape_properties.transform.as_ref(),
            );
            let old = ShapeTreeChild::Shape(old);
            let new = ShapeTreeChild::Shape(new);
            if actual.id != old.non_visual_id()
                || actual.id != new.non_visual_id()
                || actual.has_table
                || actual.paragraph_runs != old_runs
                || actual.paragraph_runs != new_runs
            {
                return Ok(false);
            }
            let name = shape.non_visual_name() == old.non_visual_name()
                && old.non_visual_name() != new.non_visual_name();
            let alt = shape.alt_text() == old_alt && old_alt != new_alt;
            Ok(name || alt || geometry)
        }
        rpptx::ShapeKind::GraphicFrame => {
            use rpptx_oxml::graphic_frame::{CT_GraphicFrame, GraphicDataPayload};
            let (Ok(old), Ok(new)) = (
                CT_GraphicFrame::from_xml(before),
                CT_GraphicFrame::from_xml(after),
            ) else {
                return Ok(false);
            };
            let old_table = matches!(old.graphic_data.payload(), GraphicDataPayload::Table(_));
            let new_table = matches!(new.graphic_data.payload(), GraphicDataPayload::Table(_));
            let payload_same = old.graphic_data == new.graphic_data;
            let geometry = geometry_witness(&shape, Some(&old.transform), Some(&new.transform));
            let old = ShapeTreeChild::GraphicFrame(Box::new(old));
            let new = ShapeTreeChild::GraphicFrame(Box::new(new));
            if actual.id != old.non_visual_id() || actual.has_table != old_table {
                return Ok(false);
            }
            if old_table != new_table {
                return Err(pyo3::exceptions::PyNotImplementedError::new_err(
                    "Shape.element.replace_xml does not support changing table payload presence",
                ));
            }
            if actual.id != new.non_visual_id() || !payload_same {
                return Ok(false);
            }
            Ok((shape.non_visual_name() == old.non_visual_name()
                && old.non_visual_name() != new.non_visual_name())
                || geometry)
        }
        _ => Ok(false),
    }
}

// Existing presentation.rs helpers. No current-owner history is retained.
impl PyPresentation {
    fn replace_text_candidate(
        &mut self,
        py: Python<'_>,
        placeholder: &str,
        replacement: &str,
        expect: Option<usize>,
        operation: &'static str,
    ) -> PyResult<usize> {
        let mut candidate = self.inner.clone();
        let count = py
            .detach(|| candidate.try_replace_text(placeholder, replacement))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        if let Some(expected) = expect.filter(|expected| *expected != count) {
            return Err(replacement_count_to_pyerr(py, placeholder, expected, count));
        }
        if count > 0 {
            let survivors = self.replaced_text_owners(&candidate)?;
            let prepared =
                self.prepare_owners(&candidate, operation, &survivors, false, false, false)?;
            self.publish_owners(candidate, prepared)?;
        }
        Ok(count)
    }

    /// Footer staging only removes owned footer shapes, appends fresh copies,
    /// and preserves every surviving physical identity. Match that explicit
    /// native guarantee and compare complete text owners before publication.
    ///
    /// A part the candidate still shares keeps its owners unchanged, so only
    /// the other parts are walked and matched. A physical identity names its
    /// part, and a slide ID is held by one slide, so an owner of a matched
    /// part can only match within that part, and gets the proof that
    /// matching every part gives. A kept part is not matched at all, so a
    /// physical identity two of its owners share, such as a repeated shape
    /// ID or two `mc:AlternateContent` shapes, no longer refuses the edit as
    /// ambiguous. Debug builds check both against matching every part.
    pub(crate) fn footer_owners(&self, candidate: &rpptx::Presentation) -> PyResult<OwnerProofs> {
        let carried = self.carried_parts(candidate, &OwnerProofs::edit(PartEdit::Unchanged));
        let kept = carried.kept_live_slides(self.owners.parts.slides.len());
        let changed = self
            .prefix_parts(&[])
            .into_iter()
            .filter(|part| match part {
                OwnerPart::Slide(index) => !kept[*index],
                OwnerPart::Masters => !carried.masters,
            })
            .collect::<Vec<_>>();
        let mut next = Vec::new();
        for (at, origin) in carried.slides.iter().enumerate() {
            if origin.is_none() {
                slide_owner_addresses(candidate, at, &mut next)?;
            }
        }
        if !carried.masters {
            master_owner_addresses(candidate, &mut next)?;
        }
        let proofs = self.matched_footer_owners(candidate, changed.clone(), &next);
        #[cfg(debug_assertions)]
        without_work_count(|| self.assert_footer_owners(candidate, &carried, &changed, &proofs))?;
        proofs
    }

    /// Debug builds check footer proofs against matching every part: when
    /// that succeeds the proofs are the same, those of a kept part being its
    /// unchanged ones, and when it refuses, matching the kept parts alone
    /// refuses too, the one difference footer_owners states.
    #[cfg(debug_assertions)]
    fn assert_footer_owners(
        &self,
        candidate: &rpptx::Presentation,
        carried: &CarriedParts,
        changed: &[OwnerPart],
        proofs: &PyResult<OwnerProofs>,
    ) -> PyResult<()> {
        let mut every = Vec::new();
        let mut kept_next = Vec::new();
        for (at, origin) in carried.slides.iter().enumerate() {
            slide_owner_addresses(candidate, at, &mut every)?;
            if origin.is_some() {
                slide_owner_addresses(candidate, at, &mut kept_next)?;
            }
        }
        master_owner_addresses(candidate, &mut every)?;
        if carried.masters {
            master_owner_addresses(candidate, &mut kept_next)?;
        }
        let whole = self.matched_footer_owners(candidate, self.prefix_parts(&[]), &every);
        match (whole, proofs) {
            (Ok(whole), Ok(proofs)) => {
                for (part, list) in &whole.parts {
                    match proofs.parts.iter().find(|(listed, _)| listed == part) {
                        Some((_, listed)) => assert_eq!(listed, list, "footer proofs of {part:?}"),
                        None => {
                            let owners = self.owners.parts.part(*part).expect("live part");
                            let unchanged = owners
                                .owners
                                .iter()
                                .filter_map(|owner| PartEdit::Unchanged.proof(owner))
                                .collect::<Vec<_>>();
                            assert_eq!(&unchanged, list, "footer proofs of kept {part:?}");
                        }
                    }
                }
            }
            (Err(whole), Ok(_)) => {
                let kept_parts = self
                    .prefix_parts(&[])
                    .into_iter()
                    .filter(|part| !changed.contains(part))
                    .collect();
                assert!(
                    self.matched_footer_owners(candidate, kept_parts, &kept_next)
                        .is_err(),
                    "footer proofs accept what matching every part refuses: {whole}"
                );
            }
            (Ok(_), Err(error)) => {
                panic!("footer proofs refuse what matching every part accepts: {error}")
            }
            (Err(whole), Err(error)) => assert_eq!(whole.to_string(), error.to_string()),
        }
        Ok(())
    }

    /// Survival proofs for the owners of `parts`, each matched by physical
    /// identity to its one address in `next`.
    fn matched_footer_owners(
        &self,
        candidate: &rpptx::Presentation,
        parts: Vec<OwnerPart>,
        next: &[OwnerAddress],
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::ContentPath;
        let mut by_physical =
            HashMap::<&PhysicalOwner, smallvec::SmallVec<[&OwnerAddress; 1]>>::new();
        by_physical.try_reserve(next.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("footer owner allocation failed")
        })?;
        for address in next {
            by_physical
                .entry(&address.physical)
                .or_default()
                .push(address);
        }
        let mut changed_bodies = Vec::<smallvec::SmallVec<[PathSeg; 5]>>::new();
        self.part_owner_proofs(parts, |old, proof| {
            // The native footer provider retains the body but replaces its
            // text children. Earlier nodes do not survive by ordinal.
            if matches!(
                old.address.physical,
                PhysicalOwner::Paragraph { .. } | PhysicalOwner::Run { .. }
            ) && changed_bodies
                .iter()
                .any(|base| old.address.path.starts_with(base))
            {
                return Ok(false);
            }
            let after = match by_physical
                .get(&old.address.physical)
                .map(|found| found.as_slice())
            {
                None | Some([]) => return Ok(false),
                Some([after]) => *after,
                Some(_) => {
                    return Err(PyValueError::new_err("footer owner identity is ambiguous"));
                }
            };
            let before_path = ContentPath::new(old.address.path.clone(), self.revisions.current());
            let after_path = ContentPath::new(after.path.clone(), self.revisions.current());
            let text_changed = match &old.address.physical {
                PhysicalOwner::Shape { .. } => {
                    crate::shape::native_shape_ref_at(&self.inner, &before_path)
                        .and_then(|shape| shape.text_frame())
                        != crate::shape::native_shape_ref_at(candidate, &after_path)
                            .and_then(|shape| shape.text_frame())
                }
                PhysicalOwner::TextBody { cell, .. } => {
                    let frame = |inner, path| {
                        let shape = crate::shape::native_shape_ref_at(inner, path)?;
                        match cell {
                            Some((row, col)) => shape.table()?.cell(*row, *col)?.text_frame(),
                            None => shape.text_frame(),
                        }
                    };
                    frame(&self.inner, &before_path) != frame(candidate, &after_path)
                }
                // Footer publication does not change notes or any table cell.
                PhysicalOwner::Slide(_)
                | PhysicalOwner::Cell { .. }
                | PhysicalOwner::Layout(_)
                | PhysicalOwner::Master(_)
                | PhysicalOwner::Paragraph { .. }
                | PhysicalOwner::Run { .. } => false,
            };
            if text_changed && matches!(old.address.physical, PhysicalOwner::TextBody { .. }) {
                changed_bodies.push(old.address.path.clone());
            }
            proof.destination = after.path.clone();
            proof.layout_index = after.layout_index;
            proof.text_changed = text_changed;
            Ok(true)
        })
    }

    /// The parts holding the owners under a path prefix, every part for an
    /// empty prefix.
    fn prefix_parts(&self, prefix: &[PathSeg]) -> Vec<OwnerPart> {
        if prefix.is_empty() {
            self.owners.parts.iter().map(|(part, _)| part).collect()
        } else {
            vec![OwnerPart::of(prefix)]
        }
    }

    /// Explicit proofs for every owner of `parts`, each starting as the
    /// owner's unchanged proof. `edit` adjusts one proof, or returns false to
    /// retire its owner. Owners of other parts survive unchanged.
    fn part_owner_proofs(
        &self,
        parts: Vec<OwnerPart>,
        mut edit: impl FnMut(&LiveOwner, &mut OwnerSurvival) -> PyResult<bool>,
    ) -> PyResult<OwnerProofs> {
        let mut listed = Vec::new();
        for part in parts {
            let Some(owners) = self.owners.parts.part(part) else {
                continue;
            };
            let mut proofs = Vec::new();
            proofs.try_reserve_exact(owners.owners.len()).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("owner proofs allocation failed")
            })?;
            for owner in &owners.owners {
                let mut proof = PartEdit::Unchanged
                    .proof(owner)
                    .expect("an unchanged owner survives");
                if edit(owner, &mut proof)? {
                    proofs.push(proof);
                }
            }
            listed.push((part, proofs));
        }
        Ok(OwnerProofs {
            edit: PartEdit::Unchanged,
            parts: listed,
        })
    }

    pub(crate) fn unchanged_owners(&self) -> PyResult<OwnerProofs> {
        Ok(OwnerProofs::edit(PartEdit::Unchanged))
    }

    /// Retires one removed layout and its descendants, preserving exact surviving part owners.
    pub(crate) fn removed_layout_owners(&self, removed: usize) -> PyResult<OwnerProofs> {
        Ok(OwnerProofs::edit(PartEdit::RemoveLayout(removed)))
    }

    /// Remap exactly one existing sibling ordinal. The operation must use this
    /// same permutation on its staged native tree before preparing publication.
    pub(crate) fn moved_shape_owners(
        &self,
        parent: &[oxml_py_support::PathSeg],
        from: usize,
        to: usize,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        self.part_owner_proofs(self.prefix_parts(parent), |_, proof| {
            if proof.layout_index.is_some() || !proof.destination.starts_with(parent) {
                return Ok(true);
            }
            let Some(PathSeg::Shape(index)) = proof.destination.get_mut(parent.len()) else {
                return Ok(true);
            };
            *index = if *index == from {
                to
            } else if from < to && *index > from && *index <= to {
                *index - 1
            } else if to < from && *index >= to && *index < from {
                *index + 1
            } else {
                *index
            };
            Ok(true)
        })
    }

    /// Retires one shape replaced in place with every owner under it.
    pub(crate) fn replaced_shape_owners(
        &self,
        shape: &[oxml_py_support::PathSeg],
    ) -> PyResult<OwnerProofs> {
        self.part_owner_proofs(self.prefix_parts(shape), |_, proof| {
            Ok(proof.layout_index.is_some() || !proof.destination.starts_with(shape))
        })
    }

    pub(crate) fn removed_shape_owners(
        &self,
        parent: &[oxml_py_support::PathSeg],
        removed: usize,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        self.part_owner_proofs(self.prefix_parts(parent), |_, proof| {
            if proof.layout_index.is_some() || !proof.destination.starts_with(parent) {
                return Ok(true);
            }
            let Some(PathSeg::Shape(index)) = proof.destination.get_mut(parent.len()) else {
                return Ok(true);
            };
            if *index == removed {
                return Ok(false);
            }
            if *index > removed {
                *index -= 1;
            }
            Ok(true)
        })
    }

    /// Apply the same explicit ordinal operation to every surviving physical
    /// slide descendant. Removed members never contribute a survival proof.
    pub(crate) fn slide_owner_edit(&self, edit: SlideOwnerEdit) -> PyResult<OwnerProofs> {
        Ok(OwnerProofs::edit(PartEdit::Slides(edit)))
    }

    /// Literal edits preserve bodies but retire the actual old text nodes in
    /// each changed body. Matching ordinals are not survival evidence. A part
    /// whose native records the candidate still shares has no changed text.
    pub(crate) fn replaced_text_owners(
        &self,
        candidate: &rpptx::Presentation,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::ContentPath;
        let origins = candidate.unchanged_slides_since(&self.inner);
        let masters = candidate.unchanged_layouts_and_masters_since(&self.inner);
        let parts = self
            .owners
            .parts
            .iter()
            .map(|(part, _)| part)
            .filter(|part| match part {
                OwnerPart::Slide(index) => origins.get(*index) != Some(&Some(*index)),
                OwnerPart::Masters => !masters,
            })
            .collect();
        let mut proofs = self.part_owner_proofs(parts, |owner, proof| {
            let path = ContentPath::new(owner.address.path.clone(), self.revisions.current());
            proof.text_changed = match &owner.address.physical {
                PhysicalOwner::Slide(_) => {
                    let index = crate::shape::slide_index(&path)?;
                    match (self.inner.slide(index), candidate.slide(index)) {
                        (Some(before), Some(after)) => !before.notes_text_matches(&after),
                        (None, None) => false,
                        _ => true,
                    }
                }
                PhysicalOwner::TextBody { cell, .. } => {
                    let frame = |inner| {
                        let shape = crate::shape::native_shape_ref_at(inner, &path)?;
                        match cell {
                            Some((row, col)) => shape.table()?.cell(*row, *col)?.text_frame(),
                            None => shape.text_frame(),
                        }
                    };
                    frame(&self.inner) != frame(candidate)
                }
                _ => false,
            };
            Ok(true)
        })?;
        // A changed body's paragraphs and runs share its part.
        for (part, list) in &mut proofs.parts {
            let owners = &self
                .owners
                .parts
                .part(*part)
                .expect("listed live part")
                .owners;
            let changed = owners
                .iter()
                .zip(list.iter())
                .filter(|(owner, proof)| {
                    proof.text_changed
                        && matches!(owner.address.physical, PhysicalOwner::TextBody { .. })
                })
                .map(|(owner, _)| owner.address.path.clone())
                .collect::<Vec<_>>();
            let mut owners = owners.iter();
            list.retain(|_| {
                let owner = owners.next().expect("one proof per owner");
                !matches!(
                    owner.address.physical,
                    PhysicalOwner::Paragraph { .. } | PhysicalOwner::Run { .. }
                ) || !changed
                    .iter()
                    .any(|base| owner.address.path.starts_with(base))
            });
        }
        Ok(proofs)
    }

    pub(crate) fn cell_text_owners(
        &self,
        path: &oxml_py_support::ContentPath,
        first_row: usize,
        last_row: usize,
        first_column: usize,
        last_column: usize,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        let table_path = path
            .segs
            .iter()
            .filter(|seg| {
                matches!(
                    seg,
                    PathSeg::Slide(_) | PathSeg::Layout(_) | PathSeg::Master(_) | PathSeg::Shape(_)
                )
            })
            .copied()
            .collect::<smallvec::SmallVec<[PathSeg; 5]>>();
        self.part_owner_proofs(self.prefix_parts(&table_path), |owner, proof| {
            if owner.address.layout_index.is_some() || !owner.address.path.starts_with(&table_path)
            {
                return Ok(true);
            }
            let cell = match owner.address.physical {
                PhysicalOwner::Cell { row, column, .. } => Some((row, column)),
                PhysicalOwner::TextBody { cell, .. }
                | PhysicalOwner::Paragraph { cell, .. }
                | PhysicalOwner::Run { cell, .. } => cell,
                _ => None,
            };
            let affected = cell.is_some_and(|(row, column)| {
                row >= first_row
                    && row <= last_row
                    && column >= first_column
                    && column <= last_column
            });
            if affected
                && matches!(
                    owner.address.physical,
                    PhysicalOwner::Paragraph { .. } | PhysicalOwner::Run { .. }
                )
            {
                return Ok(false);
            }
            proof.text_changed = affected;
            Ok(true)
        })
    }

    /// Keep earlier paragraph and run owners. Retire only the appended
    /// owner's cached text view, so a materialised collection is not refreshed.
    pub(crate) fn appended_text_owners(
        &self,
        path: &oxml_py_support::ContentPath,
    ) -> PyResult<OwnerProofs> {
        let generation = path
            .owner_generation()
            .ok_or_else(|| PyValueError::new_err("text append needs a bound owner"))?;
        let (part, _) = *self
            .owners
            .by_generation
            .get(&generation)
            .ok_or_else(|| PyValueError::new_err("text append owner is absent"))?;
        self.part_owner_proofs(vec![part], |_, proof| {
            proof.text_changed |= proof.generation == generation;
            Ok(true)
        })
    }

    /// Text replacement retains its physical paragraph and body ancestors,
    /// while the replaced runs and whole-body paragraph descendants retire.
    pub(crate) fn text_changed_owners(
        &self,
        path: &oxml_py_support::ContentPath,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        let current = self
            .current_path(path)
            .ok_or_else(|| PyValueError::new_err("text mutation owner is absent"))?;
        let generation = path
            .owner_generation()
            .ok_or_else(|| PyValueError::new_err("text mutation needs a bound owner"))?;
        let paragraph = current
            .segs
            .iter()
            .any(|seg| matches!(seg, PathSeg::Para(_)));
        let base = current
            .segs
            .iter()
            .filter(|seg| !matches!(seg, PathSeg::Para(_) | PathSeg::Run(_)))
            .copied()
            .collect::<smallvec::SmallVec<[PathSeg; 5]>>();
        let body = self
            .body_owner_at(&base)
            .map(|owner| owner.owner.generation());
        let owner = self.owners.by_generation.get(&generation);
        // The slide scalar setter edits only the notes part. Slide body
        // owners are unrelated and cannot be retired by that operation.
        if let Some((part, index)) = owner
            && matches!(
                self.owners
                    .parts
                    .part(*part)
                    .map(|owners| &owners.owners[*index].address.physical),
                Some(PhysicalOwner::Slide(_))
            )
        {
            return self.part_owner_proofs(vec![*part], |_, proof| {
                proof.text_changed |= proof.generation == generation;
                Ok(true)
            });
        }
        let mut parts = self.prefix_parts(&base);
        if let Some((part, _)) = owner
            && !parts.contains(part)
        {
            parts.push(*part);
        }
        self.part_owner_proofs(parts, |old, proof| {
            let removed = if paragraph {
                matches!(old.address.physical, PhysicalOwner::Run { .. })
                    && old.address.path.starts_with(&current.segs)
            } else {
                matches!(
                    old.address.physical,
                    PhysicalOwner::Paragraph { .. } | PhysicalOwner::Run { .. }
                ) && old.address.path.starts_with(&base)
            };
            if removed {
                return Ok(false);
            }
            if proof.generation == generation || !paragraph && Some(proof.generation) == body {
                proof.text_changed = true;
            }
            Ok(true)
        })
    }

    pub(crate) fn grid_changed_owners(
        &self,
        path: &oxml_py_support::ContentPath,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        let current = self
            .current_path(path)
            .ok_or_else(|| pyo3::exceptions::PyValueError::new_err("grid owner is absent"))?;
        let table_path = current
            .segs
            .iter()
            .filter(|seg| {
                matches!(
                    seg,
                    PathSeg::Slide(_) | PathSeg::Layout(_) | PathSeg::Master(_) | PathSeg::Shape(_)
                )
            })
            .copied()
            .collect::<smallvec::SmallVec<[PathSeg; 5]>>();
        let table = self
            .owner_at(&table_path, None)
            .ok_or_else(|| pyo3::exceptions::PyValueError::new_err("table owner is absent"))?
            .owner
            .generation();
        // Every old physical cell retires on grid restructure. The table
        // survives with a fresh grid generation, and new cells get births.
        self.part_owner_proofs(self.prefix_parts(&table_path), |_, proof| {
            if proof.destination.starts_with(&table_path)
                && proof
                    .destination
                    .get(table_path.len())
                    .is_some_and(|seg| matches!(seg, PathSeg::Row(_)))
            {
                return Ok(false);
            }
            proof.grid_changed |= proof.generation == table;
            Ok(true)
        })
    }
}

// Replacement of existing shape.rs move_shape body AFTER range validation.
// Concrete source insertion, not a forwarding wrapper.
// let parent = smallvec![PathSeg::Slide(slide_index)];
// let mut presentation = self.presentation.borrow_mut(py);
// let mut candidate = presentation.inner.clone();
// candidate.move_shape(slide_index, from_, to).map_err(|e| rpptx_to_pyerr(py, e))?;
// let proofs = presentation.moved_shape_owners(&parent, from_, to)?;
// let prepared = presentation.prepare_owners(&candidate, "ShapeCollection.move", &proofs, false, false, false)?;
// presentation.publish_owners(candidate, prepared)
// Preserve the existing native no-op definition for from_==to before serial prep.

// Removal cannot use remove_shape alone as final acceptance. It must first
// detach incoming endpoints and preserve their last resolved geometry in this
// SAME candidate. The final authored removal method returns only after both
// detachment and native structural removal succeed. The owner proof excludes
// the removed subtree and shifts surviving ordinals. Publication is one serial.

// Candidate additions inside existing presentation.rs impl, not a new module.
impl PyPresentation {
    pub(crate) fn grouped_shape_owners(
        &self,
        parent: &[oxml_py_support::PathSeg],
        start: usize,
        count: usize,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        let end = start
            .checked_add(count)
            .ok_or_else(|| pyo3::exceptions::PyOverflowError::new_err("group range overflow"))?;
        if count == 0 {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "empty group selection",
            ));
        }
        self.part_owner_proofs(self.prefix_parts(parent), |_, proof| {
            if proof.layout_index.is_some() || !proof.destination.starts_with(parent) {
                return Ok(true);
            }
            let Some(PathSeg::Shape(index)) = proof.destination.get(parent.len()).copied() else {
                return Ok(true);
            };
            if index >= start && index < end {
                proof.destination.try_reserve(1).map_err(|_| {
                    pyo3::exceptions::PyMemoryError::new_err("group path allocation failed")
                })?;
                proof.destination[parent.len()] = PathSeg::Shape(start);
                proof
                    .destination
                    .insert(parent.len() + 1, PathSeg::Shape(index - start));
            } else if index >= end {
                proof.destination[parent.len()] = PathSeg::Shape(index - (count - 1));
            }
            Ok(true)
        })
    }
    pub(crate) fn ungrouped_shape_owners(
        &self,
        parent: &[oxml_py_support::PathSeg],
        index: usize,
        count: usize,
    ) -> PyResult<OwnerProofs> {
        use oxml_py_support::PathSeg;
        let group = |destination: &[PathSeg]| {
            destination.starts_with(parent)
                && destination.len() == parent.len() + 1
                && destination[parent.len()] == PathSeg::Shape(index)
        };
        // Retire the group first, then remap, so that a remap error names
        // the first proof the original two passes would have rejected.
        let mut proofs = self.part_owner_proofs(self.prefix_parts(parent), |_, proof| {
            Ok(!group(&proof.destination))
        })?;
        for (_, list) in &mut proofs.parts {
            for proof in list {
                if proof.layout_index.is_some() || !proof.destination.starts_with(parent) {
                    continue;
                }
                let Some(PathSeg::Shape(at)) = proof.destination.get(parent.len()).copied() else {
                    continue;
                };
                if at == index {
                    let Some(PathSeg::Shape(child)) =
                        proof.destination.get(parent.len() + 1).copied()
                    else {
                        return Err(pyo3::exceptions::PyValueError::new_err(
                            "invalid group child proof",
                        ));
                    };
                    if child >= count {
                        return Err(pyo3::exceptions::PyValueError::new_err(
                            "group child proof out of range",
                        ));
                    }
                    proof.destination[parent.len()] =
                        PathSeg::Shape(index.checked_add(child).ok_or_else(|| {
                            pyo3::exceptions::PyOverflowError::new_err("ungroup index overflow")
                        })?);
                    proof.destination.remove(parent.len() + 1);
                } else if at > index {
                    proof.destination[parent.len()] = PathSeg::Shape(
                        at.checked_sub(1)
                            .and_then(|v| v.checked_add(count))
                            .ok_or_else(|| {
                                pyo3::exceptions::PyOverflowError::new_err("ungroup index overflow")
                            })?,
                    );
                }
            }
        }
        Ok(proofs)
    }
}

pub(crate) fn collect_caller_fonts<'py>(
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
        .map_err(|error| rpptx_value_to_pyerr(py, error.to_string()))?;
    Ok((!selected.is_empty()).then_some(selected))
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

fn font_kind_name(kind: rpptx::EmbeddedFontKind) -> &'static str {
    match kind {
        rpptx::EmbeddedFontKind::Regular => "regular",
        rpptx::EmbeddedFontKind::Bold => "bold",
        rpptx::EmbeddedFontKind::Italic => "italic",
        rpptx::EmbeddedFontKind::BoldItalic => "bold_italic",
    }
}
fn checked_font_kind(kind: &str) -> PyResult<rpptx::EmbeddedFontKind> {
    match kind {
        "regular" => Ok(rpptx::EmbeddedFontKind::Regular),
        "bold" => Ok(rpptx::EmbeddedFontKind::Bold),
        "italic" => Ok(rpptx::EmbeddedFontKind::Italic),
        "bold_italic" => Ok(rpptx::EmbeddedFontKind::BoldItalic),
        _ => Err(PyValueError::new_err(
            "font kind must be regular, bold, italic or bold_italic",
        )),
    }
}
fn checked_font_license(license: &PyFontEmbeddingLicense) -> PyResult<rpptx::FontEmbeddingLicense> {
    if !license.authorized || license.identity.trim().is_empty() {
        return Err(PyValueError::new_err(
            "font embedding requires explicit authorization and nonempty license identity",
        ));
    }
    Ok(rpptx::FontEmbeddingLicense::new(
        license.authorized,
        license.identity.clone(),
    ))
}

#[pyclass(name = "EmbeddedFont", frozen, get_all, skip_from_py_object)]
pub struct PyEmbeddedFont {
    pub kind: &'static str,
    pub data: Py<PyBytes>,
    pub license: Py<PyFontEmbeddingLicense>,
}
#[pymethods]
impl PyEmbeddedFont {
    #[new]
    #[pyo3(signature = (font, kind="regular", *, license))]
    fn new(
        py: Python<'_>,
        font: &Bound<'_, PyAny>,
        kind: &str,
        license: PyRef<'_, PyFontEmbeddingLicense>,
    ) -> PyResult<Self> {
        let kind = checked_font_kind(kind)?;
        checked_font_license(&license)?;
        let bytes = checked_font_input(font)?;
        let data = PyBytes::new_with(py, bytes.len(), |out| {
            out.copy_from_slice(&bytes);
            Ok(())
        })?
        .unbind();
        let license = Py::new(py, license.clone())?;
        Ok(Self {
            kind: font_kind_name(kind),
            data,
            license,
        })
    }
}
#[pyclass(name = "EmbeddedFontFace", frozen, get_all, skip_from_py_object)]
pub struct PyEmbeddedFontFace {
    pub kind: &'static str,
    pub relationship_id: String,
    pub part_name: Option<String>,
    pub format: &'static str,
    pub stored_len: usize,
    pub data: Option<Py<PyBytes>>,
}
impl PyEmbeddedFontFace {
    fn snapshot(
        py: Python<'_>,
        font: rpptx::EmbeddedFontFace,
        bodies: &mut std::collections::BTreeMap<String, Py<PyBytes>>,
    ) -> PyResult<Py<Self>> {
        let data = if let Some(bytes) = &font.data {
            let part = font
                .part_name
                .as_ref()
                .ok_or_else(|| PyValueError::new_err("decoded font has no source part"))?;
            if let Some(bytes) = bodies.get(part) {
                Some(bytes.clone_ref(py))
            } else {
                let body = PyBytes::new_with(py, bytes.len(), |out| {
                    out.copy_from_slice(bytes);
                    Ok(())
                })?
                .unbind();
                bodies.insert(part.clone(), body.clone_ref(py));
                Some(body)
            }
        } else {
            None
        };
        let format = match font.format {
            rpptx::EmbeddedFontFormat::Eot => "eot",
            rpptx::EmbeddedFontFormat::CompressedEot => "compressed_eot",
            rpptx::EmbeddedFontFormat::Sfnt => "sfnt",
            rpptx::EmbeddedFontFormat::Unknown => "unknown",
        };
        Py::new(
            py,
            Self {
                kind: font_kind_name(font.kind),
                relationship_id: font.relationship_id,
                part_name: font.part_name,
                format,
                stored_len: font.stored_len,
                data,
            },
        )
    }
}
#[pyclass(name = "EmbeddedTypeface", frozen, get_all, skip_from_py_object)]
pub struct PyEmbeddedTypeface {
    pub typeface: String,
    pub panose: Option<String>,
    pub pitch_family: Option<i8>,
    pub charset: Option<i8>,
    pub faces: Py<PyTuple>,
}
impl PyEmbeddedTypeface {
    fn snapshot(
        py: Python<'_>,
        font: rpptx::EmbeddedTypeface,
        bodies: &mut std::collections::BTreeMap<String, Py<PyBytes>>,
    ) -> PyResult<Py<Self>> {
        let mut faces = Vec::new();
        faces.try_reserve_exact(font.faces.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("font snapshot allocation failed")
        })?;
        for face in font.faces {
            faces.push(PyEmbeddedFontFace::snapshot(py, face, bodies)?);
        }
        let faces = PyTuple::new(py, faces)?.unbind();
        Py::new(
            py,
            Self {
                typeface: font.typeface,
                panose: font.panose,
                pitch_family: font.pitch_family,
                charset: font.charset,
                faces,
            },
        )
    }
}
impl PyPresentation {
    fn publish_font_candidate(
        &mut self,
        py: Python<'_>,
        candidate: rpptx::Presentation,
        operation: &'static str,
    ) -> PyResult<()> {
        if self
            .inner
            .saves_as(&candidate)
            .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let survivors = self.unchanged_owners()?;
        let prepared =
            self.prepare_owners(&candidate, operation, &survivors, false, false, false)?;
        self.publish_owners(candidate, prepared)
    }
}

#[cfg(test)]
mod fx260_visit_tests {
    use super::*;
    use oxml_py_support::PathSeg;

    /// Edits on the first slide walk its part and the layouts and masters,
    /// and match their proofs, whatever the number of later slides.
    #[test]
    fn issue403_shape_edits_do_the_same_inventory_work_at_n_and_4n_slides() {
        Python::initialize();
        Python::attach(|py| -> PyResult<()> {
            let mut expected_work = None;
            for slides in [3usize, 12] {
                let mut native = rpptx::Presentation::new().unwrap();
                for slide in 0..slides {
                    native.add_slide(6).unwrap();
                    let mut shapes = native.shapes_mut(slide, &[]).unwrap();
                    for _ in 0..6 {
                        shapes
                            .add_textbox(
                                rpptx::Emu(0),
                                rpptx::Emu(0),
                                rpptx::Emu(100),
                                rpptx::Emu(100),
                            )
                            .unwrap();
                    }
                    shapes
                        .add_table(
                            3,
                            3,
                            rpptx::Emu(0),
                            rpptx::Emu(0),
                            rpptx::Emu(900),
                            rpptx::Emu(900),
                        )
                        .unwrap();
                }
                let presentation = Py::new(py, PyPresentation::from_presentation_owned(native)?)?;
                let deck = presentation.bind(py);
                let shapes = deck.getattr("slides")?.get_item(0)?.getattr("shapes")?;
                for index in 0..6 {
                    shapes.get_item(index)?.setattr("text", "before")?;
                }
                let mut work = Vec::new();
                let mut count = |label: &'static str, edit: &dyn Fn() -> PyResult<()>| {
                    reset_fx260_ppt_visits();
                    edit()?;
                    work.push((label, fx260_ppt_visits()));
                    PyResult::Ok(())
                };
                count("ShapeCollection.add_shape", &|| {
                    shapes.call_method1("add_shape", ("rect", 0, 0, 100, 100))?;
                    Ok(())
                })?;
                count("Paragraph.add_run", &|| {
                    let paragraph = shapes
                        .get_item(0)?
                        .getattr("text_frame")?
                        .getattr("paragraphs")?
                        .get_item(0)?;
                    paragraph.call_method1("add_run", ("run",))?;
                    Ok(())
                })?;
                count("Shape.text", &|| {
                    shapes.get_item(1)?.setattr("text", "after")
                })?;
                count("Cell.text", &|| {
                    shapes
                        .get_item(6)?
                        .getattr("table")?
                        .call_method1("cell", (1, 1))?
                        .setattr("text", "cell")
                })?;
                count("Element.replace_xml (shape text)", &|| {
                    let element = shapes.get_item(2)?.getattr("element")?;
                    let xml = element.getattr("xml")?.extract::<String>()?;
                    assert!(xml.contains(">before<"), "{xml}");
                    element.call_method1("replace_xml", (xml.replace(">before<", ">after<"),))?;
                    Ok(())
                })?;
                count("SlideCollection.add_slide", &|| {
                    let layout = deck.getattr("slide_layouts")?.get_item(6)?;
                    deck.getattr("slides")?
                        .call_method1("add_slide", (layout,))?;
                    Ok(())
                })?;
                assert!(work.iter().all(|(_, visits)| *visits > 0), "{work:?}");
                if let Some(expected) = &expected_work {
                    assert_eq!(&work, expected, "later slides changed the inventory work");
                } else {
                    expected_work = Some(work);
                }
            }
            Ok(())
        })
        .unwrap();
    }

    #[test]
    fn issue389_local_edits_do_not_scale_with_unrelated_package() {
        Python::initialize();
        Python::attach(|py| -> PyResult<()> {
            let mut expected_work = None;
            for (slides, boxes) in [(0usize, 40usize), (40usize, 40usize)] {
                let mut work = Vec::new();
                let mut native = rpptx::Presentation::new().unwrap();
                native.add_slide(6).unwrap();
                for _ in 0..slides {
                    native.add_slide(6).unwrap();
                }
                for slide in 1..=slides {
                    for _ in 0..boxes {
                        native
                            .shapes_mut(slide, &[])
                            .unwrap()
                            .add_textbox(
                                rpptx::Emu(0),
                                rpptx::Emu(0),
                                rpptx::Emu(100),
                                rpptx::Emu(100),
                            )
                            .unwrap();
                    }
                }
                let presentation = Py::new(py, PyPresentation::from_presentation_owned(native)?)?;
                let path = presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Slide(0)])?;
                let mut shapes =
                    crate::shape::PyShapeCollection::new(presentation.clone_ref(py), path);
                reset_fx260_ppt_visits();
                let shape = shapes.add_textbox(py, 0, 0, 100, 100)?;
                work.push(fx260_ppt_visits());
                // Regular runs, breaks and fields must not collapse to one count.
                // This setup stages before the measured scalar edit below.
                let body = r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>regular</a:t></a:r><a:br/><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>2</a:t></a:fld></a:p></p:txBody>"#;
                let element = shape.borrow(py).element(py)?;
                let xml = element.borrow(py).xml(py)?;
                let start = xml.find("<p:txBody").expect("physical text body");
                let end = xml.find("</p:txBody>").expect("text body end") + "</p:txBody>".len();
                let rich = format!("{}{}{}", &xml[..start], body, &xml[end..]);
                element.borrow(py).replace_xml(py, &rich)?;
                reset_fx260_ppt_visits();
                shape.borrow(py).set_alt_text(py, Some("alt"))?;
                work.push(fx260_ppt_visits());
                reset_fx260_ppt_visits();
                shape.borrow(py).set_left(py, 200_000)?;
                work.push(fx260_ppt_visits());
                reset_fx260_ppt_visits();
                let fill = shape.borrow(py).fill(py)?;
                fill.borrow(py).solid(py)?;
                work.push(fx260_ppt_visits());
                reset_fx260_ppt_visits();
                let element = shape.borrow(py).element(py)?;
                let xml = element.borrow(py).xml(py)?;
                assert!(xml.contains("TextBox "), "{xml}");
                let changed = xml.replacen("TextBox ", "TextBax ", 1);
                element.borrow(py).replace_xml(py, &changed)?;
                work.push(fx260_ppt_visits());
                let root_shapes = presentation.bind(py).getattr("slides")?
                    .get_item(0)?.getattr("shapes")?;
                let frame = root_shapes.call_method1("add_table", (2, 2, 0, 0, 1000, 1000))?
                    .extract::<Py<crate::shape::PyShape>>()?;
                frame.bind(py).getattr("table")?.call_method1("cell", (0, 0))?
                    .setattr("text", "payload")?;
                // Appending on the earlier slide makes its owners noncontiguous.
                root_shapes.call_method1("add_textbox", (0, 0, 100, 100))?;
                let held_cell = frame.bind(py).getattr("table")?.call_method1("cell", (0, 0))?;
                reset_fx260_ppt_visits();
                let element = frame.borrow(py).element(py)?;
                let xml = element.borrow(py).xml(py)?;
                let changed = xml.replacen("Table ", "ChangedTable ", 1);
                assert_ne!(xml, changed);
                element.borrow(py).replace_xml(py, &changed)?;
                work.push(fx260_ppt_visits());
                assert!(held_cell.getattr("text").unwrap_err().to_string().contains("stale"));
                assert_eq!(frame.bind(py).getattr("table")?.call_method1("cell", (0, 0))?
                    .getattr("text")?.extract::<String>()?, "payload");
                reset_fx260_ppt_visits();
                let element = frame.borrow(py).element(py)?;
                let xml = element.borrow(py).xml(py)?;
                let changed = xml.replacen(" x=\"0\"", " x=\"123\"", 1);
                assert_ne!(xml, changed);
                element.borrow(py).replace_xml(py, &changed)?;
                work.push(fx260_ppt_visits());
                assert!(
                    work.iter().all(|&visits| visits > 0 && visits <= 128),
                    "{work:?}"
                );
                if let Some(expected) = &expected_work {
                    assert_eq!(&work, expected, "unrelated slides changed traversal work");
                } else {
                    expected_work = Some(work);
                }
            }
            Ok(())
        })
        .unwrap();
    }
}
