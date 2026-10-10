use std::path::PathBuf;

use oxml_py_support::RevisionCounter;
use pyo3::exceptions::{PyOverflowError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyList, PyTuple};
use smallvec::smallvec;

use crate::layout::PyTextFrameLayout;
use crate::shape::length;
use crate::slide::{PySlideCollection, PySlideLayoutCollection};
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
        };
        Self {
            kind,
            message: format!("{issue:?}"),
        }
    }
}

/// One presentation section: a named, contiguous run of slides.
#[pyclass(name = "Section", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PySection {
    /// The braced GUID PowerPoint identifies the section by.
    pub id: String,
    pub name: String,
    /// The zero-based indices of the section's slides, in order.
    pub slide_indices: Vec<usize>,
}

#[pymethods]
impl PySection {
    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!(
            "Section(id={}, name={}, slide_indices={:?})",
            pyo3::types::PyString::new(py, &self.id).repr()?,
            pyo3::types::PyString::new(py, &self.name).repr()?,
            self.slide_indices
        ))
    }
}

/// One ODP conversion diagnostic as a `(path, message)` pair.
type OdpDiagnostics = Vec<(String, String)>;

fn odp_diagnostics(diagnostics: &[rpptx::OdpDiagnostic]) -> OdpDiagnostics {
    diagnostics
        .iter()
        .map(|diagnostic| (diagnostic.path.clone(), diagnostic.message.clone()))
        .collect()
}

/// A section id derived from the section's name and place, so the same
/// sections give the same file.
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

/// Refuses handout rendering for a deck without a handout master, saying
/// how to get one.
fn require_handout_master(py: Python<'_>, presentation: &rpptx::Presentation) -> PyResult<()> {
    if presentation.has_handout_master() {
        return Ok(());
    }
    Err(rpptx_value_to_pyerr(
        py,
        "the presentation has no handout master, which handouts are laid out from. \
         Open the deck in PowerPoint, show View > Handout Master once and save it, \
         or use to_pdf() for one slide per page"
            .to_owned(),
    ))
}

/// The JPEG quality used when the caller names none.
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

type CoreField = fn(&mut rpptx::CoreProperties) -> &mut Option<String>;

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

/// The package core properties (`docProps/core.xml`) under python-pptx's
/// attribute names. Ported from the rdocx binding, to share once a common
/// helper exists.
///
/// Text properties read as an empty string when absent, `revision` as zero,
/// and dates as `None`. Assigning `None` or empty text removes a property.
/// A write replaces the native model and creates the part, its package
/// relationship and its content type when the presentation has none. It
/// changes no content, so handles stay valid.
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
        *presentation.inner.core_properties_mut() = properties;
        Ok(())
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

#[pyclass(name = "Presentation")]
pub struct PyPresentation {
    pub(crate) inner: rpptx::Presentation,
    pub(crate) revisions: RevisionCounter,
}

impl PyPresentation {
    fn from_presentation(inner: rpptx::Presentation) -> Self {
        Self {
            inner,
            revisions: RevisionCounter::new(),
        }
    }

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
                .map(Self::from_presentation)
                .map_err(|error| rpptx_to_pyerr(py, error)),
            None => rpptx::Presentation::new()
                .map(Self::from_presentation)
                .map_err(|error| rpptx_to_pyerr(py, error)),
        }
    }

    #[staticmethod]
    fn from_bytes(bytes: &[u8], py: Python<'_>) -> PyResult<Self> {
        rpptx::Presentation::from_bytes(bytes)
            .map(Self::from_presentation)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    fn save(&self, path: PathBuf, py: Python<'_>) -> PyResult<()> {
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

    fn to_pdf<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        py.detach(|| self.inner.to_pdf_deterministic())
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (slide_index, dpi = 150.0))]
    fn render_slide_to_png<'py>(
        &self,
        py: Python<'py>,
        slide_index: usize,
        dpi: f64,
    ) -> PyResult<Option<Bound<'py, PyBytes>>> {
        py.detach(|| self.inner.slide_png_deterministic(slide_index, dpi))
            .map(|bytes| bytes.map(|bytes| PyBytes::new(py, &bytes)))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (dpi = 150.0))]
    fn render_all_slides<'py>(&self, py: Python<'py>, dpi: f64) -> PyResult<Bound<'py, PyList>> {
        let slides = py
            .detach(|| self.inner.slide_pngs_deterministic(dpi))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, slides.iter().map(|slide| PyBytes::new(py, slide)))
    }

    #[pyo3(signature = (*, width_factor = 1.0))]
    fn text_layout<'py>(
        &self,
        py: Python<'py>,
        width_factor: f64,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let frames = py
            .detach(|| self.inner.text_layout_deterministic(width_factor))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyTuple::new(py, frames.iter().map(PyTextFrameLayout::from))
    }

    fn to_notes_pdf<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        py.detach(|| self.inner.to_notes_pdf_deterministic())
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (dpi = 150.0))]
    fn render_all_notes<'py>(&self, py: Python<'py>, dpi: f64) -> PyResult<Bound<'py, PyList>> {
        let notes = py
            .detach(|| self.inner.notes_page_pngs_deterministic(dpi))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, notes.iter().map(|page| PyBytes::new(py, page)))
    }

    /// Replaces literal text in slides and speaker notes and returns the count.
    ///
    /// With `expect`, the replacement runs on a clone, so a count that
    /// differs raises and leaves the presentation and its revision as they
    /// were. Without it, the staged facade call runs in place. The revision
    /// advances once only when something was replaced.
    #[pyo3(signature = (placeholder, replacement, *, expect = None))]
    fn try_replace_text(
        &mut self,
        py: Python<'_>,
        placeholder: &str,
        replacement: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        let count = match expect {
            None => py
                .detach(|| self.inner.try_replace_text(placeholder, replacement))
                .map_err(|error| rpptx_to_pyerr(py, error))?,
            Some(expected) => {
                let (candidate, count) = py
                    .detach(|| {
                        let mut candidate = self.inner.clone();
                        candidate
                            .try_replace_text(placeholder, replacement)
                            .map(|count| (candidate, count))
                    })
                    .map_err(|error| rpptx_to_pyerr(py, error))?;
                if count != expected {
                    return Err(replacement_count_to_pyerr(py, placeholder, expected, count));
                }
                self.inner = candidate;
                count
            }
        };
        if count > 0 {
            self.revisions.bump();
        }
        Ok(count)
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
        self.try_replace_text(py, placeholder, replacement, expect)
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
        self.inner
            .add_comment_author(author)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.revisions.bump();
        Ok(())
    }

    #[getter]
    fn slide_layouts(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideLayoutCollection>> {
        let path = slf.borrow(py).revisions.capture(smallvec![]);
        Py::new(py, PySlideLayoutCollection::new(slf, path))
    }

    /// The package core properties under python-pptx's names.
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

    /// Renders slides as PNG or JPEG images, one per slide, or as one
    /// multi-page TIFF. `slides` picks zero-based slides, all by default.
    /// `quality` (JPEG, 90 by default) and `transparent` (PNG) raise for
    /// another format rather than being ignored.
    #[pyo3(signature = (*, dpi = 150.0, format = "png", quality = None, transparent = false, slides = None))]
    fn render_slides(
        &self,
        py: Python<'_>,
        dpi: f64,
        format: &str,
        quality: Option<i64>,
        transparent: bool,
        slides: Option<Vec<usize>>,
    ) -> PyResult<Py<PyAny>> {
        let format = raster_format(format, quality, transparent)?;
        let slides = slides.unwrap_or_else(|| (0..self.inner.len()).collect());
        let rendered = py
            .detach(|| {
                self.inner
                    .render_slides_deterministic(&slides, rpptx::RasterOptions { dpi, format })
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

    /// Renders an archival PDF, `pdfa-2b` or `pdfa-3b`.
    #[pyo3(signature = (profile = "pdfa-2b"))]
    fn to_pdfa<'py>(&self, py: Python<'py>, profile: &str) -> PyResult<Bound<'py, PyBytes>> {
        let profile = match profile {
            "pdfa-2b" => rpptx::PdfConformance::PdfA2b,
            "pdfa-3b" => rpptx::PdfConformance::PdfA3b,
            _ => {
                return Err(PyValueError::new_err(format!(
                    "profile must be pdfa-2b or pdfa-3b, not {profile:?}"
                )));
            }
        };
        py.detach(|| self.inner.to_pdfa_deterministic(profile))
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Renders audience handouts with 1, 2, 3, 4, 6, or 9 slides a page.
    #[pyo3(signature = (slides_per_page = 6))]
    fn to_handout_pdf<'py>(
        &self,
        py: Python<'py>,
        slides_per_page: u8,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let layout = handout_layout(slides_per_page)?;
        require_handout_master(py, &self.inner)?;
        py.detach(|| self.inner.to_handout_pdf_deterministic(layout))
            .map(|bytes| PyBytes::new(py, &bytes))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Renders one PNG per audience handout page.
    #[pyo3(signature = (slides_per_page = 6, dpi = 150.0))]
    fn render_all_handouts<'py>(
        &self,
        py: Python<'py>,
        slides_per_page: u8,
        dpi: f64,
    ) -> PyResult<Bound<'py, PyList>> {
        let layout = handout_layout(slides_per_page)?;
        require_handout_master(py, &self.inner)?;
        let pages = py
            .detach(|| self.inner.handout_page_pngs_deterministic(layout, dpi))
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, pages.iter().map(|page| PyBytes::new(py, page)))
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
        Ok((Self::from_presentation(result.presentation), diagnostics))
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

    #[getter]
    fn slides(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideCollection>> {
        let path = slf.borrow(py).revisions.capture(smallvec![]);
        Py::new(py, PySlideCollection::new(slf, path))
    }
}
