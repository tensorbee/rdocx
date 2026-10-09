use std::path::PathBuf;

use oxml_py_support::RevisionCounter;
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

    #[pyo3(signature = (*, fonts = None, font_dir = None))]
    fn to_pdf<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let fonts = caller_fonts(fonts, font_dir)?;
        py.detach(|| {
            self.inner
                .to_pdf_deterministic_with_fonts(&rpptx::FontFile::as_refs(&fonts))
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (slide_index, dpi = 150.0, *, fonts = None, font_dir = None))]
    fn render_slide_to_png<'py>(
        &self,
        py: Python<'py>,
        slide_index: usize,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
    ) -> PyResult<Option<Bound<'py, PyBytes>>> {
        let fonts = caller_fonts(fonts, font_dir)?;
        py.detach(|| {
            self.inner.slide_png_deterministic_with_fonts(
                slide_index,
                dpi,
                &rpptx::FontFile::as_refs(&fonts),
            )
        })
        .map(|bytes| bytes.map(|bytes| PyBytes::new(py, &bytes)))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None))]
    fn render_all_slides<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
    ) -> PyResult<Bound<'py, PyList>> {
        let fonts = caller_fonts(fonts, font_dir)?;
        let slides = py
            .detach(|| {
                self.inner
                    .slide_pngs_deterministic_with_fonts(dpi, &rpptx::FontFile::as_refs(&fonts))
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyList::new(py, slides.iter().map(|slide| PyBytes::new(py, slide)))
    }

    #[pyo3(signature = (*, width_factor = 1.0, fonts = None, font_dir = None))]
    fn text_layout<'py>(
        &self,
        py: Python<'py>,
        width_factor: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        let fonts = caller_fonts(fonts, font_dir)?;
        let frames = py
            .detach(|| {
                self.inner.text_layout_deterministic_with_fonts(
                    width_factor,
                    &rpptx::FontFile::as_refs(&fonts),
                )
            })
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyTuple::new(py, frames.iter().map(PyTextFrameLayout::from))
    }

    #[pyo3(signature = (*, fonts = None, font_dir = None))]
    fn to_notes_pdf<'py>(
        &self,
        py: Python<'py>,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
    ) -> PyResult<Bound<'py, PyBytes>> {
        let fonts = caller_fonts(fonts, font_dir)?;
        py.detach(|| {
            self.inner
                .to_notes_pdf_deterministic_with_fonts(&rpptx::FontFile::as_refs(&fonts))
        })
        .map(|bytes| PyBytes::new(py, &bytes))
        .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[pyo3(signature = (dpi = 150.0, *, fonts = None, font_dir = None))]
    fn render_all_notes<'py>(
        &self,
        py: Python<'py>,
        dpi: f64,
        fonts: Option<Vec<(String, Bound<'py, PyBytes>)>>,
        font_dir: Option<PathBuf>,
    ) -> PyResult<Bound<'py, PyList>> {
        let fonts = caller_fonts(fonts, font_dir)?;
        let notes = py
            .detach(|| {
                self.inner.notes_page_pngs_deterministic_with_fonts(
                    dpi,
                    &rpptx::FontFile::as_refs(&fonts),
                )
            })
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

    #[getter]
    fn slides(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<PySlideCollection>> {
        let path = slf.borrow(py).revisions.capture(smallvec![]);
        Py::new(py, PySlideCollection::new(slf, path))
    }
}

/// The fonts of a `fonts=` and `font_dir=` pair, the `(family, bytes)` pairs
/// first. A missing `font_dir` raises `FileNotFoundError` and a file
/// `NotADirectoryError`, before any layout.
fn caller_fonts(
    fonts: Option<Vec<(String, Bound<'_, PyBytes>)>>,
    font_dir: Option<PathBuf>,
) -> PyResult<Vec<rpptx::FontFile>> {
    let mut files = fonts
        .unwrap_or_default()
        .into_iter()
        .map(|(family, data)| rpptx::FontFile {
            family,
            data: data.as_bytes().to_vec(),
        })
        .collect::<Vec<_>>();
    if let Some(font_dir) = font_dir {
        files.extend(rpptx::FontFile::load_dir(&font_dir)?);
    }
    Ok(files)
}
