use oxml_py_support::{ContentPath, PathSeg};
use pyo3::PyClass;
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBytes, PyList, PySlice, PyTuple};
use smallvec::smallvec;

use crate::dml::{FillTarget, PyFillFormat};
use crate::normalize_index;
use crate::presentation::{PyComment, PyPresentation};
use crate::replacement_count_to_pyerr;
use crate::rpptx_to_pyerr;
use crate::shape::{MediaFile, PyPlaceholderCollection, PyShapeCollection};
use crate::validate_path;

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PySlideLayout>()?;
    module.add_class::<PySlideLayoutCollection>()?;
    module.add_class::<PySlide>()?;
    module.add_class::<PySlideCollection>()?;
    module.add_class::<PyBackground>()?;
    module.add_class::<PyMediaInfo>()?;
    Ok(())
}

/// One video or audio clip on a slide.
#[pyclass(name = "MediaInfo", frozen, get_all, eq, skip_from_py_object)]
#[derive(Clone, PartialEq, Eq)]
pub struct PyMediaInfo {
    /// The id of the picture shape that plays the clip.
    pub shape_id: u32,
    /// `video` or `audio`.
    pub kind: &'static str,
    /// The content type of an embedded clip, or `None` for a linked one.
    pub content_type: Option<String>,
    /// Whether the clip is linked rather than embedded.
    pub linked: bool,
    /// The package part of an embedded clip or the target of a linked one.
    pub target: String,
}

#[pymethods]
impl PyMediaInfo {
    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let text = |value: &str| pyo3::types::PyString::new(py, value).repr();
        Ok(format!(
            "MediaInfo(shape_id={}, kind={}, content_type={}, linked={}, target={})",
            self.shape_id,
            text(self.kind)?,
            match &self.content_type {
                Some(value) => text(value)?.to_string(),
                None => "None".to_owned(),
            },
            if self.linked { "True" } else { "False" },
            text(&self.target)?
        ))
    }
}

impl From<&rpptx::MediaInfo> for PyMediaInfo {
    fn from(info: &rpptx::MediaInfo) -> Self {
        let (content_type, linked, target) = match &info.source {
            rpptx::MediaLocation::Embedded {
                part_name,
                content_type,
            } => (Some(content_type.clone()), false, part_name.clone()),
            rpptx::MediaLocation::Linked { target } => (None, true, target.clone()),
        };
        Self {
            shape_id: info.shape_id,
            kind: match info.kind {
                rpptx::MediaKind::Audio => "audio",
                rpptx::MediaKind::Video => "video",
            },
            content_type,
            linked,
            target,
        }
    }
}

#[pyclass(name = "SlideLayout")]
pub struct PySlideLayout {
    pub(crate) presentation: Py<PyPresentation>,
    pub(crate) index: usize,
    path: ContentPath,
}

impl PySlideLayout {
    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "slide layout",
            &format!(".slide_layouts[{}]", self.index),
        )
    }
}

#[pymethods]
impl PySlideLayout {
    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.validate(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .layout_name(self.index)
            .map(str::to_owned))
    }

    /// Two handles are equal when they name the same layout of one presentation.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, PySlideLayout>>()
            .is_ok_and(|other| {
                other.presentation.is(&self.presentation) && other.index == self.index
            })
    }
}

#[pyclass(name = "SlideLayoutCollection")]
pub struct PySlideLayoutCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PySlideLayoutCollection {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "slide layout collection",
            ".slide_layouts",
        )?;
        Ok(self.presentation.borrow(py).inner.layout_count())
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PySlideLayout>> {
        Py::new(
            py,
            PySlideLayout {
                presentation: self.presentation.clone_ref(py),
                index,
                path: self.path.clone(),
            },
        )
    }
}

#[pymethods]
impl PySlideLayoutCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        sequence_item(py, key, self.len(py)?, "slide layout", |index| {
            self.item(py, index)
        })
    }

    /// Returns the zero-based index of a layout of this presentation.
    fn index(&self, py: Python<'_>, slide_layout: &Bound<'_, PyAny>) -> PyResult<usize> {
        self.len(py)?;
        let layout = slide_layout.extract::<PyRef<'_, PySlideLayout>>()?;
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "layout not in this SlideLayouts collection",
            ));
        }
        layout.validate(py)?;
        Ok(layout.index)
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PySlideLayoutIterator>> {
        self.len(py)?;
        Py::new(
            py,
            PySlideLayoutIterator {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PySlideLayoutIterator {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    index: usize,
}

#[pymethods]
impl PySlideLayoutIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PySlideLayout>>> {
        let collection =
            PySlideLayoutCollection::new(self.presentation.clone_ref(py), self.path.clone());
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

#[pyclass(name = "Slide")]
pub struct PySlide {
    pub(crate) presentation: Py<PyPresentation>,
    pub(crate) path: ContentPath,
}

impl PySlide {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "slide", "")?;
        self.path
            .segs
            .iter()
            .find_map(|segment| match segment {
                PathSeg::Slide(index) => Some(*index),
                _ => None,
            })
            .ok_or_else(|| PyIndexError::new_err("slide index is missing"))
    }
}

#[pymethods]
impl PySlide {
    /// Two current handles are equal when they name the same slide of one
    /// presentation. Like python-pptx `Slide`, a handle is not hashable.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, PySlide>>()
            .is_ok_and(|other| other.presentation.is(&self.presentation) && other.path == self.path)
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<Py<PyShapeCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    #[getter]
    fn notes_text(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let index = self.validate(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .slide(index)
            .and_then(|slide| slide.notes_text()))
    }

    /// Replaces the speaker notes, creating the notes slide when absent.
    #[setter]
    fn set_notes_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .set_notes_text(index, text)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    /// Replaces literal text on this slide, and in its speaker notes when
    /// `notes` is true, and returns the count.
    ///
    /// The contract is `Presentation.try_replace_text` restricted to this
    /// slide: with `expect`, a count that differs raises and leaves the
    /// presentation and its revision as they were, and the revision advances
    /// once only when something was replaced. The native call works on a
    /// copy of this slide alone.
    #[pyo3(signature = (placeholder, replacement, *, expect = None, notes = true))]
    fn try_replace_text(
        &self,
        py: Python<'_>,
        placeholder: &str,
        replacement: &str,
        expect: Option<usize>,
        notes: bool,
    ) -> PyResult<usize> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let inner = &mut presentation.inner;
        let count = py
            .detach(|| inner.try_replace_slide_text(index, placeholder, replacement, notes, expect))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        if let Some(expected) = expect.filter(|&expected| expected != count) {
            return Err(replacement_count_to_pyerr(py, placeholder, expected, count));
        }
        if count > 0 {
            presentation.revisions.bump();
        }
        Ok(count)
    }

    #[getter]
    fn slide_layout(&self, py: Python<'_>) -> PyResult<Py<PySlideLayout>> {
        let index = self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let layout = presentation
            .inner
            .slide_layout_index(index)
            .ok_or_else(|| {
                crate::rpptx_value_to_pyerr(
                    py,
                    format!("slide {index} has no layout reachable through the slide masters"),
                )
            })?;
        let path = presentation.revisions.capture(smallvec![]);
        drop(presentation);
        Py::new(
            py,
            PySlideLayout {
                presentation: self.presentation.clone_ref(py),
                index: layout,
                path,
            },
        )
    }

    /// Moves the slide to another layout of this presentation.
    ///
    /// Placeholders the new layout does not place keep the geometry they
    /// inherited, and the revision advances once.
    #[setter]
    fn set_slide_layout(&self, py: Python<'_>, layout: &Bound<'_, PyAny>) -> PyResult<()> {
        let index = self.validate(py)?;
        let layout = layout.extract::<PyRef<'_, PySlideLayout>>()?;
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "slide layout is not in this presentation",
            ));
        }
        layout.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .set_slide_layout(index, layout.index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    #[getter]
    fn hidden(&self, py: Python<'_>) -> PyResult<bool> {
        let index = self.validate(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .slide(index)
            .is_some_and(|slide| slide.hidden()))
    }

    #[setter]
    fn set_hidden(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let index = self.validate(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .slide_mut(index)
            .ok_or_else(|| PyIndexError::new_err(format!("slide index {index} is out of range")))?
            .set_hidden(value);
        Ok(())
    }

    /// The slide id, unique in the presentation and stable when slides move,
    /// as python-pptx `Slide.slide_id`.
    #[getter]
    fn slide_id(&self, py: Python<'_>) -> PyResult<u32> {
        let index = self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .slide(index)
            .map(|slide| slide.id())
            .ok_or_else(|| PyIndexError::new_err(format!("slide index {index} is out of range")))
    }

    /// The video and audio clips on the slide, in z-order.
    #[getter]
    fn media<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let index = self.validate(py)?;
        let media = self
            .presentation
            .borrow(py)
            .inner
            .media(index)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        PyTuple::new(py, media.iter().map(PyMediaInfo::from))
    }

    /// Returns the bytes of an embedded clip, found by shape id, or `None`
    /// for a linked one.
    fn extract_media<'py>(
        &self,
        py: Python<'py>,
        shape_id: u32,
    ) -> PyResult<Option<Bound<'py, PyBytes>>> {
        let index = self.validate(py)?;
        let bytes = self
            .presentation
            .borrow(py)
            .inner
            .extract_media(index, shape_id)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        Ok(bytes.map(|bytes| PyBytes::new(py, &bytes)))
    }

    /// Replaces the clip of one media shape, keeping its poster, position,
    /// and playback settings.
    #[pyo3(signature = (shape_id, media_file, mime_type = None))]
    fn replace_media(
        &self,
        py: Python<'_>,
        shape_id: u32,
        media_file: &Bound<'_, PyAny>,
        mime_type: Option<&str>,
    ) -> PyResult<()> {
        let index = self.validate(py)?;
        let media = MediaFile::read(media_file, mime_type)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .replace_media(index, shape_id, media.input())
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Removes one media shape with its clip, poster, and timing.
    fn remove_media(&self, py: Python<'_>, shape_id: u32) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .remove_media(index, shape_id)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    #[getter]
    fn background(&self, py: Python<'_>) -> PyResult<Py<PyBackground>> {
        self.validate(py)?;
        Py::new(
            py,
            PyBackground {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    /// Whether the slide inherits its background from its layout and master.
    #[getter]
    fn follow_master_background(&self, py: Python<'_>) -> PyResult<bool> {
        let index = self.validate(py)?;
        Ok(!self
            .presentation
            .borrow(py)
            .inner
            .slide(index)
            .is_some_and(|slide| slide.has_explicit_background()))
    }

    #[setter]
    fn set_follow_master_background(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let explicit = presentation
            .inner
            .slide(index)
            .is_some_and(|slide| slide.has_explicit_background());
        let mut slide = presentation
            .inner
            .slide_mut(index)
            .ok_or_else(|| PyIndexError::new_err(format!("slide index {index} is out of range")))?;
        if value {
            slide.remove_background();
            // A picture background leaves an image nothing shows any more.
            presentation
                .inner
                .release_unused_slide_images(index)
                .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        } else if !explicit {
            slide
                .set_background(rpptx::Fill::NoFill(rpptx::NoFill::default()))
                .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        }
        Ok(())
    }

    #[getter]
    fn comments<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let index = self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        PyTuple::new(
            py,
            presentation
                .inner
                .comments(index)
                .unwrap_or_default()
                .iter()
                .map(PyComment::from),
        )
    }

    // PyO3 exposes these as named Python arguments, so a Rust options wrapper
    // would only hide the public signature from this boundary.
    #[allow(clippy::too_many_arguments)]
    #[pyo3(signature = (*, id, author_id, created, text, shape_id=None, text_start=None, text_length=None))]
    fn add_comment(
        &self,
        id: String,
        author_id: String,
        created: String,
        text: &str,
        shape_id: Option<u32>,
        text_start: Option<usize>,
        text_length: Option<usize>,
        py: Python<'_>,
    ) -> PyResult<()> {
        let index = self.validate(py)?;
        let comment = rpptx::Comment::new(id, author_id, created, text)
            .map_err(|error| crate::rpptx_value_to_pyerr(py, error.to_string()))?;
        if text_start.is_some() != text_length.is_some()
            || (text_start.is_some() && shape_id.is_none())
        {
            return Err(PyValueError::new_err(
                "text_start and text_length require each other and shape_id",
            ));
        }
        let mut presentation = self.presentation.borrow_mut(py);
        match shape_id {
            Some(shape_id) => presentation.inner.add_comment_at_shape(
                index,
                comment,
                shape_id,
                text_start.zip(text_length),
            ),
            None => presentation.inner.add_comment(index, comment),
        }
        .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    #[pyo3(signature = (comment_id, *, id, author_id, created, text))]
    fn reply_to_comment(
        &self,
        comment_id: &str,
        id: String,
        author_id: String,
        created: String,
        text: &str,
        py: Python<'_>,
    ) -> PyResult<()> {
        let index = self.validate(py)?;
        let reply = rpptx::CommentReply::new(id, author_id, created, text)
            .map_err(|error| crate::rpptx_value_to_pyerr(py, error.to_string()))?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .reply_to_comment(index, comment_id, reply)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    /// Marks one comment thread resolved. A reply id is an unknown id.
    fn resolve_comment(&self, comment_id: &str, py: Python<'_>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .resolve_comment(index, comment_id)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    /// Removes one comment thread with its replies, or one reply.
    fn remove_comment(&self, comment_id: &str, py: Python<'_>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .remove_comment(index, comment_id)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    fn move_comment(&self, from_: usize, to: usize, py: Python<'_>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .move_comment(index, from_, to)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    fn move_reply(
        &self,
        comment_id: &str,
        from_: usize,
        to: usize,
        py: Python<'_>,
    ) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .move_reply(index, comment_id, from_, to)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }
}

#[pyclass(name = "SlideCollection")]
pub struct PySlideCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PySlideCollection {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "slide collection",
            ".slides",
        )?;
        Ok(self.presentation.borrow(py).inner.len())
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PySlide>> {
        let path = self
            .presentation
            .borrow(py)
            .revisions
            .capture(smallvec![PathSeg::Slide(index)]);
        Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))
    }
}

#[pymethods]
impl PySlideCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        sequence_item(py, key, self.len(py)?, "slide", |index| {
            self.item(py, index)
        })
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PySlideIterator>> {
        self.len(py)?;
        Py::new(
            py,
            PySlideIterator {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                index: 0,
            },
        )
    }

    fn add_slide(&self, py: Python<'_>, layout: &Bound<'_, PyAny>) -> PyResult<Py<PySlide>> {
        let layout = layout.extract::<PyRef<'_, PySlideLayout>>()?;
        self.len(py)?;
        layout.validate(py)?;
        let (index, path) = {
            let mut presentation = self.presentation.borrow_mut(py);
            let index = presentation.inner.len();
            presentation
                .inner
                .add_slide(layout.index)
                .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
            presentation.revisions.bump();
            let path = presentation
                .revisions
                .capture(smallvec![PathSeg::Slide(index)]);
            (index, path)
        };
        debug_assert!(matches!(path.segs.last(), Some(PathSeg::Slide(value)) if *value == index));
        Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))
    }

    /// Imports a slide of another presentation, using this presentation's theme.
    ///
    /// The copy takes `layout` when given. Otherwise a slide of this
    /// presentation keeps its own layout, and a slide of another presentation
    /// takes this presentation's first layout named like its layout. It is
    /// inserted before `index` as by `list.insert`, counted from the end when
    /// negative, or appended when that is `None`. An index outside the
    /// collection raises `IndexError`.
    #[pyo3(signature = (slide, layout=None, index=None))]
    fn import_slide(
        &self,
        py: Python<'_>,
        slide: &Bound<'_, PyAny>,
        layout: Option<&Bound<'_, PyAny>>,
        index: Option<isize>,
    ) -> PyResult<Py<PySlide>> {
        let len = self.len(py)?;
        let slide = slide.extract::<PyRef<'_, PySlide>>()?;
        let source_index = slide.validate(py)?;
        let layout_index = match layout {
            Some(layout) => {
                let layout = layout.extract::<PyRef<'_, PySlideLayout>>()?;
                if !layout.presentation.is(&self.presentation) {
                    return Err(PyValueError::new_err(
                        "layout is not a layout of this presentation",
                    ));
                }
                layout.validate(py)?;
                Some(layout.index)
            }
            None => None,
        };
        let inserted = match index {
            None => len,
            Some(index) => {
                let normalized = if index < 0 {
                    len as isize + index
                } else {
                    index
                };
                if normalized < 0 || normalized > len as isize {
                    return Err(PyIndexError::new_err("slide index out of range"));
                }
                normalized as usize
            }
        };
        let source = slide.presentation.clone_ref(py);
        drop(slide);
        let path = {
            let mut presentation = self.presentation.borrow_mut(py);
            let result = if source.is(&self.presentation) {
                let copy = presentation.inner.clone();
                let layout_index = layout_index.or_else(|| copy.slide_layout_index(source_index));
                presentation
                    .inner
                    .import_slide(&copy, source_index, layout_index, Some(inserted))
                    .map(|_| ())
            } else {
                let source = source.borrow(py);
                presentation
                    .inner
                    .import_slide(&source.inner, source_index, layout_index, Some(inserted))
                    .map(|_| ())
            };
            result.map_err(|error| crate::rpptx_to_pyerr(py, error))?;
            presentation.revisions.bump();
            presentation
                .revisions
                .capture(smallvec![PathSeg::Slide(inserted)])
        };
        Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))
    }

    /// Removes one slide of this presentation with its notes and owned media.
    fn remove(&self, py: Python<'_>, slide: &Bound<'_, PyAny>) -> PyResult<()> {
        self.len(py)?;
        let slide = slide.extract::<PyRef<'_, PySlide>>()?;
        if !slide.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("slide is not in this collection"));
        }
        let index = slide.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .remove_slide(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    /// Duplicates one slide of this presentation with its notes right after
    /// the source, and returns the new slide.
    fn duplicate(&self, py: Python<'_>, slide: &Bound<'_, PyAny>) -> PyResult<Py<PySlide>> {
        self.len(py)?;
        let slide = slide.extract::<PyRef<'_, PySlide>>()?;
        if !slide.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("slide is not in this collection"));
        }
        let index = slide.validate(py)?;
        let path = {
            let mut presentation = self.presentation.borrow_mut(py);
            presentation
                .inner
                .duplicate_slide(index)
                .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
            presentation.revisions.bump();
            presentation
                .revisions
                .capture(smallvec![PathSeg::Slide(index + 1)])
        };
        Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))
    }

    /// Moves the slide at `from_` so that it ends up at index `to`.
    #[pyo3(name = "move")]
    fn move_slide(&self, py: Python<'_>, from_: isize, to: isize) -> PyResult<()> {
        let len = self.len(py)?;
        let from_ = normalize_index(from_, len, "slide")?;
        let to = normalize_index(to, len, "slide")?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .move_slide(from_, to)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }
}

/// The background of one slide, like python-pptx `_Background`.
#[pyclass(name = "Background")]
pub struct PyBackground {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyBackground {
    /// The slide's own background fill. Reading it never changes the slide.
    #[getter]
    fn fill(&self, py: Python<'_>) -> PyResult<Py<PyFillFormat>> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "background",
            ".background",
        )?;
        Py::new(
            py,
            PyFillFormat::new(
                self.presentation.clone_ref(py),
                self.path.clone(),
                FillTarget::Background,
            ),
        )
    }
}

#[pyclass]
struct PySlideIterator {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    index: usize,
}

#[pymethods]
impl PySlideIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PySlide>>> {
        let collection = PySlideCollection::new(self.presentation.clone_ref(py), self.path.clone());
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

fn sequence_item<T, F>(
    py: Python<'_>,
    key: &Bound<'_, PyAny>,
    len: usize,
    kind: &str,
    mut item: F,
) -> PyResult<Py<PyAny>>
where
    T: PyClass,
    F: FnMut(usize) -> PyResult<Py<T>>,
{
    if let Ok(index) = key.extract::<isize>() {
        return Ok(item(normalize_index(index, len, kind)?)?.into_any());
    }
    if key.is_instance_of::<PySlice>() {
        let (start, stop, step): (isize, isize, isize) =
            key.call_method1("indices", (len,))?.extract()?;
        let items = PyList::empty(py);
        let mut index = start;
        while if step > 0 { index < stop } else { index > stop } {
            items.append(item(index as usize)?)?;
            index += step;
        }
        return Ok(items.into_any().unbind());
    }
    Err(PyTypeError::new_err(format!(
        "{kind} indices must be integers or slices"
    )))
}
