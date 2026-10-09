use oxml_py_support::{ContentPath, PathSeg};
use pyo3::PyClass;
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice, PyTuple};
use smallvec::smallvec;

use crate::dml::{FillTarget, PyFillFormat};
use crate::normalize_index;
use crate::presentation::{PyComment, PyPresentation};
use crate::replacement_count_to_pyerr;
use crate::shape::{PyPlaceholderCollection, PyShapeCollection};
use crate::validate_path;

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PySlideLayout>()?;
    module.add_class::<PySlideLayoutCollection>()?;
    module.add_class::<PySlide>()?;
    module.add_class::<PySlideCollection>()?;
    module.add_class::<PyBackground>()?;
    module.add_class::<PyHeaderFooter>()?;
    module.add_class::<PySlideTransition>()?;
    module.add_class::<PySlideMaster>()?;
    module.add_class::<PySlideMasterCollection>()?;
    module.add_class::<PyTheme>()?;
    module.add_class::<PyThemeFonts>()?;
    module.add_class::<PyThemeFontSet>()?;
    module.add_class::<PyThemeColors>()?;
    Ok(())
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

    /// The path of the layout's own content at the current revision.
    fn part_path(&self, py: Python<'_>) -> ContentPath {
        self.presentation
            .borrow(py)
            .revisions
            .capture(smallvec![PathSeg::Layout(self.index)])
    }
}

/// Whether a slide or layout follows the background above it.
fn follow_master_background(
    py: Python<'_>,
    presentation: &Py<PyPresentation>,
    part: rpptx::PartRef,
) -> PyResult<bool> {
    presentation
        .borrow(py)
        .inner
        .part_has_background(part)
        .map(|explicit| !explicit)
        .map_err(|error| crate::rpptx_to_pyerr(py, error))
}

/// Following removes the part's own background, and not following gives it
/// an empty no-fill one, as python-pptx does.
fn set_follow_master_background(
    py: Python<'_>,
    presentation: &Py<PyPresentation>,
    part: rpptx::PartRef,
    value: bool,
) -> PyResult<()> {
    let mut presentation = presentation.borrow_mut(py);
    let explicit = presentation
        .inner
        .part_has_background(part)
        .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
    let result = if value {
        presentation.inner.remove_part_background(part)
    } else if !explicit {
        presentation
            .inner
            .set_part_background(part, rpptx::Fill::NoFill(rpptx::NoFill::default()))
    } else {
        Ok(())
    };
    result.map_err(|error| crate::rpptx_to_pyerr(py, error))
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

    // Masters, layouts and themes (#312).

    /// Renames the layout, as PowerPoint's Rename Layout does.
    #[setter]
    fn set_name(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.validate(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_layout_name(self.index, value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    /// The slide master this layout belongs to.
    #[getter]
    fn slide_master(&self, py: Python<'_>) -> PyResult<Py<PySlideMaster>> {
        self.validate(py)?;
        let master = self
            .presentation
            .borrow(py)
            .inner
            .layout_master(self.index)
            .ok_or_else(|| {
                PyValueError::new_err("the layout is not reached from a slide master")
            })?;
        let path = self.presentation.borrow(py).revisions.capture(smallvec![]);
        PySlideMasterCollection::new(self.presentation.clone_ref(py), path).item(py, master)
    }

    /// The layout's own shapes, such as a logo every slide on it shows.
    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<Py<PyShapeCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeCollection::new(self.presentation.clone_ref(py), self.part_path(py)),
        )
    }

    /// The layout's placeholders, which new slides copy and slides inherit.
    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), self.part_path(py)),
        )
    }

    /// The layout's own background, which its slides follow.
    #[getter]
    fn background(&self, py: Python<'_>) -> PyResult<Py<PyBackground>> {
        self.validate(py)?;
        Py::new(
            py,
            PyBackground {
                presentation: self.presentation.clone_ref(py),
                path: self.part_path(py),
            },
        )
    }

    /// Whether the layout inherits its master's background.
    #[getter]
    fn follow_master_background(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        follow_master_background(py, &self.presentation, rpptx::PartRef::Layout(self.index))
    }

    #[setter]
    fn set_follow_master_background(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.validate(py)?;
        set_follow_master_background(
            py,
            &self.presentation,
            rpptx::PartRef::Layout(self.index),
            value,
        )
    }

    /// Whether slides on this layout show the master's shapes, such as a
    /// logo. False hides them, as PowerPoint's Hide Background Graphics
    /// does on a layout.
    #[getter]
    fn show_master_shapes(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .show_master_shapes(rpptx::PartRef::Layout(self.index))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    #[setter]
    fn set_show_master_shapes(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.validate(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_show_master_shapes(rpptx::PartRef::Layout(self.index), value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    /// The slides that use this layout, like python-pptx `used_by_slides`.
    #[getter]
    fn used_by_slides<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let slides = presentation
            .inner
            .layout_slides(self.index)
            .into_iter()
            .map(|index| {
                presentation
                    .revisions
                    .capture(smallvec![PathSeg::Slide(index)])
            })
            .collect::<Vec<_>>();
        drop(presentation);
        let items = slides
            .into_iter()
            .map(|path| Py::new(py, PySlide::new(self.presentation.clone_ref(py), path)))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, items)
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
    /// The master whose layouts this collection holds, or every master's
    /// layouts in order for `prs.slide_layouts`.
    master: Option<usize>,
}

impl PySlideLayoutCollection {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self {
            presentation,
            path,
            master: None,
        }
    }

    /// The presentation-wide indices of this collection's layouts, in order.
    fn layouts(&self, py: Python<'_>) -> PyResult<Vec<usize>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "slide layout collection",
            &match self.master {
                Some(master) => format!(".slide_masters[{master}].slide_layouts"),
                None => ".slide_layouts".to_owned(),
            },
        )?;
        match self.master {
            None => Ok((0..presentation.inner.layout_count()).collect()),
            Some(master) => presentation
                .inner
                .master_layouts(master)
                .ok_or_else(|| PyIndexError::new_err("slide master index out of range")),
        }
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.layouts(py)?.len())
    }

    /// The layout at a position of this collection.
    fn item(&self, py: Python<'_>, position: usize) -> PyResult<Py<PySlideLayout>> {
        let index = match self.master {
            None => position,
            Some(_) => *self
                .layouts(py)?
                .get(position)
                .ok_or_else(|| PyIndexError::new_err("slide layout index out of range"))?,
        };
        self.layout(py, index)
    }

    /// A handle on one layout by its presentation-wide index.
    fn layout(&self, py: Python<'_>, index: usize) -> PyResult<Py<PySlideLayout>> {
        Py::new(
            py,
            PySlideLayout {
                presentation: self.presentation.clone_ref(py),
                index,
                path: self.path.clone(),
            },
        )
    }

    /// The position of a layout in this collection and its presentation-wide index.
    fn find(&self, py: Python<'_>, slide_layout: &Bound<'_, PyAny>) -> PyResult<(usize, usize)> {
        let layouts = self.layouts(py)?;
        let layout = slide_layout.extract::<PyRef<'_, PySlideLayout>>()?;
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "layout not in this SlideLayouts collection",
            ));
        }
        layout.validate(py)?;
        let position = layouts
            .iter()
            .position(|index| *index == layout.index)
            .ok_or_else(|| PyValueError::new_err("layout not in this SlideLayouts collection"))?;
        Ok((position, layout.index))
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

    /// Returns the zero-based position of a layout in this collection.
    fn index(&self, py: Python<'_>, slide_layout: &Bound<'_, PyAny>) -> PyResult<usize> {
        Ok(self.find(py, slide_layout)?.0)
    }

    /// Returns the first layout named `name`, or `default`, like
    /// python-pptx `SlideLayouts.get_by_name`.
    #[pyo3(signature = (name, default = None))]
    fn get_by_name(
        &self,
        py: Python<'_>,
        name: &str,
        default: Option<Py<PyAny>>,
    ) -> PyResult<Option<Py<PyAny>>> {
        let found = {
            let layouts = self.layouts(py)?;
            let presentation = self.presentation.borrow(py);
            layouts
                .into_iter()
                .find(|index| presentation.inner.layout_name(*index) == Some(name))
        };
        match found {
            Some(index) => Ok(Some(self.layout(py, index)?.into_any())),
            None => Ok(default),
        }
    }

    /// Removes a layout no slide uses, like python-pptx
    /// `SlideLayouts.remove`. The last layout of a master stays. Layout
    /// handles are invalidated.
    fn remove(&self, py: Python<'_>, slide_layout: &Bound<'_, PyAny>) -> PyResult<()> {
        let (_, index) = self.find(py, slide_layout)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let used = presentation.inner.layout_slides(index);
        if !used.is_empty() {
            return Err(PyValueError::new_err(format!(
                "cannot remove slide-layout in use by one or more slides (slides {used:?}), set their slide_layout to another layout first"
            )));
        }
        let siblings = presentation
            .inner
            .layout_master(index)
            .and_then(|master| presentation.inner.master_layouts(master))
            .map_or(0, |layouts| layouts.len());
        if siblings <= 1 {
            return Err(PyValueError::new_err(
                "cannot remove the only slide-layout of its slide master",
            ));
        }
        presentation
            .inner
            .remove_layout(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    /// Copies a layout into a new layout of the same master, placed right
    /// after it and named as PowerPoint names a copy (rpptx extension).
    /// Layout handles are invalidated, and the new layout is returned.
    fn duplicate(
        &self,
        py: Python<'_>,
        slide_layout: &Bound<'_, PyAny>,
    ) -> PyResult<Py<PySlideLayout>> {
        let (_, index) = self.find(py, slide_layout)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let copy = presentation
            .inner
            .duplicate_layout(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        let path = presentation.revisions.capture(smallvec![]);
        drop(presentation);
        Py::new(
            py,
            PySlideLayout {
                presentation: self.presentation.clone_ref(py),
                index: copy,
                path,
            },
        )
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let items = self
            .layouts(py)?
            .into_iter()
            .map(|index| self.layout(py, index))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, items)?
            .into_any()
            .try_iter()
            .map(Bound::into_any)
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

    /// Whether the slide shows its layout's and master's shapes, such as a
    /// logo. False hides them, as PowerPoint's Hide Background Graphics does.
    #[getter]
    fn show_master_shapes(&self, py: Python<'_>) -> PyResult<bool> {
        let index = self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .show_master_shapes(rpptx::PartRef::Slide(index))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    #[setter]
    fn set_show_master_shapes(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let index = self.validate(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_show_master_shapes(rpptx::PartRef::Slide(index), value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
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

    // Header and footer, transition (#311).

    /// The slide's own date, footer and slide number (rpptx extension).
    #[getter]
    fn header_footer(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.validate(py)?;
        Py::new(
            py,
            PyHeaderFooter {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    /// The slide's transition and advance timing (rpptx extension).
    #[getter]
    fn transition(&self, py: Python<'_>) -> PyResult<Py<PySlideTransition>> {
        self.validate(py)?;
        Py::new(
            py,
            PySlideTransition {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
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

// Header and footer, transitions, masters and themes (#311).

/// The index path of every shape on a slide, groups included, by shape id.
fn shape_paths(
    presentation: &rpptx::Presentation,
    slide_index: usize,
) -> std::collections::HashMap<u32, Vec<usize>> {
    fn walk<'a>(
        shapes: impl Iterator<Item = rpptx::ShapeRef<'a>>,
        prefix: &[usize],
        paths: &mut std::collections::HashMap<u32, Vec<usize>>,
    ) {
        for (index, shape) in shapes.enumerate() {
            let mut path = prefix.to_vec();
            path.push(index);
            walk(shape.children(), &path, paths);
            if let Some(id) = shape.non_visual_id() {
                paths.insert(id, path);
            }
        }
    }
    let mut paths = std::collections::HashMap::new();
    if let Some(slide) = presentation.slide(slide_index) {
        walk(slide.shapes(), &[], &mut paths);
    }
    paths
}

/// Every slide's shape paths, to tell whether a header-footer edit moved any.
pub(crate) fn all_shape_paths(
    presentation: &rpptx::Presentation,
) -> Vec<std::collections::HashMap<u32, Vec<usize>>> {
    (0..presentation.len())
        .map(|index| shape_paths(presentation, index))
        .collect()
}

/// Whether every shape that survived an edit kept its index path. Date,
/// footer and slide-number placeholders are appended after the other shapes,
/// so adding them or rewriting their text leaves every held handle valid,
/// and only removing one before another shape moves anything.
pub(crate) fn paths_kept(
    before: &[std::collections::HashMap<u32, Vec<usize>>],
    after: &[std::collections::HashMap<u32, Vec<usize>>],
) -> bool {
    before.len() == after.len()
        && before.iter().zip(after).all(|(before, after)| {
            before
                .iter()
                .all(|(id, path)| after.get(id).is_none_or(|moved| moved == path))
        })
}

/// Today's local date as the cached text of a `datetime1` to `datetime7` field.
///
/// `hint` names how the caller supplies the text instead for a time format.
pub(crate) fn today_field_text(py: Python<'_>, field_type: &str, hint: &str) -> PyResult<String> {
    let today = py
        .import("datetime")?
        .getattr("date")?
        .call_method0("today")?;
    let year: i32 = today.getattr("year")?.extract()?;
    let month: u32 = today.getattr("month")?.extract()?;
    let day: u32 = today.getattr("day")?.extract()?;
    rpptx::date_field_text(field_type, year, month, day).map_err(|_| {
        PyValueError::new_err(format!(
            "{field_type} is not a date format rpptx can fill with today's date, use datetime1 to datetime7{hint}"
        ))
    })
}

/// Builds the Rust settings from python values: `date` is `None`, fixed
/// text, or `"auto"` for a `date_format` field holding today's date.
pub(crate) fn header_footer_settings(
    py: Python<'_>,
    slide_number: bool,
    footer: Option<String>,
    date: Option<String>,
    date_format: &str,
) -> PyResult<rpptx::HeaderFooter> {
    let date = match date.as_deref() {
        None => rpptx::HeaderFooterDate::Off,
        Some("auto") => rpptx::HeaderFooterDate::Automatic {
            field_type: date_format.to_owned(),
            text: today_field_text(py, date_format, "")?,
        },
        Some(text) => rpptx::HeaderFooterDate::Fixed(text.to_owned()),
    };
    Ok(rpptx::HeaderFooter {
        slide_number,
        footer,
        date,
    })
}

/// One slide's date, footer and slide number, read and written live.
///
/// Each assignment adds or removes the slide's own placeholders, copied from
/// its layout. Added placeholders go after the other shapes, so held slide
/// and shape handles stay valid. Only removing a placeholder that other
/// shapes follow advances the revision, which stales held handles.
#[pyclass(name = "HeaderFooter")]
pub struct PyHeaderFooter {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyHeaderFooter {
    fn slide_index(&self, py: Python<'_>) -> PyResult<usize> {
        PySlide::new(self.presentation.clone_ref(py), self.path.clone()).validate(py)
    }

    fn current(&self, py: Python<'_>) -> PyResult<(usize, rpptx::HeaderFooter)> {
        let index = self.slide_index(py)?;
        let settings = self
            .presentation
            .borrow(py)
            .inner
            .slide_header_footer(index)
            .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?;
        Ok((index, settings))
    }

    /// Writes the change and advances the revision. This handle names the
    /// slide, which stays where it is, so it follows the new revision.
    fn update(
        &mut self,
        py: Python<'_>,
        change: impl FnOnce(&mut rpptx::HeaderFooter),
    ) -> PyResult<()> {
        let (index, mut settings) = self.current(py)?;
        change(&mut settings);
        let mut presentation = self.presentation.borrow_mut(py);
        let before = shape_paths(&presentation.inner, index);
        presentation
            .inner
            .set_slide_header_footer(index, &settings)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let after = shape_paths(&presentation.inner, index);
        if !paths_kept(&[before], &[after]) {
            presentation.revisions.bump();
            self.path = presentation.revisions.capture(self.path.segs.clone());
        }
        Ok(())
    }
}

#[pymethods]
impl PyHeaderFooter {
    /// Whether the slide owns a slide-number placeholder.
    #[getter]
    fn slide_number(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.current(py)?.1.slide_number)
    }

    #[setter]
    fn set_slide_number(&mut self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.update(py, |settings| settings.slide_number = value)
    }

    /// The footer text, `None` without a footer placeholder.
    #[getter]
    fn footer(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.current(py)?.1.footer)
    }

    #[setter]
    fn set_footer(&mut self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        self.update(py, |settings| settings.footer = value)
    }

    /// `None`, the fixed date text, or `"auto"` for a date field.
    #[getter]
    fn date(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(match self.current(py)?.1.date {
            rpptx::HeaderFooterDate::Off => None,
            rpptx::HeaderFooterDate::Fixed(text) => Some(text),
            rpptx::HeaderFooterDate::Automatic { .. } => Some("auto".to_owned()),
        })
    }

    /// Sets `None`, fixed text, or `"auto"` for a field PowerPoint refreshes,
    /// caching today's date in the current `date_format`, else `datetime1`.
    #[setter]
    fn set_date(&mut self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        let date_format = self
            .date_format(py)?
            .unwrap_or_else(|| "datetime1".to_owned());
        let date = header_footer_settings(py, false, None, value, &date_format)?.date;
        self.update(py, |settings| settings.date = date)
    }

    /// The date field type, such as `datetime1`, `None` for a fixed or no date.
    #[getter]
    fn date_format(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(match self.current(py)?.1.date {
            rpptx::HeaderFooterDate::Automatic { field_type, .. } => Some(field_type),
            _ => None,
        })
    }

    #[setter]
    fn set_date_format(&mut self, py: Python<'_>, value: &str) -> PyResult<()> {
        if self.date_format(py)?.is_none() {
            return Err(PyValueError::new_err(
                "the slide has no date field, set header_footer.date = \"auto\" first",
            ));
        }
        let date = header_footer_settings(py, false, None, Some("auto".to_owned()), value)?.date;
        self.update(py, |settings| settings.date = date)
    }
}

/// One slide's transition, read and written live (rpptx extension).
///
/// `type` is one of `fade`, `push`, `wipe`, `split`, `cover`, `uncover`,
/// `cut`, `zoom`, or `None`. Another producer's effect reads as its element
/// name and cannot be assigned. Durations are in seconds.
#[pyclass(name = "SlideTransition")]
pub struct PySlideTransition {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PySlideTransition {
    fn current(&self, py: Python<'_>) -> PyResult<(usize, rpptx::SlideTransition)> {
        let index =
            PySlide::new(self.presentation.clone_ref(py), self.path.clone()).validate(py)?;
        let transition = self
            .presentation
            .borrow(py)
            .inner
            .slide(index)
            .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?
            .transition()
            .unwrap_or_default();
        Ok((index, transition))
    }

    fn update(
        &self,
        py: Python<'_>,
        change: impl FnOnce(&mut rpptx::SlideTransition) -> PyResult<()>,
    ) -> PyResult<()> {
        let (index, mut transition) = self.current(py)?;
        change(&mut transition)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .slide_mut(index)
            .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?
            .set_transition(Some(&transition))
            // The only failures are values the effect does not take.
            .map_err(|error| PyValueError::new_err(error.to_string()))
    }
}

fn seconds_to_ms(name: &str, value: Option<f64>) -> PyResult<Option<u64>> {
    value
        .map(|seconds| {
            if seconds.is_finite() && (0.0..=86_400.0).contains(&seconds) {
                Ok((seconds * 1000.0).round() as u64)
            } else {
                Err(PyValueError::new_err(format!(
                    "{name} must be between 0 and 86400 seconds, got {seconds}"
                )))
            }
        })
        .transpose()
}

#[pymethods]
impl PySlideTransition {
    #[getter]
    fn r#type(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.current(py)?.1.kind.map(|kind| kind.name().to_owned()))
    }

    /// Sets the effect, keeping the direction when the new effect takes it.
    #[setter]
    fn set_type(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        let kind = value
            .map(|name| {
                rpptx::TransitionKind::parse(&name).ok_or_else(|| {
                    PyValueError::new_err(format!(
                        "unknown transition type {name:?}, use one of {}",
                        rpptx::TransitionKind::NAMES.join(", ")
                    ))
                })
            })
            .transpose()?;
        self.update(py, |transition| {
            let keeps_direction = transition.direction.is_some_and(|direction| {
                kind.as_ref()
                    .is_some_and(|kind| kind.directions().contains(&direction))
            });
            if !keeps_direction {
                transition.direction = None;
            }
            if kind.is_none() {
                transition.duration_ms = None;
            }
            transition.kind = kind;
            Ok(())
        })
    }

    /// The direction the incoming slide moves, such as `left` (`dir="l"`,
    /// PowerPoint's "From Right"), `None` for the effect's default.
    #[getter]
    fn direction(&self, py: Python<'_>) -> PyResult<Option<&'static str>> {
        Ok(self
            .current(py)?
            .1
            .direction
            .map(rpptx::TransitionDirection::name))
    }

    #[setter]
    fn set_direction(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        let direction = value
            .map(|name| {
                rpptx::TransitionDirection::parse(&name).ok_or_else(|| {
                    PyValueError::new_err(format!(
                        "unknown transition direction {name:?}, use one of {}",
                        rpptx::TransitionDirection::NAMES.join(", ")
                    ))
                })
            })
            .transpose()?;
        self.update(py, |transition| {
            transition.direction = direction;
            Ok(())
        })
    }

    /// The effect duration in seconds, `None` for the default speed.
    #[getter]
    fn duration(&self, py: Python<'_>) -> PyResult<Option<f64>> {
        Ok(self.current(py)?.1.duration_ms.map(|ms| ms as f64 / 1000.0))
    }

    #[setter]
    fn set_duration(&self, py: Python<'_>, value: Option<f64>) -> PyResult<()> {
        let duration = seconds_to_ms("duration", value)?;
        self.update(py, |transition| {
            transition.duration_ms = duration;
            Ok(())
        })
    }

    /// Whether a click advances the slide show, true by default.
    #[getter]
    fn advance_on_click(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.current(py)?.1.advance_on_click)
    }

    #[setter]
    fn set_advance_on_click(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.update(py, |transition| {
            transition.advance_on_click = value;
            Ok(())
        })
    }

    /// Seconds after which the slide show advances by itself, or `None`.
    #[getter]
    fn advance_after(&self, py: Python<'_>) -> PyResult<Option<f64>> {
        Ok(self
            .current(py)?
            .1
            .advance_after_ms
            .map(|ms| ms as f64 / 1000.0))
    }

    #[setter]
    fn set_advance_after(&self, py: Python<'_>, value: Option<f64>) -> PyResult<()> {
        let after = seconds_to_ms("advance_after", value)?;
        self.update(py, |transition| {
            transition.advance_after_ms = after;
            Ok(())
        })
    }

    /// Gives every slide this slide's transition, as PowerPoint's Apply To All.
    fn apply_to_all(&self, py: Python<'_>) -> PyResult<()> {
        let (_, transition) = self.current(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_all_transitions(Some(&transition))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }
}

/// A slide master: its layouts, shapes, background, text styles and theme.
#[pyclass(name = "SlideMaster")]
pub struct PySlideMaster {
    presentation: Py<PyPresentation>,
    index: usize,
    path: ContentPath,
}

impl PySlideMaster {
    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "slide master",
            &format!(".slide_masters[{}]", self.index),
        )
    }

    /// The path of the master's own content at the current revision.
    fn part_path(&self, py: Python<'_>) -> ContentPath {
        self.presentation
            .borrow(py)
            .revisions
            .capture(smallvec![PathSeg::Master(self.index)])
    }
}

#[pymethods]
impl PySlideMaster {
    /// The theme this master uses. Its colours and fonts are read and
    /// written live.
    #[getter]
    fn theme(&self, py: Python<'_>) -> PyResult<Py<PyTheme>> {
        self.validate(py)?;
        let theme = PyTheme {
            presentation: self.presentation.clone_ref(py),
            master: self.index,
            path: self.path.clone(),
        };
        theme.read(py)?;
        Py::new(py, theme)
    }

    // Masters, layouts and themes (#312).

    /// The master's own shapes, such as a logo every slide shows.
    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<Py<PyShapeCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeCollection::new(self.presentation.clone_ref(py), self.part_path(py)),
        )
    }

    /// The master's placeholders, which layouts and slides inherit.
    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), self.part_path(py)),
        )
    }

    /// The master's background, which layouts and slides follow.
    #[getter]
    fn background(&self, py: Python<'_>) -> PyResult<Py<PyBackground>> {
        self.validate(py)?;
        Py::new(
            py,
            PyBackground {
                presentation: self.presentation.clone_ref(py),
                path: self.part_path(py),
            },
        )
    }

    /// The master's title, body and other text styles (rpptx extension).
    #[getter]
    fn text_styles(&self, py: Python<'_>) -> PyResult<Py<crate::text::PyMasterTextStyles>> {
        self.validate(py)?;
        Py::new(
            py,
            crate::text::PyMasterTextStyles::new(
                self.presentation.clone_ref(py),
                self.part_path(py),
            ),
        )
    }

    /// The layouts this master owns, like python-pptx
    /// `slide_master.slide_layouts`, with `remove`, `duplicate`,
    /// `get_by_name` and `index` scoped to this master.
    #[getter]
    fn slide_layouts(&self, py: Python<'_>) -> PyResult<Py<PySlideLayoutCollection>> {
        self.validate(py)?;
        let mut collection =
            PySlideLayoutCollection::new(self.presentation.clone_ref(py), self.path.clone());
        collection.master = Some(self.index);
        collection.layouts(py)?;
        Py::new(py, collection)
    }

    /// Two handles are equal when they name the same master of one presentation.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, PySlideMaster>>()
            .is_ok_and(|other| {
                other.presentation.is(&self.presentation) && other.index == self.index
            })
    }
}

#[pyclass(name = "SlideMasterCollection")]
pub struct PySlideMasterCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PySlideMasterCollection {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "slide master collection",
            ".slide_masters",
        )?;
        Ok(self.presentation.borrow(py).inner.master_count())
    }

    pub(crate) fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PySlideMaster>> {
        Py::new(
            py,
            PySlideMaster {
                presentation: self.presentation.clone_ref(py),
                index,
                path: self.path.clone(),
            },
        )
    }
}

#[pymethods]
impl PySlideMasterCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        sequence_item(py, key, self.len(py)?, "slide master", |index| {
            self.item(py, index)
        })
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let items = (0..self.len(py)?)
            .map(|index| self.item(py, index))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, items)?
            .into_any()
            .try_iter()
            .map(Bound::into_any)
    }
}

/// One master's theme, read and written live: name, colour scheme and fonts.
#[pyclass(name = "Theme")]
pub struct PyTheme {
    presentation: Py<PyPresentation>,
    master: usize,
    path: ContentPath,
}

impl PyTheme {
    fn read(&self, py: Python<'_>) -> PyResult<rpptx::Theme> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "theme",
            &format!(".slide_masters[{}].theme", self.master),
        )?;
        presentation
            .inner
            .theme(self.master)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("slide master index out of range"))
    }

    fn handle<T: PyClass + Into<pyo3::PyClassInitializer<T>>>(
        &self,
        py: Python<'_>,
        build: impl FnOnce(Py<PyPresentation>, usize, ContentPath) -> T,
    ) -> PyResult<Py<T>> {
        self.read(py)?;
        Py::new(
            py,
            build(
                self.presentation.clone_ref(py),
                self.master,
                self.path.clone(),
            ),
        )
    }
}

#[pymethods]
impl PyTheme {
    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.read(py)?.name)
    }

    /// `dk1`, `lt1`, `dk2`, `lt2`, `accent1` to `accent6`, `hlink` and
    /// `folHlink` mapped to an `RGBColor`, or `None` for a colour rpptx
    /// cannot resolve. Assigning a colour changes the theme, and every
    /// shape and text that uses it follows.
    #[getter]
    fn colors(&self, py: Python<'_>) -> PyResult<Py<PyThemeColors>> {
        self.handle(py, |presentation, master, path| PyThemeColors {
            presentation,
            master,
            path,
        })
    }

    #[getter]
    fn fonts(&self, py: Python<'_>) -> PyResult<Py<PyThemeFonts>> {
        self.handle(py, |presentation, master, path| PyThemeFonts {
            presentation,
            master,
            path,
        })
    }
}

/// The theme's colour scheme as a live mapping from slot name to `RGBColor`.
#[pyclass(name = "ThemeColors", mapping)]
pub struct PyThemeColors {
    presentation: Py<PyPresentation>,
    master: usize,
    path: ContentPath,
}

impl PyThemeColors {
    fn theme(&self, py: Python<'_>) -> PyTheme {
        PyTheme {
            presentation: self.presentation.clone_ref(py),
            master: self.master,
            path: self.path.clone(),
        }
    }
}

#[pymethods]
impl PyThemeColors {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.theme(py).read(py)?.colors.len())
    }

    fn __contains__(&self, py: Python<'_>, key: &str) -> PyResult<bool> {
        Ok(self
            .theme(py)
            .read(py)?
            .colors
            .iter()
            .any(|(slot, _)| *slot == key))
    }

    fn __getitem__(&self, py: Python<'_>, key: &str) -> PyResult<Py<PyAny>> {
        let theme = self.theme(py).read(py)?;
        let (_, color) = theme
            .colors
            .iter()
            .find(|(slot, _)| *slot == key)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(key.to_owned()))?;
        match color {
            Some(color) => {
                let [red, green, blue] = color.components();
                Ok(py
                    .import("rpptx.dml.color")?
                    .getattr("RGBColor")?
                    .call1((red, green, blue))?
                    .unbind())
            }
            None => Ok(py.None()),
        }
    }

    /// Sets one scheme colour from an `RGBColor`, a hex string with or
    /// without `#`, or an `(r, g, b)` triple.
    fn __setitem__(&self, py: Python<'_>, key: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        self.theme(py).read(py)?;
        let color = crate::dml::color_argument(value, "theme colour")?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_theme_color(self.master, key, color)
            .map_err(|error| pyo3::exceptions::PyKeyError::new_err(error.to_string()))
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.keys(py)?.into_any().try_iter().map(Bound::into_any)
    }

    fn keys<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let theme = self.theme(py).read(py)?;
        PyList::new(py, theme.colors.iter().map(|(slot, _)| *slot))
    }

    fn items<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let keys = self.keys(py)?;
        let items = PyList::empty(py);
        for key in keys.iter() {
            let key = key.extract::<String>()?;
            let value = self.__getitem__(py, &key)?;
            items.append((key, value))?;
        }
        Ok(items)
    }
}

/// The theme's heading (`major`) and body (`minor`) fonts.
#[pyclass(name = "ThemeFonts")]
pub struct PyThemeFonts {
    presentation: Py<PyPresentation>,
    master: usize,
    path: ContentPath,
}

impl PyThemeFonts {
    fn set(&self, py: Python<'_>, role: rpptx::ThemeFontRole) -> PyResult<Py<PyThemeFontSet>> {
        Py::new(
            py,
            PyThemeFontSet {
                presentation: self.presentation.clone_ref(py),
                master: self.master,
                path: self.path.clone(),
                role,
            },
        )
    }
}

#[pymethods]
impl PyThemeFonts {
    #[getter]
    fn major(&self, py: Python<'_>) -> PyResult<Py<PyThemeFontSet>> {
        self.set(py, rpptx::ThemeFontRole::Major)
    }

    #[getter]
    fn minor(&self, py: Python<'_>) -> PyResult<Py<PyThemeFontSet>> {
        self.set(py, rpptx::ThemeFontRole::Minor)
    }
}

/// One theme font collection's typefaces, empty when the theme names none.
/// Assigning a typeface changes the theme.
#[pyclass(name = "ThemeFontSet")]
pub struct PyThemeFontSet {
    presentation: Py<PyPresentation>,
    master: usize,
    path: ContentPath,
    role: rpptx::ThemeFontRole,
}

impl PyThemeFontSet {
    fn read(&self, py: Python<'_>) -> PyResult<rpptx::ThemeFonts> {
        let theme = PyTheme {
            presentation: self.presentation.clone_ref(py),
            master: self.master,
            path: self.path.clone(),
        }
        .read(py)?;
        Ok(match self.role {
            rpptx::ThemeFontRole::Major => theme.major_font,
            rpptx::ThemeFontRole::Minor => theme.minor_font,
        })
    }

    fn write(&self, py: Python<'_>, script: rpptx::ThemeFontScript, value: &str) -> PyResult<()> {
        self.read(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_theme_font(self.master, self.role, script, value)
            .map_err(|error| PyValueError::new_err(error.to_string()))
    }
}

#[pymethods]
impl PyThemeFontSet {
    #[getter]
    fn latin(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self.read(py)?.latin)
    }

    #[setter]
    fn set_latin(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.write(py, rpptx::ThemeFontScript::Latin, value)
    }

    #[getter]
    fn east_asian(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self.read(py)?.east_asian)
    }

    #[setter]
    fn set_east_asian(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.write(py, rpptx::ThemeFontScript::EastAsian, value)
    }

    #[getter]
    fn complex_script(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self.read(py)?.complex_script)
    }

    #[setter]
    fn set_complex_script(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.write(py, rpptx::ThemeFontScript::ComplexScript, value)
    }
}
