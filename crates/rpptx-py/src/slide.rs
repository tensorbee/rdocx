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
use crate::shape::{PyPlaceholderCollection, PyShapeCollection};
use crate::validate_path;

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PySlideMaster>()?;
    module.add_class::<PySlideMasterCollection>()?;
    module.add_class::<PyTheme>()?;
    module.add_class::<PyThemeColors>()?;
    module.add_class::<PyThemeFonts>()?;
    module.add_class::<PyThemeFontSet>()?;
    module.add_class::<PyHeaderFooter>()?;
    module.add_class::<PySlideTransition>()?;
    module.add_class::<PySlideLayout>()?;
    module.add_class::<PySlideLayoutCollection>()?;
    module.add_class::<PySlide>()?;
    module.add_class::<PyMediaInfo>()?;
    module.add_class::<PySlideCollection>()?;
    module.add_class::<PyBackground>()?;
    Ok(())
}

#[pyclass(name = "SlideLayout")]
pub struct PySlideLayout {
    pub(crate) presentation: Py<PyPresentation>,
    pub(crate) index: usize,
    path: ContentPath,
}

impl PySlideLayout {
    fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "slide layout",
            &format!(".slide_layouts[{}]", self.index),
        )?;
        presentation
            .current_layout_index(&self.path)
            .ok_or_else(|| PyIndexError::new_err("slide layout owner is absent"))
    }
}

#[pymethods]
impl PySlideLayout {
    #[getter]
    fn element(&self, py: Python<'_>) -> PyResult<Py<crate::presentation::PyPptElement>> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_path(smallvec![PathSeg::Layout(index)], None)?;
        crate::presentation::PyPptElement::new_checked(
            py,
            self.presentation.clone_ref(py),
            path,
            crate::presentation::PptElementKind::Layout,
        )
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, name: Option<&str>) -> PyResult<()> {
        let index = self.validate(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .set_layout_name(index, name)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    #[getter]
    fn slide_master(&self, py: Python<'_>) -> PyResult<Py<PySlideMaster>> {
        let index = self.validate(py)?;
        let master = self
            .presentation
            .borrow(py)
            .inner
            .layout_master_index(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        PySlideMasterCollection::new(
            self.presentation.clone_ref(py),
            self.presentation
                .borrow(py)
                .capture_path(smallvec![], None)?,
        )
        .item(py, master)
    }

    #[getter]
    fn used_by_slides<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let index = self.validate(py)?;
        let indices = self
            .presentation
            .borrow(py)
            .inner
            .layout_used_by_slides(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let items = indices
            .into_iter()
            .map(|index| {
                let path = self
                    .presentation
                    .borrow(py)
                    .capture_path(smallvec![PathSeg::Slide(index)], None)?;
                Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, items)
    }

    #[getter]
    fn background(&self, py: Python<'_>) -> PyResult<Py<PyBackground>> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(smallvec![PathSeg::Layout(index)])?;
        Py::new(
            py,
            PyBackground {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
    }

    #[getter]
    fn follow_master_background(&self, py: Python<'_>) -> PyResult<bool> {
        let index = self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .has_explicit_background_at(rpptx::ShapeOwner::Layout(index))
            .map(|explicit| !explicit)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    #[setter]
    fn set_follow_master_background(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_follow_master_background_at(rpptx::ShapeOwner::Layout(index), value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideLayout.follow_master_background",
            &proofs,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }

    #[getter]
    fn show_master_shapes(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        let index = self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .show_master_shapes_at(rpptx::ShapeOwner::Layout(index))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    #[setter]
    fn set_show_master_shapes(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut p = self.presentation.borrow_mut(py);
        let mut candidate = p.inner.clone();
        candidate
            .set_show_master_shapes_at(rpptx::ShapeOwner::Layout(index), value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = p.unchanged_owners()?;
        let prepared = p.prepare_owners(
            &candidate,
            "SlideLayout.show_master_shapes",
            &proofs,
            false,
            false,
            false,
        )?;
        p.publish_owners(candidate, prepared)
    }

    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(smallvec![PathSeg::Layout(index)])?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), path),
        )
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<PyShapeCollection> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(smallvec![PathSeg::Layout(index)])?;
        Ok(PyShapeCollection::new(
            self.presentation.clone_ref(py),
            path,
        ))
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let index = self.validate(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .layout_name(index)
            .map(str::to_owned))
    }

    /// Two handles are equal when they name the same layout of one presentation.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, PySlideLayout>>()
            .is_ok_and(|other| {
                other.presentation.is(&self.presentation)
                    && other.path.owner_generation() == self.path.owner_generation()
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
                path: self.presentation.borrow(py).capture_layout(index)?,
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
        layout.validate(py)
    }

    /// Duplicates one layout into a fresh part owned by the same master.
    fn duplicate(&self, py: Python<'_>, layout: &Bound<'_, PyAny>) -> PyResult<Py<PySlideLayout>> {
        self.len(py)?;
        let layout = layout.extract::<PyRef<'_, PySlideLayout>>()?;
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("layout is not in this presentation"));
        }
        let index = layout.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        let inserted = candidate
            .duplicate_layout(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideLayoutCollection.duplicate",
            &proofs,
            false,
            true,
            false,
        )?;
        let path = presentation.capture_prepared_path(
            &prepared,
            smallvec![PathSeg::Layout(inserted)],
            None,
        )?;
        let result = Py::new(
            py,
            PySlideLayout {
                presentation: self.presentation.clone_ref(py),
                index: inserted,
                path,
            },
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    /// Removes an unused layout only after the complete inbound graph is checked.
    fn remove(&self, py: Python<'_>, layout: &Bound<'_, PyAny>) -> PyResult<()> {
        self.len(py)?;
        let layout = layout.extract::<PyRef<'_, PySlideLayout>>()?;
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("layout is not in this presentation"));
        }
        let index = layout.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .remove_layout(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = presentation.removed_layout_owners(index)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideLayoutCollection.remove",
            &proofs,
            false,
            true,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
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

/// An immutable media inventory snapshot in slide draw order.
#[pyclass(name = "MediaInfo", frozen, get_all)]
pub struct PyMediaInfo {
    slide_index: usize,
    shape_id: u32,
    kind: String,
    part_name: Option<String>,
    content_type: Option<String>,
    target: Option<String>,
    poster_relationship_id: Option<String>,
    trim_start_ms: Option<u64>,
    trim_end_ms: Option<u64>,
    volume: u32,
    looped: bool,
    show_when_stopped: bool,
    trigger: &'static str,
    diagnostics: Py<PyTuple>,
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
        presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?
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
    #[getter]
    fn slide_id(&self, py: Python<'_>) -> PyResult<u32> {
        let index = self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .slide(index)
            .map(|slide| slide.id())
            .ok_or_else(|| PyIndexError::new_err("slide owner is absent"))
    }

    #[getter]
    fn element(&self, py: Python<'_>) -> PyResult<Py<crate::presentation::PyPptElement>> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_path(smallvec![PathSeg::Slide(index)], None)?;
        crate::presentation::PyPptElement::new_checked(
            py,
            self.presentation.clone_ref(py),
            path,
            crate::presentation::PptElementKind::Slide,
        )
    }

    #[getter]
    fn media<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let index = self.validate(py)?;
        let inventory = self
            .presentation
            .borrow(py)
            .inner
            .media(index)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?;
        let mut result = Vec::with_capacity(inventory.len());
        for info in inventory {
            let (part_name, content_type, target) = match info.source {
                rpptx::MediaLocation::Embedded {
                    part_name,
                    content_type,
                } => (Some(part_name), Some(content_type), None),
                rpptx::MediaLocation::Linked { target } => (None, None, Some(target)),
            };
            result.push(Py::new(
                py,
                PyMediaInfo {
                    slide_index: info.slide_index,
                    shape_id: info.shape_id,
                    kind: if info.kind == rpptx::MediaKind::Audio {
                        "audio"
                    } else {
                        "video"
                    }
                    .into(),
                    part_name,
                    content_type,
                    target,
                    poster_relationship_id: info.poster_relationship_id,
                    trim_start_ms: info.settings.trim_start_ms,
                    trim_end_ms: info.settings.trim_end_ms,
                    volume: info.settings.volume,
                    looped: info.settings.looped,
                    show_when_stopped: info.settings.show_when_stopped,
                    trigger: match info.settings.trigger {
                        rpptx::MediaPlaybackTrigger::Automatic => "automatic",
                        rpptx::MediaPlaybackTrigger::OnClick => "on_click",
                        rpptx::MediaPlaybackTrigger::InClickSequence => "in_click_sequence",
                    },
                    diagnostics: PyTuple::new(
                        py,
                        info.diagnostics.into_iter().map(|d| match d {
                            rpptx::MediaDiagnostic::UnsupportedContentType { content_type } => {
                                ("unsupported_content_type", Some(content_type))
                            }
                            rpptx::MediaDiagnostic::MissingPlaybackTiming => {
                                ("missing_playback_timing", None)
                            }
                            rpptx::MediaDiagnostic::UnsupportedPlaybackCommand { command } => {
                                ("unsupported_playback_command", Some(command))
                            }
                            rpptx::MediaDiagnostic::MissingPoster => ("missing_poster", None),
                        }),
                    )?
                    .unbind(),
                },
            )?);
        }
        PyTuple::new(py, result)
    }

    fn extract_media(&self, py: Python<'_>, shape_id: u32) -> PyResult<Option<Py<PyBytes>>> {
        let index = self.validate(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .extract_media(index, shape_id)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
            .map(|bytes| PyBytes::new(py, &bytes).unbind()))
    }

    #[pyo3(signature = (shape_id, media_file, mime_type=None))]
    fn replace_media(
        &self,
        py: Python<'_>,
        shape_id: u32,
        media_file: &Bound<'_, PyAny>,
        mime_type: Option<&str>,
    ) -> PyResult<()> {
        let index = self.validate(py)?;
        let media = crate::shape::MediaFile::read(media_file, mime_type)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .replace_media(index, shape_id, media.input())
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.replace_media",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }

    fn remove_media(&self, py: Python<'_>, shape_id: u32) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let shape_index = presentation
            .inner
            .slide(index)
            .and_then(|slide| {
                slide
                    .shapes()
                    .position(|shape| shape.non_visual_id() == Some(shape_id))
            })
            .ok_or_else(|| PyValueError::new_err("media shape is absent"))?;
        let mut candidate = presentation.inner.clone();
        candidate
            .remove_media(index, shape_id)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?;
        let survivors = presentation.removed_shape_owners(&[PathSeg::Slide(index)], shape_index)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.remove_media",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }

    #[getter]
    fn transition(&self, py: Python<'_>) -> PyResult<PySlideTransition> {
        self.validate(py)?;
        Ok(PySlideTransition {
            presentation: self.presentation.clone_ref(py),
            path: self
                .presentation
                .borrow(py)
                .capture_xml_children_path(smallvec![PathSeg::Slide(self.validate(py)?)])?,
        })
    }
    #[getter]
    fn header_footer(&self, py: Python<'_>) -> PyResult<Py<PyHeaderFooter>> {
        self.validate(py)?;
        Py::new(
            py,
            PyHeaderFooter {
                presentation: self.presentation.clone_ref(py),
                path: self
                    .presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Slide(self.validate(py)?)])?,
            },
        )
    }

    /// Two current handles are equal when they name the same slide of one
    /// presentation. Like python-pptx `Slide`, a handle is not hashable.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other.extract::<PyRef<'_, PySlide>>().is_ok_and(|other| {
            other.presentation.is(&self.presentation)
                && other.path.owner_generation() == self.path.owner_generation()
        })
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<Py<PyShapeCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeCollection::new(
                self.presentation.clone_ref(py),
                self.presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Slide(self.validate(py)?)])?,
            ),
        )
    }

    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyPlaceholderCollection::new(
                self.presentation.clone_ref(py),
                self.presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Slide(self.validate(py)?)])?,
            ),
        )
    }

    #[getter]
    fn has_notes_slide(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.notes_text(py)?.is_some())
    }

    #[getter]
    fn notes_slide(&self, py: Python<'_>) -> PyResult<Py<PyNotesSlide>> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        if presentation
            .inner
            .slide(index)
            .is_some_and(|slide| slide.notes_text().is_some())
        {
            return Py::new(
                py,
                PyNotesSlide {
                    presentation: self.presentation.clone_ref(py),
                    path: self.path.clone(),
                },
            );
        }
        let mut candidate = presentation.inner.clone();
        candidate
            .set_notes_text(index, "")
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.text_changed_owners(&self.path)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.notes_slide",
            &survivors,
            false,
            false,
            false,
        )?;
        let path = presentation.capture_prepared_path(
            &prepared,
            smallvec![PathSeg::Slide(index)],
            None,
        )?;
        let result = Py::new(
            py,
            PyNotesSlide {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
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
        let mut candidate = presentation.inner.clone();
        candidate
            .set_notes_text(index, text)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.text_changed_owners(&self.path)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.notes_text",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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
        let mut candidate = presentation.inner.clone();
        let inner = &mut candidate;
        let count = py
            .detach(|| inner.try_replace_slide_text(index, placeholder, replacement, notes, expect))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        if let Some(expected) = expect.filter(|&expected| expected != count) {
            return Err(replacement_count_to_pyerr(py, placeholder, expected, count));
        }
        if count > 0 {
            let survivors = presentation.replaced_text_owners(&candidate)?;
            let prepared = presentation.prepare_owners(
                &candidate,
                "Slide.try_replace_text",
                &survivors,
                false,
                false,
                false,
            )?;
            presentation.publish_owners(candidate, prepared)?;
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
        let path = presentation.capture_layout(layout)?;
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
        let layout_index = layout.validate(py)?;
        if self.presentation.borrow(py).inner.slide_layout_index(index) == Some(layout_index) {
            return Ok(());
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_slide_layout(index, layout_index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.slide_layout",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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
                path: self
                    .presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Slide(self.validate(py)?)])?,
            },
        )
    }

    #[getter]
    fn show_master_shapes(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        let index = self.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .show_master_shapes_at(rpptx::ShapeOwner::Slide(index))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }

    #[setter]
    fn set_show_master_shapes(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut p = self.presentation.borrow_mut(py);
        let mut candidate = p.inner.clone();
        candidate
            .set_show_master_shapes_at(rpptx::ShapeOwner::Slide(index), value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = p.unchanged_owners()?;
        let prepared = p.prepare_owners(
            &candidate,
            "Slide.show_master_shapes",
            &proofs,
            false,
            false,
            false,
        )?;
        p.publish_owners(candidate, prepared)
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
        let mut candidate = presentation.inner.clone();
        candidate
            .set_follow_master_background_at(rpptx::ShapeOwner::Slide(index), value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.follow_master_background",
            &proofs,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
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
        let mut candidate = presentation.inner.clone();
        match shape_id {
            Some(shape_id) => candidate.add_comment_at_shape(
                index,
                comment,
                shape_id,
                text_start.zip(text_length),
            ),
            None => candidate.add_comment(index, comment),
        }
        .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.add_comment",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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
        let mut candidate = presentation.inner.clone();
        candidate
            .reply_to_comment(index, comment_id, reply)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.reply_to_comment",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(())
    }

    /// Marks one comment thread resolved. A reply id is an unknown id.
    fn resolve_comment(&self, comment_id: &str, py: Python<'_>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .resolve_comment(index, comment_id)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.resolve_comment",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(())
    }

    /// Removes one comment thread with its replies, or one reply.
    fn remove_comment(&self, comment_id: &str, py: Python<'_>) -> PyResult<()> {
        let index = self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .remove_comment(index, comment_id)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.remove_comment",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(())
    }

    fn move_comment(&self, from_: usize, to: usize, py: Python<'_>) -> PyResult<()> {
        let index = self.validate(py)?;
        if from_ == to {
            return Ok(());
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .move_comment(index, from_, to)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.move_comment",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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
        if from_ == to {
            return Ok(());
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .move_reply(index, comment_id, from_, to)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Slide.move_reply",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(())
    }
}

#[pyclass(name = "SlideCollection")]
pub struct PySlideCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

struct SlideImportRequest {
    source: rpptx::Presentation,
    indexes: Vec<usize>,
    options: rpptx::ImportSlideOptions,
    slide_map: Vec<(usize, usize)>,
    operation: &'static str,
    single: bool,
}

struct ImportedSlides {
    slides: Vec<Py<PySlide>>,
    output: Py<PyTuple>,
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
            .capture_path(smallvec![PathSeg::Slide(index)], None)?;
        Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))
    }
    fn import_options(
        &self,
        py: Python<'_>,
        layout: Option<&Bound<'_, PyAny>>,
        index: Option<isize>,
        keep_source_formatting: bool,
        keep_sections: bool,
    ) -> PyResult<rpptx::ImportSlideOptions> {
        let layout = layout
            .map(|layout| layout.extract::<PyRef<'_, PySlideLayout>>())
            .transpose()?;
        if layout
            .as_ref()
            .is_some_and(|layout| !layout.presentation.is(&self.presentation))
        {
            return Err(PyValueError::new_err(
                "layout is not a layout of this presentation",
            ));
        }
        let len = self.len(py)?;
        let layout_index = layout.map(|layout| layout.validate(py)).transpose()?;
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
        Ok(rpptx::ImportSlideOptions {
            layout_index,
            insert_index: Some(inserted),
            keep_source_formatting,
            keep_sections,
        })
    }

    fn single_import_request(
        &self,
        py: Python<'_>,
        slide: &Bound<'_, PyAny>,
        layout: Option<&Bound<'_, PyAny>>,
        index: Option<isize>,
        keep_source_formatting: bool,
        operation: &'static str,
    ) -> PyResult<SlideImportRequest> {
        let mut options = self.import_options(py, layout, index, keep_source_formatting, false)?;
        let slide = slide.extract::<PyRef<'_, PySlide>>()?;
        let source_index = slide.validate(py)?;
        let source = slide.presentation.borrow(py).inner.clone();
        if slide.presentation.is(&self.presentation)
            && options.layout_index.is_none()
            && !keep_source_formatting
        {
            options.layout_index = source.slide_layout_index(source_index);
        }
        Ok(SlideImportRequest {
            source,
            indexes: vec![source_index],
            options,
            slide_map: Vec::new(),
            operation,
            single: true,
        })
    }

    fn publish_import(
        &self,
        py: Python<'_>,
        request: SlideImportRequest,
    ) -> PyResult<ImportedSlides> {
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        let (indexes, diagnostics) = candidate
            .import_slides(
                &request.source,
                &request.indexes,
                &request.options,
                &request.slide_map,
            )
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let diagnostics = PyTuple::new(py, diagnostics)?;
        if indexes.is_empty() {
            let slides = PyTuple::empty(py);
            let output = PyTuple::new(py, [slides.as_any(), diagnostics.as_any()])?.unbind();
            return Ok(ImportedSlides {
                slides: Vec::new(),
                output,
            });
        }
        let survivors =
            presentation.slide_owner_edit(crate::presentation::SlideOwnerEdit::InsertMany {
                at: indexes[0],
                count: indexes.len(),
            })?;
        let prepared = presentation.prepare_owners(
            &candidate,
            request.operation,
            &survivors,
            true,
            candidate.layout_count() != presentation.inner.layout_count(),
            candidate.master_count() != presentation.inner.master_count(),
        )?;
        let mut slides = Vec::new();
        slides.try_reserve_exact(indexes.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("imported Slide allocation failed")
        })?;
        for index in indexes {
            let path = presentation.capture_prepared_path(
                &prepared,
                smallvec![PathSeg::Slide(index)],
                None,
            )?;
            slides.push(Py::new(
                py,
                PySlide::new(self.presentation.clone_ref(py), path),
            )?);
        }
        let first = if request.single {
            slides[0].bind(py).as_any().clone()
        } else {
            PyTuple::new(py, slides.iter().map(|slide| slide.bind(py)))?.into_any()
        };
        let output = PyTuple::new(py, [first.as_any(), diagnostics.as_any()])?.unbind();
        presentation.publish_owners(candidate, prepared)?;
        Ok(ImportedSlides { slides, output })
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
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "layout belongs to a different presentation",
            ));
        }
        self.len(py)?;
        let layout_index = layout.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let index = presentation.inner.len();
        let mut candidate = presentation.inner.clone();
        candidate
            .add_slide(layout_index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideCollection.add_slide",
            &survivors,
            true,
            false,
            false,
        )?;
        let path = presentation.capture_prepared_path(
            &prepared,
            smallvec![PathSeg::Slide(index)],
            None,
        )?;
        let result = Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    /// Inserts a new slide at a final zero-based position, including append.
    fn add_slide_at(
        &self,
        py: Python<'_>,
        layout: &Bound<'_, PyAny>,
        position: usize,
    ) -> PyResult<Py<PySlide>> {
        let layout = layout.extract::<PyRef<'_, PySlideLayout>>()?;
        if !layout.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "layout belongs to a different presentation",
            ));
        }
        self.len(py)?;
        let layout_index = layout.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        let index = candidate
            .add_slide_at(layout_index, position)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors =
            presentation.slide_owner_edit(crate::presentation::SlideOwnerEdit::Insert(index))?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideCollection.add_slide_at",
            &survivors,
            true,
            false,
            false,
        )?;
        let path = presentation.capture_prepared_path(
            &prepared,
            smallvec![PathSeg::Slide(index)],
            None,
        )?;
        let result = Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    /// Imports one slide, appending by default, with optional source formatting.
    #[pyo3(signature = (slide, layout=None, index=None, *, keep_source_formatting=false))]
    fn import_slide(
        &self,
        py: Python<'_>,
        slide: &Bound<'_, PyAny>,
        layout: Option<&Bound<'_, PyAny>>,
        index: Option<isize>,
        keep_source_formatting: bool,
    ) -> PyResult<Py<PySlide>> {
        let request = self.single_import_request(
            py,
            slide,
            layout,
            index,
            keep_source_formatting,
            "SlideCollection.import_slide",
        )?;
        let mut result = self.publish_import(py, request)?;
        Ok(result
            .slides
            .pop()
            .expect("single import allocated one Slide"))
    }

    /// Imports one slide and returns every dropped-jump diagnostic.
    #[pyo3(signature = (slide, layout=None, index=None, *, keep_source_formatting=false))]
    fn import_slide_with_diagnostics(
        &self,
        py: Python<'_>,
        slide: &Bound<'_, PyAny>,
        layout: Option<&Bound<'_, PyAny>>,
        index: Option<isize>,
        keep_source_formatting: bool,
    ) -> PyResult<Py<PyTuple>> {
        let request = self.single_import_request(
            py,
            slide,
            layout,
            index,
            keep_source_formatting,
            "SlideCollection.import_slide_with_diagnostics",
        )?;
        Ok(self.publish_import(py, request)?.output)
    }

    /// Imports selected source indexes in caller order, or all source slides.
    /// Slide maps contain source and existing destination indexes before insertion.
    #[pyo3(signature = (source, indexes=None, *, layout=None, index=None, keep_source_formatting=false, keep_sections=false, slide_map=None))]
    // The approved Python signature exposes these independent keyword options.
    #[allow(clippy::too_many_arguments)]
    fn import_slides(
        &self,
        py: Python<'_>,
        source: &Bound<'_, PyAny>,
        indexes: Option<Vec<isize>>,
        layout: Option<&Bound<'_, PyAny>>,
        index: Option<isize>,
        keep_source_formatting: bool,
        keep_sections: bool,
        slide_map: Option<Vec<(isize, isize)>>,
    ) -> PyResult<Py<PyTuple>> {
        let options =
            self.import_options(py, layout, index, keep_source_formatting, keep_sections)?;
        let destination_len = self.len(py)?;
        let source = source.extract::<PyRef<'_, PyPresentation>>()?.inner.clone();
        let indexes = match indexes {
            None => (0..source.len()).collect(),
            Some(indexes) => indexes
                .into_iter()
                .map(|index| normalize_index(index, source.len(), "source slide"))
                .collect::<PyResult<Vec<_>>>()?,
        };
        let slide_map = slide_map
            .unwrap_or_default()
            .into_iter()
            .map(|(from, to)| {
                Ok((
                    normalize_index(from, source.len(), "source slide")?,
                    normalize_index(to, destination_len, "destination slide")?,
                ))
            })
            .collect::<PyResult<Vec<_>>>()?;
        let request = SlideImportRequest {
            source,
            indexes,
            options,
            slide_map,
            operation: "SlideCollection.import_slides",
            single: false,
        };
        Ok(self.publish_import(py, request)?.output)
    }

    #[getter(_sldIdLst)]
    fn unsupported_slide_xml_list(&self, py: Python<'_>) -> PyResult<()> {
        self.len(py)?;
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "use prs.slides.remove(slide) to delete a slide and prs.slides.move(old_index, new_index) to reorder",
        ))
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
        let mut candidate = presentation.inner.clone();
        candidate
            .remove_slide(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors =
            presentation.slide_owner_edit(crate::presentation::SlideOwnerEdit::Remove(index))?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideCollection.remove",
            &survivors,
            true,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .duplicate_slide(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation
            .slide_owner_edit(crate::presentation::SlideOwnerEdit::Insert(index + 1))?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideCollection.duplicate",
            &survivors,
            true,
            false,
            false,
        )?;
        let path = presentation.capture_prepared_path(
            &prepared,
            smallvec![PathSeg::Slide(index + 1)],
            None,
        )?;
        let result = Py::new(py, PySlide::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    /// Moves the slide at `from_` so that it ends up at index `to`.
    #[pyo3(name = "move")]
    fn move_slide(&self, py: Python<'_>, from_: isize, to: isize) -> PyResult<()> {
        let len = self.len(py)?;
        let from_ = normalize_index(from_, len, "slide")?;
        let to = normalize_index(to, len, "slide")?;
        if from_ == to {
            return Ok(());
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .move_slide(from_, to)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let survivors = presentation
            .slide_owner_edit(crate::presentation::SlideOwnerEdit::Move { from: from_, to })?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "SlideCollection.move",
            &survivors,
            true,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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

#[pyclass(name = "SlideMaster")]
pub struct PySlideMaster {
    presentation: Py<PyPresentation>,
    part_name: String,
    path: ContentPath,
}

impl PySlideMaster {
    fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let p = self.presentation.borrow(py);
        validate_path(py, &p, &self.path, "slide master", ".slide_masters")?;
        for index in 0..p.inner.slide_master_count() {
            if p.inner
                .slide_master_part_name(index)
                .map_err(|e| crate::rpptx_to_pyerr(py, e))?
                .as_deref()
                == Some(self.part_name.as_str())
            {
                return Ok(index);
            }
        }
        Err(PyIndexError::new_err("slide master owner is absent"))
    }
}

#[pymethods]
impl PySlideMaster {
    #[getter]
    fn element(&self, py: Python<'_>) -> PyResult<Py<crate::presentation::PyPptElement>> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_path(smallvec![PathSeg::Master(index)], None)?;
        crate::presentation::PyPptElement::new_checked(
            py,
            self.presentation.clone_ref(py),
            path,
            crate::presentation::PptElementKind::Master,
        )
    }

    #[getter]
    fn text_styles(&self, py: Python<'_>) -> PyResult<Py<crate::text::PyMasterTextStyles>> {
        self.validate(py)?;
        Py::new(
            py,
            crate::text::PyMasterTextStyles {
                presentation: self.presentation.clone_ref(py),
                path: self
                    .presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Master(self.validate(py)?)])?,
            },
        )
    }

    #[getter]
    fn background(&self, py: Python<'_>) -> PyResult<Py<PyBackground>> {
        self.validate(py)?;
        Py::new(
            py,
            PyBackground {
                presentation: self.presentation.clone_ref(py),
                path: self
                    .presentation
                    .borrow(py)
                    .capture_xml_children_path(smallvec![PathSeg::Master(self.validate(py)?)])?,
            },
        )
    }

    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(smallvec![PathSeg::Master(index)])?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), path),
        )
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<PyShapeCollection> {
        let index = self.validate(py)?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(smallvec![PathSeg::Master(index)])?;
        Ok(PyShapeCollection::new(
            self.presentation.clone_ref(py),
            path,
        ))
    }

    /// The theme this master uses, read when accessed.
    #[getter]
    fn theme(&self, py: Python<'_>) -> PyResult<PyTheme> {
        self.validate(py)?;
        Ok(PyTheme {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self
                .presentation
                .borrow(py)
                .capture_xml_children_path(smallvec![PathSeg::Master(self.validate(py)?)])?,
        })
    }

    /// The layouts this master owns, as `prs.slide_layouts` members.
    #[getter]
    fn slide_layouts<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let index = self.validate(py)?;
        let layouts = self
            .presentation
            .borrow(py)
            .inner
            .master_layout_indices(index)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let collection =
            PySlideLayoutCollection::new(self.presentation.clone_ref(py), self.path.clone());
        let items = layouts
            .into_iter()
            .map(|index| collection.item(py, index))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, items)
    }

    /// Two handles are equal when they name the same master of one presentation.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, PySlideMaster>>()
            .is_ok_and(|other| {
                other.presentation.is(&self.presentation) && other.part_name == self.part_name
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
        Ok(self.presentation.borrow(py).inner.slide_master_count())
    }

    pub(crate) fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PySlideMaster>> {
        Py::new(
            py,
            PySlideMaster {
                presentation: self.presentation.clone_ref(py),
                part_name: self
                    .presentation
                    .borrow(py)
                    .inner
                    .slide_master_part_name(index)
                    .map_err(|e| crate::rpptx_to_pyerr(py, e))?
                    .ok_or_else(|| PyIndexError::new_err("slide master index out of range"))?,
                path: self
                    .presentation
                    .borrow(py)
                    .capture_path(smallvec![PathSeg::Master(index)], None)?,
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

/// A live view that resolves its original master part on every access.
#[pyclass(name = "Theme")]
pub struct PyTheme {
    presentation: Py<PyPresentation>,
    part_name: String,
    path: ContentPath,
}

impl PyTheme {
    fn current(&self, py: Python<'_>) -> PyResult<rpptx::Theme> {
        let master = PySlideMaster {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        };
        let index = master.validate(py)?;
        self.presentation
            .borrow(py)
            .inner
            .theme(index)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
            .ok_or_else(|| PyIndexError::new_err("slide master owner is absent"))
    }
}

#[pymethods]
impl PyTheme {
    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.current(py)?.name)
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let master = PySlideMaster {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        };
        let index = master.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_theme_name(index, value)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        if candidate
            .saves_as(&presentation.inner)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?
        {
            return Ok(());
        }
        let proofs = presentation.unchanged_owners()?;
        let prepared =
            presentation.prepare_owners(&candidate, "Theme.name", &proofs, false, false, false)?;
        presentation.publish_owners(candidate, prepared)
    }

    #[getter]
    fn colors(&self, py: Python<'_>) -> PyResult<PyThemeColors> {
        self.current(py)?;
        Ok(PyThemeColors {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        })
    }

    #[getter]
    fn fonts(&self, py: Python<'_>) -> PyResult<PyThemeFonts> {
        self.current(py)?;
        Ok(PyThemeFonts {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        })
    }
}

/// Live access to the twelve concrete color-scheme slots of the owning theme.
#[pyclass(name = "ThemeColors")]
pub struct PyThemeColors {
    presentation: Py<PyPresentation>,
    part_name: String,
    path: ContentPath,
}

#[pymethods]
impl PyThemeColors {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.theme(py).current(py)?.colors.len())
    }

    fn __getitem__(&self, py: Python<'_>, key: &str) -> PyResult<Py<PyAny>> {
        let current = self.theme(py).current(py)?;
        let color = current
            .colors
            .into_iter()
            .find(|(slot, _)| *slot == key)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(key.to_owned()))?
            .1;
        match color {
            Some(color) => {
                let [r, g, b] = color.components();
                py.import("rpptx.dml.color")?
                    .getattr("RGBColor")?
                    .call1((r, g, b))
                    .map(Bound::unbind)
            }
            None => Ok(py.None()),
        }
    }

    fn __setitem__(&self, py: Python<'_>, key: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let master = PySlideMaster {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        };
        let index = master.validate(py)?;
        let (r, g, b) = oxml_py_support::rgb_channels(value, "Theme.colors")?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_theme_color(index, key, rpptx::RgbColor::new(r, g, b))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Theme.colors",
            &proofs,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }

    fn keys<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(
            py,
            self.theme(py)
                .current(py)?
                .colors
                .into_iter()
                .map(|(slot, _)| slot),
        )
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.keys(py)?.into_any().try_iter().map(Bound::into_any)
    }
}

impl PyThemeColors {
    fn theme(&self, py: Python<'_>) -> PyTheme {
        PyTheme {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        }
    }
}

/// The owning theme's heading and body font collections.
#[pyclass(name = "ThemeFonts")]
pub struct PyThemeFonts {
    presentation: Py<PyPresentation>,
    part_name: String,
    path: ContentPath,
}

#[pymethods]
impl PyThemeFonts {
    #[getter]
    fn major(&self, py: Python<'_>) -> PyThemeFontSet {
        self.collection(py, true)
    }
    #[getter]
    fn minor(&self, py: Python<'_>) -> PyThemeFontSet {
        self.collection(py, false)
    }
}

impl PyThemeFonts {
    fn collection(&self, py: Python<'_>, major: bool) -> PyThemeFontSet {
        PyThemeFontSet {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
            major,
        }
    }
}

/// Live typefaces in one theme collection, empty when the source names none.
#[pyclass(name = "ThemeFontSet")]
pub struct PyThemeFontSet {
    presentation: Py<PyPresentation>,
    part_name: String,
    path: ContentPath,
    major: bool,
}

impl PyThemeFontSet {
    fn current(&self, py: Python<'_>) -> PyResult<rpptx::ThemeFonts> {
        let theme = PyTheme {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        }
        .current(py)?;
        Ok(if self.major {
            theme.major_font
        } else {
            theme.minor_font
        })
    }

    fn write(&self, py: Python<'_>, script: &str, typeface: &str) -> PyResult<()> {
        let master = PySlideMaster {
            presentation: self.presentation.clone_ref(py),
            part_name: self.part_name.clone(),
            path: self.path.clone(),
        };
        let index = master.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_theme_font(index, self.major, script, typeface)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        let proofs = presentation.unchanged_owners()?;
        let prepared =
            presentation.prepare_owners(&candidate, "Theme.fonts", &proofs, false, false, false)?;
        presentation.publish_owners(candidate, prepared)
    }
}

#[pymethods]
impl PyThemeFontSet {
    #[getter]
    fn latin(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self.current(py)?.latin)
    }
    #[setter]
    fn set_latin(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.write(py, "latin", value)
    }
    #[getter]
    fn east_asian(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self.current(py)?.east_asian)
    }
    #[setter]
    fn set_east_asian(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.write(py, "east_asian", value)
    }
    #[getter]
    fn complex_script(&self, py: Python<'_>) -> PyResult<String> {
        Ok(self.current(py)?.complex_script)
    }
    #[setter]
    fn set_complex_script(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.write(py, "complex_script", value)
    }
}

pub(crate) fn header_footer_settings(
    slide_number: bool,
    footer: Option<String>,
    date: Option<String>,
    date_format: &str,
    date_cache: Option<String>,
) -> rpptx::HeaderFooter {
    let date = match date.as_deref() {
        None => rpptx::HeaderFooterDate::Off,
        Some("auto") => rpptx::HeaderFooterDate::Automatic {
            field_type: date_format.into(),
            text: date_cache,
        },
        Some(text) => rpptx::HeaderFooterDate::Fixed(text.into()),
    };
    rpptx::HeaderFooter {
        slide_number,
        footer,
        date,
    }
}

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

    /// Patch the latest owning slide and publish all causal scope changes once.
    fn update(
        &mut self,
        py: Python<'_>,
        change: impl FnOnce(&mut rpptx::HeaderFooter),
    ) -> PyResult<()> {
        let (index, mut settings) = self.current(py)?;
        change(&mut settings);
        let mut p = self.presentation.borrow_mut(py);
        let mut candidate = p.inner.clone();
        if candidate
            .set_slide_header_footer(index, &settings)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
        {
            let proofs = p.footer_owners(&candidate)?;
            let prepared = p.prepare_owners(
                &candidate,
                "Slide.header_footer",
                &proofs,
                false,
                false,
                false,
            )?;
            p.publish_owners(candidate, prepared)?;
        }
        Ok(())
    }
}

#[pymethods]
impl PyHeaderFooter {
    #[getter]
    fn date_cache(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(match self.current(py)?.1.date {
            rpptx::HeaderFooterDate::Automatic { text, .. } => text,
            _ => None,
        })
    }

    #[getter]
    fn date_mode(&self, py: Python<'_>) -> PyResult<&'static str> {
        Ok(match self.current(py)?.1.date {
            rpptx::HeaderFooterDate::Off => "off",
            rpptx::HeaderFooterDate::Fixed(_) => "fixed",
            rpptx::HeaderFooterDate::Automatic { .. } => "automatic",
        })
    }

    #[getter]
    fn source(&self, py: Python<'_>) -> PyResult<&'static str> {
        let current = self.current(py)?.1;
        if current.slide_number
            || current.footer.is_some()
            || !matches!(current.date, rpptx::HeaderFooterDate::Off)
        {
            return Ok("slide");
        }
        self.policy_source(py)
    }

    /// Legal inherited policy is independent of materialized slide-owned content.
    #[getter]
    fn policy_source(&self, py: Python<'_>) -> PyResult<&'static str> {
        let index = self.slide_index(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .slide_header_footer_policy(index)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
            .map_or("absent", |(source, _)| source))
    }

    #[getter]
    fn slide_number_policy(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        let index = self.slide_index(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .slide_header_footer_policy(index)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
            .and_then(|(_, flags)| flags.slide_number))
    }

    #[getter]
    fn footer_policy(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        let index = self.slide_index(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .slide_header_footer_policy(index)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
            .and_then(|(_, flags)| flags.footer))
    }

    #[getter]
    fn date_policy(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        let index = self.slide_index(py)?;
        Ok(self
            .presentation
            .borrow(py)
            .inner
            .slide_header_footer_policy(index)
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
            .and_then(|(_, flags)| flags.date))
    }

    #[getter]
    fn footer_present(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.current(py)?.1.footer.is_some())
    }

    #[getter]
    fn date_present(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(!matches!(
            self.current(py)?.1.date,
            rpptx::HeaderFooterDate::Off
        ))
    }

    #[getter]
    fn slide_number_present(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.current(py)?.1.slide_number)
    }

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
    /// retaining a compatible cache in the current `date_format`, else `datetime1`.
    #[setter]
    fn set_date(&mut self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        let date_format = self
            .date_format(py)?
            .unwrap_or_else(|| "datetime1".to_owned());
        let date = header_footer_settings(false, None, value, &date_format, None).date;
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
        let date = header_footer_settings(false, None, Some("auto".to_owned()), value, None).date;
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
        let mut candidate = presentation.inner.clone();
        if candidate
            .set_slide_transition(index, Some(&transition))
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
        {
            let proofs = presentation.footer_owners(&candidate)?;
            let prepared = presentation.prepare_owners(
                &candidate,
                "Slide.transition",
                &proofs,
                false,
                false,
                false,
            )?;
            presentation.publish_owners(candidate, prepared)?;
        }
        Ok(())
    }
}

fn seconds_to_ms(name: &str, value: Option<f64>) -> PyResult<Option<u64>> {
    value
        .map(|seconds| {
            if seconds.is_finite() && seconds >= 0.0 && seconds * 1000.0 < u64::MAX as f64 {
                Ok((seconds * 1000.0).trunc() as u64)
            } else {
                Err(PyValueError::new_err(format!(
                    "{name} must be finite, nonnegative and fit milliseconds, got {seconds}"
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
                *transition = rpptx::SlideTransition::default();
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
    fn advance_on_click(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.current(py)?.1.advance_on_click)
    }

    #[setter]
    fn set_advance_on_click(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
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
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        if candidate
            .set_all_transitions(Some(&transition))
            .map_err(|e| crate::rpptx_to_pyerr(py, e))?
        {
            let proofs = presentation.footer_owners(&candidate)?;
            let prepared = presentation.prepare_owners(
                &candidate,
                "SlideTransition.apply_to_all",
                &proofs,
                false,
                false,
                false,
            )?;
            presentation.publish_owners(candidate, prepared)?;
        }
        Ok(())
    }
}

/// The notes owner of one checked physical slide.
#[pyclass(name = "NotesSlide")]
pub struct PyNotesSlide {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyNotesSlide {
    #[getter]
    fn notes_text_frame(&self, py: Python<'_>) -> PyResult<Option<Py<PyNotesTextFrame>>> {
        let index =
            PySlide::new(self.presentation.clone_ref(py), self.path.clone()).validate(py)?;
        let presentation = self.presentation.borrow(py);
        let text = presentation
            .inner
            .slide(index)
            .ok_or_else(|| PyIndexError::new_err("slide owner is absent"))?
            .notes_text_frame_text()
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        text.map(|_| {
            Py::new(
                py,
                PyNotesTextFrame {
                    presentation: self.presentation.clone_ref(py),
                    path: self.path.clone(),
                },
            )
        })
        .transpose()
    }
}

/// A scalar frame admitted only for one exact direct notes-body owner.
#[pyclass(name = "NotesTextFrame")]
pub struct PyNotesTextFrame {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyNotesTextFrame {
    fn validated_text(&self, py: Python<'_>) -> PyResult<(usize, String)> {
        let index =
            PySlide::new(self.presentation.clone_ref(py), self.path.clone()).validate(py)?;
        let presentation = self.presentation.borrow(py);
        let text = presentation.inner.slide(index).ok_or_else(|| PyIndexError::new_err("slide owner is absent"))?
            .notes_text_frame_text().map_err(|error| crate::rpptx_to_pyerr(py, error))?
            .ok_or_else(|| PyValueError::new_err("notes body is absent; inspect slide.notes_slide.notes_text_frame again or use Slide.notes_text"))?;
        Ok((index, text))
    }
}

#[pymethods]
impl PyNotesTextFrame {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.validated_text(py).map(|(_, text)| text)
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        let (index, _) = self.validated_text(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_notes_text(index, text)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?;
        if candidate
            .saves_as(&presentation.inner)
            .map_err(|error| crate::rpptx_to_pyerr(py, error))?
        {
            return Ok(());
        }
        let survivors = presentation.text_changed_owners(&self.path)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "NotesTextFrame.text",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(())
    }
}
