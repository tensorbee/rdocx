use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, paragraph_location};
use crate::{normalize_index, stale_to_pyerr};

fn checked_break_kind(value: &Bound<'_, PyAny>) -> PyResult<&'static str> {
    if !value.is_instance_of::<pyo3::types::PyBool>() {
        if let Ok(name) = value.extract::<String>() {
            match name.as_str() {
                "line" => return Ok("line"),
                "page" => return Ok("page"),
                "column" => return Ok("column"),
                _ => {}
            }
        } else if let Ok(number) = value.extract::<i64>() {
            match number {
                6 => return Ok("line"),
                7 => return Ok("page"),
                8 => return Ok("column"),
                _ => {}
            }
        }
    }
    Err(PyValueError::new_err(
        "use WD_BREAK.LINE/PAGE/COLUMN or line/page/column strings",
    ))
}

#[pyclass(name = "Run")]
pub struct PyRun {
    document: Py<PyDocument>,
    path: ContentPath,
    run_path: rdocx::AcceptedRunPath,
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyRun {
    pub(crate) fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        run_path: rdocx::AcceptedRunPath,
        related: Option<rdocx::ContentLocation>,
    ) -> Self {
        Self {
            document,
            path,
            run_path,
            related,
            owner: None,
            furniture_scope: None,
        }
    }

    pub(crate) fn bind_current(
        mut self,
        py: Python<'_>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> PyResult<Self> {
        self.furniture_scope = furniture_scope;
        let paragraph = self.related.as_ref().map_or_else(
            || paragraph_location(&self.path),
            |location| Ok(ParagraphLocation::Related(location.clone())),
        )?;
        {
            let document = self.document.borrow(py);
            match paragraph.native(&document.inner) {
                Ok(location) => {
                    let (path, owner) = document.capture_owner(
                        &crate::document::WordOwnerLocation::Run(location, self.run_path.clone()),
                        self.path.segs.clone(),
                        crate::document::WordView::Owner,
                        self.furniture_scope.as_ref(),
                    )?;
                    self.path = path;
                    self.owner = Some(owner);
                }
                Err(_) => {
                    self.path = document.capture_snapshot(self.path.segs.clone())?;
                }
            }
        }
        Ok(self)
    }

    pub(crate) fn bind_prepared(
        mut self,
        py: Python<'_>,
        document: &PyDocument,
        prepared: &crate::document::PreparedWordRevision,
        candidate: &rdocx::Document,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> PyResult<Self> {
        self.furniture_scope =
            document.prepared_furniture_scope(prepared, furniture_scope.as_ref())?;
        let paragraph = self.related.as_ref().map_or_else(
            || paragraph_location(&self.path),
            |location| Ok(ParagraphLocation::Related(location.clone())),
        )?;
        let location = document
            .paragraph_route(&paragraph, Some((prepared, candidate)))
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let (path, owner) = document.capture_prepared_owner(
            prepared,
            &crate::document::WordOwnerLocation::Run(location, self.run_path.clone()),
            self.path.segs.clone(),
            crate::document::WordView::Owner,
            self.furniture_scope.as_ref(),
        )?;
        self.path = path;
        self.owner = Some(owner);
        Ok(self)
    }

    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<(ParagraphLocation, usize)> {
        let document = self.document.borrow(py);
        let recovery_hint = if self.related.is_some() {
            "Re-fetch it from the current rich story paragraph's runs.".to_owned()
        } else {
            match paragraph_location(&self.path)? {
                ParagraphLocation::Body(paragraph) => {
                    format!("Re-fetch it with doc.paragraphs[{paragraph}].runs[i].")
                }
                ParagraphLocation::Cell {
                    table,
                    row,
                    cell,
                    paragraph,
                } => format!(
                    "Re-fetch it with doc.tables[{table}].rows[{row}].cells[{cell}].paragraphs[{paragraph}].runs[i]."
                ),
                ParagraphLocation::Related(_) => {
                    "Re-fetch it from the current rich story paragraph's runs.".to_owned()
                }
            }
        };
        self.path
            .validate_scopes(&document.revisions, "run", &recovery_hint)
            .map_err(|error| stale_to_pyerr(py, error))?;
        let location = match &self.related {
            Some(location) => match &self.owner {
                Some(owner) => ParagraphLocation::Related(document.current_run_owner(owner)?.0),
                None => ParagraphLocation::Related(location.clone()),
            },
            None => paragraph_location(&self.path)?,
        };
        let index = location
            .read(&document.inner, |paragraph| {
                (0..paragraph.run_count())
                    .find(|index| paragraph.run_path(*index).as_ref() == Some(&self.run_path))
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("run physical owner is no longer present"))?;
        Ok((location, index))
    }

    pub(crate) fn belongs_to(&self, py: Python<'_>, document: &Py<PyDocument>) -> bool {
        self.document.bind(py).is(document.bind(py))
    }

    /// Apply one edit to this run where it lives, in the body or in a cell.
    fn edit(&self, py: Python<'_>, edit: impl FnOnce(&mut rdocx::run::Run<'_>)) -> PyResult<()> {
        let (location, _) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        location
            .edit(&mut candidate, |paragraph| {
                paragraph.edit_run(&self.run_path, edit)
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let text = location
            .read(&candidate, |paragraph| paragraph.text())
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        document.publish_run_content_candidate(candidate, self.owner.as_ref(), text)
    }
}

#[pymethods]
impl PyRun {
    #[getter]
    fn element<'py>(slf: &Bound<'py, Self>) -> Bound<'py, Self> {
        slf.clone()
    }

    #[getter]
    fn xml(&self, py: Python<'_>) -> PyResult<String> {
        self.validate(py)?;
        let document = self.document.borrow(py);
        let owner = self
            .owner
            .as_ref()
            .ok_or_else(|| PyValueError::new_err("run XML owner route is unsupported"))?;
        let current = || -> PyResult<Vec<u8>> {
            crate::document::note_word_whole_source_read();
            let (location, path) = document.current_run_owner(owner)?;
            document
                .inner
                .run_element_xml(&location, &path)
                .map_err(|error| crate::rdocx_to_pyerr(py, error))
        };
        // A run of a direct body paragraph serializes with that paragraph alone.
        let bytes = match document
            .direct_body_run_slot(owner)
            .and_then(|(slot, path)| {
                document
                    .inner
                    .direct_body_run_xml(slot, &path)
                    .ok()
                    .flatten()
            }) {
            Some(bytes) => {
                #[cfg(debug_assertions)]
                crate::document::assert_direct_read(&bytes, current, "run");
                bytes
            }
            None => current()?,
        };
        String::from_utf8(bytes).map_err(|error| PyValueError::new_err(error.to_string()))
    }

    #[getter]
    fn _r(&self, py: Python<'_>) -> PyResult<()> {
        self.validate(py)?;
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "read run.xml and write run.replace_xml(xml) instead of lxml",
        ))
    }

    #[getter]
    fn _element(&self, py: Python<'_>) -> PyResult<()> {
        self.validate(py)?;
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "read run.xml and write run.replace_xml(xml) instead of lxml",
        ))
    }

    fn replace_xml(&self, py: Python<'_>, xml: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let owner = self
            .owner
            .as_ref()
            .ok_or_else(|| PyValueError::new_err("run XML owner route is unsupported"))?;
        let (location, path) = document.current_run_owner(owner)?;
        let current = document
            .inner
            .story_range_paragraph_snapshot(&location)
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let paragraph = rdocx_oxml::CT_P::from_xml_fragment(current.xml())
            .map_err(|error| crate::rdocx_to_pyerr(py, error.into()))?;
        let run_index = paragraph
            .accepted_run_paths()
            .iter()
            .position(|candidate| candidate == &path)
            .ok_or_else(|| PyValueError::new_err("run XML owner is absent"))?;
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .replace_run_xml(
                &rdocx::StoryRunPosition {
                    location: location.clone(),
                    run_index,
                },
                xml,
            )
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        document.publish_owner_candidate(
            py,
            candidate,
            "Run.replace_xml",
            crate::document::WordPublication::ElementXml(crate::document::WordOwnerLocation::Run(
                location, path,
            )),
        )
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let (location, run_index) = self.validate(py)?;
        let document = self.document.borrow(py);
        location
            .read(&document.inner, |paragraph| {
                paragraph.run(run_index).map(|run| run.text())
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("run index out of range"))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        let (location, _) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let publication = match &self.owner {
            Some(owner) => {
                let (location, path) = document.current_run_owner(owner)?;
                crate::document::WordPublication::RunText(crate::document::WordOwnerLocation::Run(
                    location, path,
                ))
            }
            None => crate::document::WordPublication::Conservative,
        };
        let mut candidate = document.inner.clone_for_staging();
        location
            .edit(&mut candidate, |paragraph| {
                paragraph.edit_run(&self.run_path, |run| run.set_text(text))
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        document.publish_owner_candidate(py, candidate, "Run.text", publication)
    }

    #[pyo3(signature = (data, filename=None, width=None, height=None, *, options=None))]
    fn add_picture(
        &self,
        py: Python<'_>,
        data: &Bound<'_, PyAny>,
        filename: Option<&str>,
        width: Option<&Bound<'_, PyAny>>,
        height: Option<&Bound<'_, PyAny>>,
        options: Option<PyRef<'_, crate::document::PyPictureOptions>>,
    ) -> PyResult<Py<crate::document::PyPicture>> {
        let (location, index) = self.validate(py)?;
        let (data, filename) = crate::document::picture_input(data, filename)?;
        let width = crate::document::checked_section_length(width, "picture width")?
            .map(rdocx::Length::emu);
        let height = crate::document::checked_section_length(height, "picture height")?
            .map(rdocx::Length::emu);
        let mut document = self.document.borrow_mut(py);
        let position = rdocx::StoryRunPosition {
            location: location
                .native(&document.inner)
                .map_err(|error| crate::rdocx_to_pyerr(py, error))?,
            run_index: index + 1,
        };
        let mut candidate = document.inner.clone_for_staging();
        let picture = candidate
            .insert_picture_with_options_at(
                &position,
                &data,
                &filename,
                width,
                height,
                options.map(|o| o.inner.clone()),
            )
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Run.add_picture",
            crate::document::WordPublication::Conservative,
        )?;
        let result = Py::new(
            py,
            crate::document::PyPicture {
                inner: picture,
                binding: document.capture_prepared_snapshot(&prepared, smallvec::smallvec![])?,
            },
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    // A tab is appended inside this run, so no run index moves and live
    // handles stay valid.
    fn add_tab(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |run| run.add_tab())
    }

    #[pyo3(signature = (kind = "line"))]
    fn add_break(
        &self,
        py: Python<'_>,
        #[pyo3(from_py_with = checked_break_kind)] kind: &str,
    ) -> PyResult<()> {
        let kind = match kind {
            "line" => rdocx::BreakKind::Line,
            "page" => rdocx::BreakKind::Page,
            "column" => rdocx::BreakKind::Column,
            _ => {
                return Err(pyo3::exceptions::PyValueError::new_err(
                    "break kind must be line, page or column",
                ));
            }
        };
        if kind == rdocx::BreakKind::Page {
            let (location, _) = self.validate(py)?;
            if let ParagraphLocation::Related(location) = location {
                let document = self.document.borrow(py);
                let in_header = document
                    .inner
                    .stories()
                    .map_err(|error| crate::rdocx_to_pyerr(py, error))?
                    .iter()
                    .any(|story| {
                        story.kind() == rdocx::StoryKind::Header
                            && story.part_name() == location.story().part_name()
                    });
                if in_header {
                    return Err(PyValueError::new_err(
                        "a page break cannot be added inside a header",
                    ));
                }
            }
        }
        self.edit(py, |run| run.add_break(kind))
    }

    #[pyo3(signature = (instruction, cached_result = ""))]
    fn add_field(&self, py: Python<'_>, instruction: &str, cached_result: &str) -> PyResult<()> {
        let (location, _) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        location
            .edit(&mut candidate, |paragraph| {
                let mut added = Ok(());
                paragraph.edit_run(&self.run_path, |run| {
                    added = run.add_field(instruction, cached_result)
                })?;
                added
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Run.add_field",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }

    /// Remove this run from its paragraph, keeping the markers around it.
    fn remove(&self, py: Python<'_>) -> PyResult<()> {
        let (location, run_index) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        location
            .edit(&mut candidate, |paragraph| paragraph.remove_run(run_index))
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Run.remove",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }

    #[getter]
    fn bold(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.font(py)?.borrow(py).bold(py)
    }
    #[setter]
    fn set_bold(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.font(py)?.borrow(py).set_bold(py, value)
    }
    #[getter]
    fn italic(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.font(py)?.borrow(py).italic(py)
    }
    #[setter]
    fn set_italic(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.font(py)?.borrow(py).set_italic(py, value)
    }
    #[getter]
    fn underline(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.font(py)?.borrow(py).underline(py)
    }
    #[setter]
    fn set_underline(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.font(py)?.borrow(py).set_underline(py, value)
    }

    #[getter]
    fn font(&self, py: Python<'_>) -> PyResult<Py<crate::formatting::PyFont>> {
        self.validate(py)?;
        let path = match &self.owner {
            Some(owner) => {
                let document = self.document.borrow(py);
                let (location, run_path) = document.current_run_owner(owner)?;
                document
                    .capture_owner(
                        &crate::document::WordOwnerLocation::Run(location, run_path),
                        self.path.segs.clone(),
                        crate::document::WordView::Content,
                        self.furniture_scope.as_ref(),
                    )?
                    .0
            }
            None => self.path.clone(),
        };
        Py::new(
            py,
            crate::formatting::PyFont::new(
                self.document.clone_ref(py),
                path,
                self.run_path.clone(),
                self.related.clone(),
                self.owner.clone(),
            ),
        )
    }

    #[getter]
    fn style_id(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let (location, run_index) = self.validate(py)?;
        Ok(crate::formatting::run_snapshot(py, &self.document, location, run_index)?.style_id)
    }

    #[setter]
    fn set_style_id(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let (location, _) = self.validate(py)?;
        crate::formatting::apply_run_update(
            py,
            &self.document,
            location,
            &self.run_path,
            crate::formatting::FontUpdate::Style(value),
        )
    }
}

#[pyclass(name = "RunCollection")]
pub struct PyRunCollection {
    document: Py<PyDocument>,
    paragraph_path: ContentPath,
    related: Option<rdocx::ContentLocation>,
    paragraph_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyRunCollection {
    pub(crate) fn new(
        document: Py<PyDocument>,
        paragraph_path: ContentPath,
        related: Option<rdocx::ContentLocation>,
        paragraph_owner: Option<oxml_py_support::RevisionState>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> Self {
        Self {
            document,
            paragraph_path,
            related,
            paragraph_owner,
            furniture_scope,
        }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<ParagraphLocation> {
        let document = self.document.borrow(py);
        self.paragraph_path
            .validate_scopes(
                &document.revisions,
                "run collection",
                "Re-fetch it with paragraph.runs.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        match &self.related {
            Some(location) => match &self.paragraph_owner {
                Some(owner) => document
                    .current_paragraph_owner(owner)
                    .map(ParagraphLocation::Related),
                None => Ok(ParagraphLocation::Related(location.clone())),
            },
            None => paragraph_location(&self.paragraph_path),
        }
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let location = self.validate(py)?;
        let document = self.document.borrow(py);
        location
            .read(&document.inner, |paragraph| paragraph.run_count())
            .map_err(|error| crate::rdocx_to_pyerr(py, error))
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyRun>> {
        let location = self.validate(py)?;
        let (path, run_path) = {
            let document = self.document.borrow(py);
            let mut segments = self.paragraph_path.segs.clone();
            segments.push(PathSeg::Run(index));
            let run_path = location
                .read(&document.inner, |paragraph| paragraph.run_path(index))
                .map_err(|error| crate::rdocx_to_pyerr(py, error))?
                .ok_or_else(|| PyIndexError::new_err("run index out of range"))?;
            (document.revisions.capture(segments), run_path)
        };
        Py::new(
            py,
            PyRun::new(
                self.document.clone_ref(py),
                path,
                run_path,
                match &location {
                    ParagraphLocation::Related(location) => Some(location.clone()),
                    _ => None,
                },
            )
            .bind_current(py, self.furniture_scope.clone())?,
        )
    }
}

#[pymethods]
impl PyRunCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.len(py)?;
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "run")?)?
                .into_any());
        }
        if key.is_instance_of::<PySlice>() {
            let (start, stop, step): (isize, isize, isize) =
                key.call_method1("indices", (len,))?.extract()?;
            let items = PyList::empty(py);
            let mut index = start;
            while if step > 0 { index < stop } else { index > stop } {
                items.append(self.item(py, index as usize)?)?;
                index += step;
            }
            return Ok(items.into_any().unbind());
        }
        Err(PyTypeError::new_err(
            "run indices must be integers or slices",
        ))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyRunIterator>> {
        self.validate(py)?;
        Py::new(
            py,
            PyRunIterator {
                document: self.document.clone_ref(py),
                paragraph_path: self.paragraph_path.clone(),
                related: self.related.clone(),
                paragraph_owner: self.paragraph_owner.clone(),
                furniture_scope: self.furniture_scope.clone(),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyRunIterator {
    document: Py<PyDocument>,
    paragraph_path: ContentPath,
    related: Option<rdocx::ContentLocation>,
    paragraph_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
    index: usize,
}

#[pymethods]
impl PyRunIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyRun>>> {
        let collection = PyRunCollection::new(
            self.document.clone_ref(py),
            self.paragraph_path.clone(),
            self.related.clone(),
            self.paragraph_owner.clone(),
            self.furniture_scope.clone(),
        );
        let len = collection.len(py)?;
        if self.index >= len {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}
