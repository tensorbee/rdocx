use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyRuntimeError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, paragraph_location};
use crate::{normalize_index, stale_to_pyerr};

fn path_indices(path: &ContentPath) -> PyResult<(ParagraphLocation, usize)> {
    let run = path.segs.iter().find_map(|segment| match segment {
        PathSeg::Run(index) => Some(*index),
        _ => None,
    });
    run.map(|run| paragraph_location(path).map(|paragraph| (paragraph, run)))
        .transpose()?
        .ok_or_else(|| PyRuntimeError::new_err("run path is incomplete"))
}

#[pyclass(name = "Run")]
pub struct PyRun {
    document: Py<PyDocument>,
    path: ContentPath,
    run_path: rdocx::AcceptedRunPath,
}

impl PyRun {
    pub(crate) fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        run_path: rdocx::AcceptedRunPath,
    ) -> Self {
        Self {
            document,
            path,
            run_path,
        }
    }

    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<(ParagraphLocation, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "run",
                "Re-fetch it with paragraph.runs[i].",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        path_indices(&self.path)
    }

    pub(crate) fn belongs_to(&self, py: Python<'_>, document: &Py<PyDocument>) -> bool {
        self.document.bind(py).is(document.bind(py))
    }

    /// Apply one edit to this run where it lives, in the body or in a cell.
    fn edit(&self, py: Python<'_>, edit: impl FnOnce(&mut rdocx::run::Run<'_>)) -> PyResult<()> {
        let (location, _) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
                .edit_run(&self.run_path, edit),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => {
                let mut table = document
                    .inner
                    .table_mut(table)
                    .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
                let mut cell = table
                    .cell(row, cell)
                    .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
                cell.paragraph_mut(paragraph)
                    .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
                    .edit_run(&self.run_path, edit)
            }
            ParagraphLocation::Story { slot, at } => {
                crate::story::edit_paragraph(py, &mut document, slot, at, |paragraph| {
                    paragraph.edit_run(&self.run_path, edit)
                })?
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
            }
        }
        .map_err(|error| crate::rdocx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyRun {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let (location, run_index) = self.validate(py)?;
        let document = self.document.borrow(py);
        match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph(index)
                .and_then(|paragraph| paragraph.run(run_index).map(|run| run.text())),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => document.inner.table(table).and_then(|table| {
                let cell = table.cell(row, cell)?;
                let paragraph = cell.paragraph(paragraph)?;
                paragraph.run(run_index).map(|run| run.text())
            }),
            ParagraphLocation::Story { slot, at } => {
                crate::story::read_paragraph(py, &document, slot, at, |paragraph| {
                    paragraph.run(run_index).map(|run| run.text())
                })?
                .flatten()
            }
        }
        .ok_or_else(|| PyIndexError::new_err("run index out of range"))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        self.edit(py, |run| run.set_text(text))
    }

    // A tab is appended inside this run, so no run index moves and live
    // handles stay valid.
    fn add_tab(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |run| run.add_tab())
    }

    /// Append a break inside this run, as python-docx `Run.add_break` does:
    /// `WD_BREAK.LINE` (the default), `PAGE`, `COLUMN`, or a text-wrapping
    /// `LINE_CLEAR_LEFT`, `LINE_CLEAR_RIGHT` or `LINE_CLEAR_ALL` break.
    /// No run index moves, so live handles stay valid.
    #[pyo3(signature = (break_type = 6))]
    fn add_break(&self, py: Python<'_>, break_type: i32) -> PyResult<()> {
        let kind = crate::story::break_kind(break_type)?;
        let (location, _) = self.validate(py)?;
        crate::story::check_break_location(location, kind)?;
        self.edit(py, |run| run.add_break(kind))
    }

    #[pyo3(signature = (instruction, cached_result = ""))]
    fn add_field(&self, py: Python<'_>, instruction: &str, cached_result: &str) -> PyResult<()> {
        let mut added = Ok(());
        self.edit(py, |run| added = run.add_field(instruction, cached_result))?;
        added.map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        // The field becomes a story item of its own, which moves the index
        // path of every later story item.
        self.document.borrow_mut(py).revisions.bump();
        Ok(())
    }

    /// Remove this run from its paragraph, keeping the markers around it.
    fn remove(&self, py: Python<'_>) -> PyResult<()> {
        let (location, run_index) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let removed = match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
                .remove_run(run_index),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => {
                let mut table = document
                    .inner
                    .table_mut(table)
                    .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
                let mut cell = table
                    .cell(row, cell)
                    .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
                cell.paragraph_mut(paragraph)
                    .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
                    .remove_run(run_index)
            }
            ParagraphLocation::Story { slot, at } => {
                crate::story::edit_paragraph(py, &mut document, slot, at, |paragraph| {
                    paragraph.remove_run(run_index)
                })?
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
            }
        };
        removed.map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        document.revisions.bump();
        Ok(())
    }

    #[getter]
    fn font(&self, py: Python<'_>) -> PyResult<Py<crate::formatting::PyFont>> {
        self.validate(py)?;
        Py::new(
            py,
            crate::formatting::PyFont::new(
                self.document.clone_ref(py),
                self.path.clone(),
                self.run_path.clone(),
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
}

impl PyRunCollection {
    pub(crate) fn new(document: Py<PyDocument>, paragraph_path: ContentPath) -> Self {
        Self {
            document,
            paragraph_path,
        }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<ParagraphLocation> {
        let document = self.document.borrow(py);
        self.paragraph_path
            .validate_revision(
                document.revisions.current(),
                "run collection",
                "Re-fetch it with paragraph.runs.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        paragraph_location(&self.paragraph_path)
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let location = self.validate(py)?;
        let document = self.document.borrow(py);
        match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph(index)
                .map(|paragraph| paragraph.run_count()),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => document.inner.table(table).and_then(|table| {
                let cell = table.cell(row, cell)?;
                cell.paragraph(paragraph)
                    .map(|paragraph| paragraph.run_count())
            }),
            ParagraphLocation::Story { slot, at } => {
                crate::story::read_paragraph(py, &document, slot, at, |paragraph| {
                    paragraph.run_count()
                })?
            }
        }
        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyRun>> {
        let location = self.validate(py)?;
        let (path, run_path) = {
            let document = self.document.borrow(py);
            let mut segments = self.paragraph_path.segs.clone();
            segments.push(PathSeg::Run(index));
            let run_path = match location {
                ParagraphLocation::Body(paragraph) => document
                    .inner
                    .paragraph(paragraph)
                    .and_then(|paragraph| paragraph.run_path(index)),
                ParagraphLocation::Cell {
                    table,
                    row,
                    cell,
                    paragraph,
                } => document.inner.table(table).and_then(|table| {
                    let cell = table.cell(row, cell)?;
                    cell.paragraph(paragraph)
                        .and_then(|paragraph| paragraph.run_path(index))
                }),
                ParagraphLocation::Story { slot, at } => {
                    crate::story::read_paragraph(py, &document, slot, at, |paragraph| {
                        paragraph.run_path(index)
                    })?
                    .flatten()
                }
            }
            .ok_or_else(|| PyIndexError::new_err("run index out of range"))?;
            (document.revisions.capture(segments), run_path)
        };
        Py::new(py, PyRun::new(self.document.clone_ref(py), path, run_path))
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
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyRunIterator {
    document: Py<PyDocument>,
    paragraph_path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyRunIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyRun>>> {
        let collection =
            PyRunCollection::new(self.document.clone_ref(py), self.paragraph_path.clone());
        let len = collection.len(py)?;
        if self.index >= len {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}
