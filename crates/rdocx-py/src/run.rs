use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyRuntimeError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, paragraph_location};
use crate::{normalize_index, rdocx_to_pyerr, stale_to_pyerr};

fn split_paragraph_run(
    py: Python<'_>,
    mut paragraph: rdocx::Paragraph<'_>,
    run_index: usize,
    offset: usize,
) -> PyResult<()> {
    if run_index >= paragraph.run_count() {
        return Err(PyIndexError::new_err("run index out of range"));
    }
    paragraph
        .split_run(run_index, offset)
        .map_err(|error| rdocx_to_pyerr(py, error))
}

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
}

impl PyRun {
    pub(crate) fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
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
        }
        .ok_or_else(|| PyIndexError::new_err("run index out of range"))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        let (location, run_index) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        match location {
            ParagraphLocation::Body(index) => {
                let mut paragraph = document
                    .inner
                    .paragraph_mut(index)
                    .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
                paragraph
                    .run_mut(run_index)
                    .map(|mut run| run.set_text(text))
            }
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
                let mut paragraph = cell
                    .paragraph_mut(paragraph)
                    .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
                paragraph
                    .run_mut(run_index)
                    .map(|mut run| run.set_text(text))
            }
        }
        .ok_or_else(|| PyIndexError::new_err("run index out of range"))?;
        Ok(())
    }

    #[getter]
    fn font(&self, py: Python<'_>) -> PyResult<Py<crate::formatting::PyFont>> {
        self.validate(py)?;
        Py::new(
            py,
            crate::formatting::PyFont::new(self.document.clone_ref(py), self.path.clone()),
        )
    }

    fn split(&self, py: Python<'_>, offset: usize) -> PyResult<Py<PyRun>> {
        let (location, run_index) = self.validate(py)?;
        let path = {
            let mut document = self.document.borrow_mut(py);
            match location {
                ParagraphLocation::Body(index) => split_paragraph_run(
                    py,
                    document
                        .inner
                        .paragraph_mut(index)
                        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?,
                    run_index,
                    offset,
                )?,
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
                    split_paragraph_run(
                        py,
                        cell.paragraph_mut(paragraph)
                            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?,
                        run_index,
                        offset,
                    )?;
                }
            }
            document.revisions.bump();
            let mut segments = self.path.segs.clone();
            segments.pop();
            segments.push(PathSeg::Run(run_index + 1));
            document.revisions.capture(segments)
        };
        Py::new(py, PyRun::new(self.document.clone_ref(py), path))
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
        }
        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyRun>> {
        self.validate(py)?;
        let path = {
            let document = self.document.borrow(py);
            let mut segments = self.paragraph_path.segs.clone();
            segments.push(PathSeg::Run(index));
            document.revisions.capture(segments)
        };
        Py::new(py, PyRun::new(self.document.clone_ref(py), path))
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
