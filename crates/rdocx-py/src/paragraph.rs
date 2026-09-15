use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyRuntimeError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};
use smallvec::smallvec;

use crate::document::PyDocument;
use crate::run::{PyRun, PyRunCollection};
use crate::{normalize_index, stale_to_pyerr};

#[derive(Clone, Copy)]
pub(crate) enum ParagraphLocation {
    Body(usize),
    Cell {
        table: usize,
        row: usize,
        cell: usize,
        paragraph: usize,
    },
}

pub(crate) fn paragraph_location(path: &ContentPath) -> PyResult<ParagraphLocation> {
    let paragraph = path
        .segs
        .iter()
        .rev()
        .find_map(|segment| match segment {
            PathSeg::Para(index) => Some(*index),
            _ => None,
        })
        .ok_or_else(|| PyRuntimeError::new_err("paragraph path has no paragraph index"))?;
    let table = path.segs.iter().find_map(|segment| match segment {
        PathSeg::Body(index) => Some(*index),
        _ => None,
    });
    let row = path.segs.iter().find_map(|segment| match segment {
        PathSeg::Row(index) => Some(*index),
        _ => None,
    });
    let cell = path.segs.iter().find_map(|segment| match segment {
        PathSeg::Cell(index) => Some(*index),
        _ => None,
    });
    match (table, row, cell) {
        (Some(table), Some(row), Some(cell)) => Ok(ParagraphLocation::Cell {
            table,
            row,
            cell,
            paragraph,
        }),
        (_, None, None) => Ok(ParagraphLocation::Body(paragraph)),
        _ => Err(PyRuntimeError::new_err("paragraph path is incomplete")),
    }
}

#[pyclass(name = "Paragraph")]
pub struct PyParagraph {
    document: Py<PyDocument>,
    path: ContentPath,
}

impl PyParagraph {
    pub(crate) fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
    }

    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<ParagraphLocation> {
        let location = paragraph_location(&self.path)?;
        let recovery_hint = match location {
            ParagraphLocation::Body(_) => "Re-fetch it with doc.paragraphs[i].".to_owned(),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => format!(
                "Re-fetch it with doc.tables[{table}].rows[{row}].cells[{cell}].paragraphs[{paragraph}]."
            ),
        };
        let document = self.document.borrow(py);
        self.path
            .validate_revision(document.revisions.current(), "paragraph", &recovery_hint)
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok(location)
    }

    /// Append one run, or one external hyperlink run when `url` is given, and
    /// return its handle.
    fn append_run(&self, py: Python<'_>, text: &str, url: Option<&str>) -> PyResult<Py<PyRun>> {
        let location = self.validate(py)?;
        let path = {
            let mut document = self.document.borrow_mut(py);
            let relationship_id = match url {
                Some(url) => {
                    // Check the paragraph first, so a failure leaves no
                    // orphaned relationship behind.
                    let exists = match location {
                        ParagraphLocation::Body(index) => document.inner.paragraph(index).is_some(),
                        ParagraphLocation::Cell {
                            table,
                            row,
                            cell,
                            paragraph,
                        } => document.inner.table(table).is_some_and(|table| {
                            table
                                .cell(row, cell)
                                .is_some_and(|cell| cell.paragraph(paragraph).is_some())
                        }),
                    };
                    if !exists {
                        return Err(PyIndexError::new_err("paragraph index out of range"));
                    }
                    Some(document.inner.add_hyperlink_relationship(url))
                }
                None => None,
            };
            let append = |paragraph: &mut rdocx::Paragraph<'_>| {
                let run_index = paragraph.run_count();
                match &relationship_id {
                    Some(relationship_id) => {
                        paragraph.add_hyperlink(text, relationship_id);
                    }
                    None => {
                        paragraph.add_run(text);
                    }
                }
                run_index
            };
            let run_index = match location {
                ParagraphLocation::Body(index) => append(
                    &mut document
                        .inner
                        .paragraph_mut(index)
                        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?,
                ),
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
                    append(
                        &mut cell
                            .paragraph_mut(paragraph)
                            .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?,
                    )
                }
            };
            document.revisions.bump();
            let mut segments = self.path.segs.clone();
            segments.push(PathSeg::Run(run_index));
            document.revisions.capture(segments)
        };
        Py::new(py, PyRun::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyParagraph {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let location = self.validate(py)?;
        let document = self.document.borrow(py);
        match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph(index)
                .map(|paragraph| paragraph.text()),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => document.inner.table(table).and_then(|table| {
                let cell = table.cell(row, cell)?;
                cell.paragraph(paragraph).map(|paragraph| paragraph.text())
            }),
        }
        .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))
    }

    #[getter]
    fn runs(&self, py: Python<'_>) -> PyResult<Py<PyRunCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyRunCollection::new(self.document.clone_ref(py), self.path.clone()),
        )
    }

    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let location = self.validate(py)?;
        let document = self.document.borrow(py);
        let alignment = match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph(index)
                .and_then(|paragraph| paragraph.alignment()),
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => document.inner.table(table).and_then(|table| {
                let cell = table.cell(row, cell)?;
                cell.paragraph(paragraph)
                    .and_then(|paragraph| paragraph.alignment())
            }),
        };
        alignment
            .map(|value| {
                crate::enum_object(
                    py,
                    "WD_ALIGN_PARAGRAPH",
                    crate::formatting::alignment_to_int(value),
                )
            })
            .transpose()
    }

    #[setter]
    fn set_alignment(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        let location = self.validate(py)?;
        let value = value
            .map(crate::formatting::alignment_from_int)
            .transpose()?;
        let mut document = self.document.borrow_mut(py);
        match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
                .set_alignment_value(value),
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
                    .set_alignment_value(value);
            }
        }
        Ok(())
    }

    fn add_run(&self, py: Python<'_>, text: &str) -> PyResult<Py<PyRun>> {
        self.append_run(py, text, None)
    }

    fn add_hyperlink(&self, py: Python<'_>, text: &str, url: &str) -> PyResult<Py<PyRun>> {
        self.append_run(py, text, Some(url))
    }

    #[getter]
    fn paragraph_format(
        &self,
        py: Python<'_>,
    ) -> PyResult<Py<crate::formatting::PyParagraphFormat>> {
        self.validate(py)?;
        Py::new(
            py,
            crate::formatting::PyParagraphFormat::new(
                self.document.clone_ref(py),
                self.path.clone(),
            ),
        )
    }
}

#[pyclass(name = "ParagraphCollection")]
pub struct PyParagraphCollection {
    document: Py<PyDocument>,
}

impl PyParagraphCollection {
    pub(crate) fn new(document: Py<PyDocument>) -> Self {
        Self { document }
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyParagraph>> {
        let path = {
            let document = self.document.borrow(py);
            document
                .revisions
                .capture(smallvec![PathSeg::Body(0), PathSeg::Para(index)])
        };
        Py::new(py, PyParagraph::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyParagraphCollection {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.document.borrow(py).inner.paragraph_count()
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.__len__(py);
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "paragraph")?)?
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
            "paragraph indices must be integers or slices",
        ))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyParagraphIterator>> {
        Py::new(
            py,
            PyParagraphIterator {
                document: self.document.clone_ref(py),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyParagraphIterator {
    document: Py<PyDocument>,
    index: usize,
}

#[pymethods]
impl PyParagraphIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyParagraph>>> {
        let len = self.document.borrow(py).inner.paragraph_count();
        if self.index >= len {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        let path = self
            .document
            .borrow(py)
            .revisions
            .capture(smallvec![PathSeg::Body(0), PathSeg::Para(index)]);
        Py::new(py, PyParagraph::new(self.document.clone_ref(py), path)).map(Some)
    }
}
