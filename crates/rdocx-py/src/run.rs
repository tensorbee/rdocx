use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyRuntimeError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBytes, PyList, PySlice};
use smallvec::SmallVec;

use crate::document::{EditScope, PyDocument, raw_xml_argument, raw_xml_error};
use crate::normalize_index;
use crate::paragraph::{ParagraphLocation, paragraph_location, xml_paragraph};

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
        path_indices(&self.resolved(py)?)
    }

    /// This handle's path at the current revision.
    pub(crate) fn resolved(&self, py: Python<'_>) -> PyResult<ContentPath> {
        self.document.borrow(py).resolve_path(
            py,
            &self.path,
            "run",
            "Re-fetch it with paragraph.runs[i].",
        )
    }

    /// The path of the paragraph that holds this run, at the current revision.
    fn paragraph_path(&self, py: Python<'_>) -> PyResult<SmallVec<[PathSeg; 5]>> {
        let mut segments = self.resolved(py)?.segs;
        segments.pop();
        Ok(segments)
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
        }
        .map_err(|error| crate::rdocx_to_pyerr(py, error))
    }
}

impl PyRun {
    fn font_attribute(&self, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        Ok(self.font(py)?.bind(py).getattr(name)?.unbind())
    }

    fn set_font_attribute(&self, py: Python<'_>, name: &str, value: Option<bool>) -> PyResult<()> {
        self.font(py)?.bind(py).setattr(name, value)
    }
}

#[pymethods]
impl PyRun {
    fn __getattr__(&self, name: &str) -> PyResult<Py<PyAny>> {
        Err(crate::missing_attribute("Run", name))
    }

    fn __setattr__(slf: &Bound<'_, Self>, name: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        crate::set_attribute(slf.as_any(), "Run", name, value)
    }

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
        self.edit(py, |run| run.set_text(text))
    }

    /// python-docx's `run.bold`, the same as `run.font.bold`.
    #[getter]
    fn bold(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.font_attribute(py, "bold")
    }

    #[setter]
    fn set_bold(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.set_font_attribute(py, "bold", value)
    }

    /// python-docx's `run.italic`, the same as `run.font.italic`.
    #[getter]
    fn italic(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.font_attribute(py, "italic")
    }

    #[setter]
    fn set_italic(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.set_font_attribute(py, "italic", value)
    }

    /// python-docx's `run.underline`, the same as `run.font.underline`.
    #[getter]
    fn underline(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.font_attribute(py, "underline")
    }

    #[setter]
    fn set_underline(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.font(py)?.bind(py).setattr("underline", value)
    }

    /// This run's `w:r` element as standalone XML bytes.
    #[getter]
    fn xml<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let (location, _) = self.validate(py)?;
        let target = rdocx::XmlTarget::Run(xml_paragraph(location), self.run_path.clone());
        let xml = self
            .document
            .borrow(py)
            .inner
            .element_xml(&target)
            .map_err(|error| raw_xml_error(py, error))?;
        Ok(PyBytes::new(py, &xml))
    }

    /// Replace this run with one `w:r` element given as XML. Every handle
    /// stays valid.
    fn replace_xml(&self, py: Python<'_>, xml: &Bound<'_, PyAny>) -> PyResult<()> {
        let xml = raw_xml_argument(xml)?;
        let (location, _) = self.validate(py)?;
        let target = rdocx::XmlTarget::Run(xml_paragraph(location), self.run_path.clone());
        self.document
            .borrow_mut(py)
            .inner
            .replace_element_xml(&target, &xml)
            .map_err(|error| raw_xml_error(py, error))
    }

    // A tab is appended inside this run, so no run index moves and live
    // handles stay valid.
    fn add_tab(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |run| run.add_tab())
    }

    #[pyo3(signature = (instruction, cached_result = ""))]
    fn add_field(&self, py: Python<'_>, instruction: &str, cached_result: &str) -> PyResult<()> {
        let mut added = Ok(());
        self.edit(py, |run| added = run.add_field(instruction, cached_result))?;
        added.map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        // The field becomes a story item of its own, which moves the index
        // path of every later story item and the runs of this paragraph.
        let paragraph = self.paragraph_path(py)?;
        self.document
            .borrow_mut(py)
            .bump_edit("Run.add_field", EditScope::Below(paragraph));
        Ok(())
    }

    /// Remove this run from its paragraph, keeping the markers around it.
    fn remove(&self, py: Python<'_>) -> PyResult<()> {
        let (location, run_index) = self.validate(py)?;
        let paragraph = self.paragraph_path(py)?;
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
        };
        removed.map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        document.bump_edit("Run.remove", EditScope::Below(paragraph));
        Ok(())
    }

    #[getter]
    fn font(&self, py: Python<'_>) -> PyResult<Py<crate::formatting::PyFont>> {
        let resolved = self.resolved(py)?;
        Py::new(
            py,
            crate::formatting::PyFont::new(
                self.document.clone_ref(py),
                resolved,
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
        paragraph_location(&self.resolved(py)?)
    }

    /// The paragraph path at the current revision.
    fn resolved(&self, py: Python<'_>) -> PyResult<ContentPath> {
        self.document.borrow(py).resolve_path(
            py,
            &self.paragraph_path,
            "run collection",
            "Re-fetch it with paragraph.runs.",
        )
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
        let resolved = self.resolved(py)?;
        let location = paragraph_location(&resolved)?;
        let (path, run_path) = {
            let document = self.document.borrow(py);
            let mut segments = resolved.segs;
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
