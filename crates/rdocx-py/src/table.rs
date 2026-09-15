use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};
use smallvec::smallvec;

use crate::document::PyDocument;
use crate::paragraph::PyParagraph;
use crate::{enum_object, length_object, normalize_index, rdocx_to_pyerr, stale_to_pyerr};

fn path_index(
    path: &ContentPath,
    expected: fn(PathSeg) -> Option<usize>,
    kind: &str,
) -> PyResult<usize> {
    path.segs
        .iter()
        .copied()
        .find_map(expected)
        .ok_or_else(|| PyIndexError::new_err(format!("{kind} index is missing")))
}

fn table_index(path: &ContentPath) -> PyResult<usize> {
    path_index(
        path,
        |segment| match segment {
            PathSeg::Body(index) => Some(index),
            _ => None,
        },
        "table",
    )
}

fn row_index(path: &ContentPath) -> PyResult<usize> {
    path_index(
        path,
        |segment| match segment {
            PathSeg::Row(index) => Some(index),
            _ => None,
        },
        "row",
    )
}

fn cell_index(path: &ContentPath) -> PyResult<usize> {
    path_index(
        path,
        |segment| match segment {
            PathSeg::Cell(index) => Some(index),
            _ => None,
        },
        "cell",
    )
}

fn alignment_from_int(value: i32) -> PyResult<rdocx::Alignment> {
    match value {
        0 => Ok(rdocx::Alignment::Left),
        1 => Ok(rdocx::Alignment::Center),
        2 => Ok(rdocx::Alignment::Right),
        _ => Err(PyValueError::new_err("unsupported table alignment")),
    }
}

fn alignment_to_int(value: rdocx::Alignment) -> Option<i32> {
    match value {
        rdocx::Alignment::Left => Some(0),
        rdocx::Alignment::Center => Some(1),
        rdocx::Alignment::Right => Some(2),
        rdocx::Alignment::Justify => None,
    }
}

fn vertical_from_int(value: i32) -> PyResult<rdocx::VerticalAlignment> {
    match value {
        0 => Ok(rdocx::VerticalAlignment::Top),
        1 => Ok(rdocx::VerticalAlignment::Center),
        3 => Ok(rdocx::VerticalAlignment::Bottom),
        _ => Err(PyValueError::new_err("unsupported cell vertical alignment")),
    }
}

fn vertical_to_int(value: rdocx::VerticalAlignment) -> i32 {
    match value {
        rdocx::VerticalAlignment::Top => 0,
        rdocx::VerticalAlignment::Center => 1,
        rdocx::VerticalAlignment::Bottom => 3,
    }
}

#[pyclass(name = "TableCollection")]
pub struct PyTableCollection {
    document: Py<PyDocument>,
}

impl PyTableCollection {
    pub(crate) fn new(document: Py<PyDocument>) -> Self {
        Self { document }
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyTable>> {
        let path = self
            .document
            .borrow(py)
            .revisions
            .capture(smallvec![PathSeg::Body(index)]);
        Py::new(py, PyTable::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyTableCollection {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.document.borrow(py).inner.table_count()
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.__len__(py);
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "table")?)?
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
            "table indices must be integers or slices",
        ))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyTableIterator>> {
        Py::new(
            py,
            PyTableIterator {
                document: self.document.clone_ref(py),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyTableIterator {
    document: Py<PyDocument>,
    index: usize,
}

#[pymethods]
impl PyTableIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyTable>>> {
        let collection = PyTableCollection::new(self.document.clone_ref(py));
        if self.index >= collection.__len__(py) {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

#[pyclass(name = "Table")]
pub struct PyTable {
    document: Py<PyDocument>,
    path: ContentPath,
}

impl PyTable {
    pub(crate) fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "table",
                "Re-fetch it with doc.tables[i].",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        table_index(&self.path)
    }
}

#[pymethods]
impl PyTable {
    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Py<PyRowCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyRowCollection::new(self.document.clone_ref(py), self.path.clone()),
        )
    }

    fn cell(&self, py: Python<'_>, row: isize, col: isize) -> PyResult<Py<PyCell>> {
        let table_index = self.validate(py)?;
        let document = self.document.borrow(py);
        let table = document
            .inner
            .table(table_index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        let row = normalize_index(row, table.row_count(), "row")?;
        let col = normalize_index(
            col,
            table.row(row).map(|row| row.cell_count()).unwrap_or(0),
            "cell",
        )?;
        let path = document.revisions.capture(smallvec![
            PathSeg::Body(table_index),
            PathSeg::Row(row),
            PathSeg::Cell(col)
        ]);
        Py::new(py, PyCell::new(self.document.clone_ref(py), path))
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let index = self.validate(py)?;
        Ok(self
            .document
            .borrow(py)
            .inner
            .table(index)
            .and_then(|table| table.style_id().map(str::to_owned)))
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let index = self.validate(py)?;
        self.document
            .borrow_mut(py)
            .inner
            .table_mut(index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
            .set_style(value);
        Ok(())
    }

    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let index = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(index)
            .and_then(|table| table.alignment())
            .and_then(alignment_to_int)
            .map(|value| enum_object(py, "WD_TABLE_ALIGNMENT", value))
            .transpose()
    }

    #[setter]
    fn set_alignment(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let index = self.validate(py)?;
        let value = alignment_from_int(value)?;
        self.document
            .borrow_mut(py)
            .inner
            .table_mut(index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
            .set_alignment(value);
        Ok(())
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let index = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(index)
            .and_then(|table| table.width())
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        let index = self.validate(py)?;
        self.document
            .borrow_mut(py)
            .inner
            .table_mut(index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
            .set_width(rdocx::Length::emu(value));
        Ok(())
    }

    #[pyo3(signature = (index, at = None))]
    fn clone_row(&self, py: Python<'_>, index: isize, at: Option<usize>) -> PyResult<Py<PyRow>> {
        let table_index = self.validate(py)?;
        let path = {
            let mut document = self.document.borrow_mut(py);
            let row_count = document
                .inner
                .table(table_index)
                .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
                .row_count();
            let index = normalize_index(index, row_count, "row")?;
            let at = at.unwrap_or(index + 1);
            if at > row_count {
                return Err(PyIndexError::new_err("row insertion index out of range"));
            }
            let inner = &mut document.inner;
            py.detach(|| inner.clone_table_row(table_index, index, at))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            document.revisions.bump();
            let mut segments = self.path.segs.clone();
            segments.push(PathSeg::Row(at));
            document.revisions.capture(segments)
        };
        Py::new(py, PyRow::new(self.document.clone_ref(py), path))
    }

    fn remove_row(&self, py: Python<'_>, index: isize) -> PyResult<()> {
        let table_index = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let row_count = document
            .inner
            .table(table_index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
            .row_count();
        let index = normalize_index(index, row_count, "row")?;
        let inner = &mut document.inner;
        py.detach(|| inner.remove_table_row(table_index, index))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.revisions.bump();
        Ok(())
    }
}

#[pyclass(name = "RowCollection")]
pub struct PyRowCollection {
    document: Py<PyDocument>,
    table_path: ContentPath,
}

impl PyRowCollection {
    fn new(document: Py<PyDocument>, table_path: ContentPath) -> Self {
        Self {
            document,
            table_path,
        }
    }
    fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let document = self.document.borrow(py);
        self.table_path
            .validate_revision(
                document.revisions.current(),
                "row collection",
                "Re-fetch it with table.rows.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        table_index(&self.table_path)
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let index = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(index)
            .map(|table| table.row_count())
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))
    }
    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyRow>> {
        self.validate(py)?;
        let mut segments = self.table_path.segs.clone();
        segments.push(PathSeg::Row(index));
        let path = self.document.borrow(py).revisions.capture(segments);
        Py::new(py, PyRow::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyRowCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }
    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.len(py)?;
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "row")?)?
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
            "row indices must be integers or slices",
        ))
    }
    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyRowIterator>> {
        self.validate(py)?;
        Py::new(
            py,
            PyRowIterator {
                document: self.document.clone_ref(py),
                table_path: self.table_path.clone(),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyRowIterator {
    document: Py<PyDocument>,
    table_path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyRowIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyRow>>> {
        let collection = PyRowCollection::new(self.document.clone_ref(py), self.table_path.clone());
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

#[pyclass(name = "Row")]
pub struct PyRow {
    document: Py<PyDocument>,
    path: ContentPath,
}

impl PyRow {
    fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
    }
    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "row",
                "Re-fetch it with table.rows[i].",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok((table_index(&self.path)?, row_index(&self.path)?))
    }
}

#[pymethods]
impl PyRow {
    #[getter]
    fn cells(&self, py: Python<'_>) -> PyResult<Py<PyCellCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyCellCollection::new(self.document.clone_ref(py), self.path.clone()),
        )
    }
}

#[pyclass(name = "CellCollection")]
pub struct PyCellCollection {
    document: Py<PyDocument>,
    row_path: ContentPath,
}

impl PyCellCollection {
    fn new(document: Py<PyDocument>, row_path: ContentPath) -> Self {
        Self { document, row_path }
    }
    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize)> {
        let document = self.document.borrow(py);
        self.row_path
            .validate_revision(
                document.revisions.current(),
                "cell collection",
                "Re-fetch it with row.cells.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok((table_index(&self.row_path)?, row_index(&self.row_path)?))
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let (table, row) = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(table)
            .and_then(|table| table.row(row).map(|row| row.cell_count()))
            .ok_or_else(|| PyIndexError::new_err("row index out of range"))
    }
    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyCell>> {
        self.validate(py)?;
        let mut segments = self.row_path.segs.clone();
        segments.push(PathSeg::Cell(index));
        let path = self.document.borrow(py).revisions.capture(segments);
        Py::new(py, PyCell::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyCellCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }
    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.len(py)?;
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "cell")?)?
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
            "cell indices must be integers or slices",
        ))
    }
    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyCellIterator>> {
        self.validate(py)?;
        Py::new(
            py,
            PyCellIterator {
                document: self.document.clone_ref(py),
                row_path: self.row_path.clone(),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyCellIterator {
    document: Py<PyDocument>,
    row_path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyCellIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyCell>>> {
        let collection = PyCellCollection::new(self.document.clone_ref(py), self.row_path.clone());
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

#[pyclass(name = "Cell")]
pub struct PyCell {
    document: Py<PyDocument>,
    path: ContentPath,
}

impl PyCell {
    fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
    }
    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "cell",
                "Re-fetch it with row.cells[i].",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok((
            table_index(&self.path)?,
            row_index(&self.path)?,
            cell_index(&self.path)?,
        ))
    }
}

#[pymethods]
impl PyCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let (table, row, cell) = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(table)
            .and_then(|table| table.cell(row, cell).map(|cell| cell.text()))
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }
    #[setter]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        {
            let mut table = document
                .inner
                .table_mut(table)
                .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
            table
                .cell(row, cell)
                .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
                .set_text(value);
        }
        document.revisions.bump();
        Ok(())
    }
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> PyResult<Py<PyCellParagraphCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyCellParagraphCollection::new(self.document.clone_ref(py), self.path.clone()),
        )
    }
    fn add_paragraph(&self, py: Python<'_>, text: &str) -> PyResult<Py<PyParagraph>> {
        let (table, row, cell) = self.validate(py)?;
        let path = {
            let mut document = self.document.borrow_mut(py);
            let paragraph = {
                let mut table = document
                    .inner
                    .table_mut(table)
                    .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
                let mut cell = table
                    .cell(row, cell)
                    .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
                let paragraph = cell.paragraph_count();
                cell.add_paragraph(text);
                paragraph
            };
            document.revisions.bump();
            document.revisions.capture(smallvec![
                PathSeg::Body(table),
                PathSeg::Row(row),
                PathSeg::Cell(cell),
                PathSeg::Para(paragraph)
            ])
        };
        Py::new(py, PyParagraph::new(self.document.clone_ref(py), path))
    }
    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let (table, row, cell) = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(table)
            .and_then(|table| table.cell(row, cell).and_then(|cell| cell.width()))
            .map(|value| length_object(py, value))
            .transpose()
    }
    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = document
            .inner
            .table_mut(table)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        table
            .cell(row, cell)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
            .set_width(rdocx::Length::emu(value));
        Ok(())
    }
    #[getter]
    fn vertical_alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let (table, row, cell) = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(table)
            .and_then(|table| {
                table
                    .cell(row, cell)
                    .and_then(|cell| cell.vertical_alignment())
            })
            .map(|value| enum_object(py, "WD_CELL_VERTICAL_ALIGNMENT", vertical_to_int(value)))
            .transpose()
    }
    #[setter]
    fn set_vertical_alignment(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let value = vertical_from_int(value)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = document
            .inner
            .table_mut(table)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        table
            .cell(row, cell)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
            .set_vertical_alignment(value);
        Ok(())
    }
}

#[pyclass(name = "CellParagraphCollection")]
pub struct PyCellParagraphCollection {
    document: Py<PyDocument>,
    cell_path: ContentPath,
}

impl PyCellParagraphCollection {
    fn new(document: Py<PyDocument>, cell_path: ContentPath) -> Self {
        Self {
            document,
            cell_path,
        }
    }
    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize, usize)> {
        let document = self.document.borrow(py);
        self.cell_path
            .validate_revision(
                document.revisions.current(),
                "cell paragraph collection",
                "Re-fetch it with cell.paragraphs.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok((
            table_index(&self.cell_path)?,
            row_index(&self.cell_path)?,
            cell_index(&self.cell_path)?,
        ))
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let (table, row, cell) = self.validate(py)?;
        self.document
            .borrow(py)
            .inner
            .table(table)
            .and_then(|table| table.cell(row, cell).map(|cell| cell.paragraph_count()))
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }
    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyParagraph>> {
        self.validate(py)?;
        let mut segments = self.cell_path.segs.clone();
        segments.push(PathSeg::Para(index));
        let path = self.document.borrow(py).revisions.capture(segments);
        Py::new(py, PyParagraph::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyCellParagraphCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }
    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.len(py)?;
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
    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyCellParagraphIterator>> {
        self.validate(py)?;
        Py::new(
            py,
            PyCellParagraphIterator {
                document: self.document.clone_ref(py),
                cell_path: self.cell_path.clone(),
                index: 0,
            },
        )
    }
}

#[pyclass]
struct PyCellParagraphIterator {
    document: Py<PyDocument>,
    cell_path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyCellParagraphIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyParagraph>>> {
        let collection =
            PyCellParagraphCollection::new(self.document.clone_ref(py), self.cell_path.clone());
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}
