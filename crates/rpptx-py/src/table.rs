use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};

use crate::dml::{FillTarget, PyFillFormat, PyLineFormat};
use crate::normalize_index;
use crate::presentation::PyPresentation;
use crate::shape::{length, shape_mut_at, shape_ref_at};
use crate::{Scope, rpptx_to_pyerr, validate_path};

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyTable>()?;
    module.add_class::<PyColumn>()?;
    module.add_class::<PyColumnCollection>()?;
    module.add_class::<PyRow>()?;
    module.add_class::<PyRowCollection>()?;
    module.add_class::<PyCell>()?;
    Ok(())
}

fn row_index(path: &ContentPath) -> Option<usize> {
    path.segs.iter().find_map(|segment| match segment {
        PathSeg::Row(index) => Some(*index),
        _ => None,
    })
}

fn cell_index(path: &ContentPath) -> Option<usize> {
    path.segs.iter().rev().find_map(|segment| match segment {
        PathSeg::Cell(index) => Some(*index),
        _ => None,
    })
}

/// Resolves the index a new row or column gets, as `list.insert` does, where
/// `None` or the length appends, but raises for an index outside `-len..=len`.
fn insertion_index(index: Option<isize>, len: usize, kind: &str) -> PyResult<usize> {
    let Some(index) = index else {
        return Ok(len);
    };
    let resolved = if index < 0 {
        len as isize + index
    } else {
        index
    };
    if resolved < 0 || resolved > len as isize {
        return Err(PyIndexError::new_err(format!("{kind} index out of range")));
    }
    Ok(resolved as usize)
}

/// Applies one row or column edit to the table at `path` and invalidates
/// row, column, cell, and text handles, because each names an index.
fn restructure(
    py: Python<'_>,
    presentation: &Py<PyPresentation>,
    path: &ContentPath,
    cause: &'static str,
    edit: impl FnOnce(&mut rpptx::TableMut<'_>) -> rpptx::Result<()>,
) -> PyResult<()> {
    let mut presentation = presentation.borrow_mut(py);
    let mut table = shape_mut_at(&mut presentation.inner, path)
        .and_then(rpptx::ShapeMut::into_table_mut)
        .ok_or_else(|| PyValueError::new_err("shape has no table"))?;
    edit(&mut table).map_err(|error| rpptx_to_pyerr(py, error))?;
    presentation.revisions.invalidate(Scope::Tables, cause);
    Ok(())
}

/// Returns the index a row or column handle names when it belongs to the
/// table at `table_path`.
fn member_index(
    table_path: &ContentPath,
    member_path: &ContentPath,
    index: fn(&PathSeg) -> Option<usize>,
) -> Option<usize> {
    let (last, table) = member_path.segs.split_last()?;
    (table == table_path.segs.as_slice()).then(|| index(last))?
}

/// Returns the path segments that name a cell's table shape.
fn table_segments(path: &ContentPath) -> impl Iterator<Item = &PathSeg> {
    path.segs
        .iter()
        .filter(|segment| !matches!(segment, PathSeg::Row(_) | PathSeg::Cell(_)))
}

pub(crate) fn cell_ref_at<'a>(
    presentation: &'a rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::TableCellRef<'a>> {
    shape_ref_at(presentation, path)?
        .table()?
        .cell(row_index(path)?, cell_index(path)?)
}

pub(crate) fn cell_mut_at<'a>(
    presentation: &'a mut rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::TableCellMut<'a>> {
    let (row, column) = (row_index(path)?, cell_index(path)?);
    shape_mut_at(presentation, path)?
        .into_table_mut()?
        .into_cell_mut(row, column)
}

/// Left, right, top, and bottom cell margins, as the facade reports them.
type Margins = (
    Option<rpptx::Emu>,
    Option<rpptx::Emu>,
    Option<rpptx::Emu>,
    Option<rpptx::Emu>,
);

#[pyclass(name = "Table")]
pub struct PyTable {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyTable {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn dimensions(&self, py: Python<'_>) -> PyResult<(usize, usize)> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "table",
            ".table",
        )?;
        shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .and_then(|shape| shape.table())
            .map(|table| (table.row_count(), table.column_count()))
            .ok_or_else(|| PyValueError::new_err("shape has no table"))
    }
}

#[pymethods]
impl PyTable {
    #[getter]
    fn columns(&self, py: Python<'_>) -> PyResult<Py<PyColumnCollection>> {
        self.dimensions(py)?;
        Py::new(
            py,
            PyColumnCollection {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Py<PyRowCollection>> {
        self.dimensions(py)?;
        Py::new(
            py,
            PyRowCollection {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    fn cell(&self, py: Python<'_>, row: isize, col: isize) -> PyResult<Py<PyCell>> {
        let (rows, columns) = self.dimensions(py)?;
        let row = normalize_index(row, rows, "row")?;
        let column = normalize_index(col, columns, "cell")?;
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Row(row));
        segments.push(PathSeg::Cell(column));
        let path = self.presentation.borrow(py).revisions.capture(segments);
        Py::new(
            py,
            PyCell {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
    }
}

#[pyclass(name = "ColumnCollection")]
pub struct PyColumnCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyColumnCollection {
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "column collection",
            ".table.columns",
        )?;
        shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .and_then(|shape| shape.table())
            .map(|table| table.column_count())
            .ok_or_else(|| PyValueError::new_err("shape has no table"))
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyColumn>> {
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Cell(index));
        let path = self.presentation.borrow(py).revisions.capture(segments);
        Py::new(
            py,
            PyColumn {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
    }
}

#[pymethods]
impl PyColumnCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.len(py)?;
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "column")?)?
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
            "column indices must be integers or slices",
        ))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyColumnIterator>> {
        self.len(py)?;
        Py::new(
            py,
            PyColumnIterator {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                index: 0,
            },
        )
    }

    /// Inserts a grid column before `index`, or appends one, and returns it.
    ///
    /// The column copies the width and cell formatting, without the text, of
    /// the column left of it, or of the first column when it becomes the
    /// first, and the table frame grows by its width. Row, column, cell, and
    /// text handles are invalidated.
    #[pyo3(signature = (index=None))]
    fn add_column(&self, py: Python<'_>, index: Option<isize>) -> PyResult<Py<PyColumn>> {
        let index = insertion_index(index, self.len(py)?, "column")?;
        restructure(
            py,
            &self.presentation,
            &self.path,
            "ColumnCollection.add_column()",
            |table| table.insert_column(index),
        )?;
        self.item(py, index)
    }

    /// Removes one column of this table and shrinks the table frame by its
    /// width. Row, column, cell, and text handles are invalidated.
    fn remove(&self, py: Python<'_>, column: &Bound<'_, PyAny>) -> PyResult<()> {
        self.len(py)?;
        let column = column.extract::<PyRef<'_, PyColumn>>()?;
        if !column.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("column is not in this collection"));
        }
        validate_path(
            py,
            &self.presentation.borrow(py),
            &column.path,
            "column",
            "",
        )?;
        let index = member_index(&self.path, &column.path, |segment| match segment {
            PathSeg::Cell(index) => Some(*index),
            _ => None,
        })
        .ok_or_else(|| PyValueError::new_err("column is not in this collection"))?;
        restructure(
            py,
            &self.presentation,
            &self.path,
            "ColumnCollection.remove()",
            |table| table.remove_column(index),
        )
    }
}

#[pyclass]
struct PyColumnIterator {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyColumnIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyColumn>>> {
        let collection = PyColumnCollection {
            presentation: self.presentation.clone_ref(py),
            path: self.path.clone(),
        };
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

#[pyclass(name = "Column")]
pub struct PyColumn {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyColumn {
    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<i64> {
        validate_path(py, &self.presentation.borrow(py), &self.path, "column", "")?;
        let column = cell_index(&self.path)
            .ok_or_else(|| PyIndexError::new_err("column index is missing"))?;
        shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .and_then(|shape| shape.table())
            .and_then(|table| table.column_width(column))
            .map(|width| width.0)
            .ok_or_else(|| PyIndexError::new_err("column index out of range"))
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, width: i64) -> PyResult<()> {
        validate_path(py, &self.presentation.borrow(py), &self.path, "column", "")?;
        let column = cell_index(&self.path)
            .ok_or_else(|| PyIndexError::new_err("column index is missing"))?;
        shape_mut_at(&mut self.presentation.borrow_mut(py).inner, &self.path)
            .and_then(rpptx::ShapeMut::into_table_mut)
            .ok_or_else(|| PyValueError::new_err("shape has no table"))?
            .set_column_width(column, rpptx::Emu(width))
            .map_err(|error| crate::rpptx_to_pyerr(py, error))
    }
}

#[pyclass(name = "RowCollection")]
pub struct PyRowCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyRowCollection {
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "row collection",
            ".table.rows",
        )?;
        shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .and_then(|shape| shape.table())
            .map(|table| table.row_count())
            .ok_or_else(|| PyValueError::new_err("shape has no table"))
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyRow>> {
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Row(index));
        let path = self.presentation.borrow(py).revisions.capture(segments);
        Py::new(
            py,
            PyRow {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
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
        self.len(py)?;
        Py::new(
            py,
            PyRowIterator {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                index: 0,
            },
        )
    }

    /// Inserts a row before `index`, or appends one, and returns it.
    ///
    /// The row copies the height and cell formatting, without the text, of
    /// the row above it, or of the first row when it becomes the first, and
    /// the table frame grows by its height. Row, column, cell, and text
    /// handles are invalidated.
    #[pyo3(signature = (index=None))]
    fn add_row(&self, py: Python<'_>, index: Option<isize>) -> PyResult<Py<PyRow>> {
        let index = insertion_index(index, self.len(py)?, "row")?;
        restructure(
            py,
            &self.presentation,
            &self.path,
            "RowCollection.add_row()",
            |table| table.insert_row(index),
        )?;
        self.item(py, index)
    }

    /// Removes one row of this table and shrinks the table frame by its
    /// height. Row, column, cell, and text handles are invalidated.
    fn remove(&self, py: Python<'_>, row: &Bound<'_, PyAny>) -> PyResult<()> {
        self.len(py)?;
        let row = row.extract::<PyRef<'_, PyRow>>()?;
        if !row.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("row is not in this collection"));
        }
        validate_path(py, &self.presentation.borrow(py), &row.path, "row", "")?;
        let index = member_index(&self.path, &row.path, |segment| match segment {
            PathSeg::Row(index) => Some(*index),
            _ => None,
        })
        .ok_or_else(|| PyValueError::new_err("row is not in this collection"))?;
        restructure(
            py,
            &self.presentation,
            &self.path,
            "RowCollection.remove()",
            |table| table.remove_row(index),
        )
    }
}

#[pyclass]
struct PyRowIterator {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyRowIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyRow>>> {
        let collection = PyRowCollection {
            presentation: self.presentation.clone_ref(py),
            path: self.path.clone(),
        };
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
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyRow {
    /// The stored row height, a minimum that PowerPoint grows to fit text.
    #[getter]
    fn height(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        validate_path(py, &self.presentation.borrow(py), &self.path, "row", "")?;
        let row =
            row_index(&self.path).ok_or_else(|| PyIndexError::new_err("row index is missing"))?;
        let height = shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .and_then(|shape| shape.table())
            .and_then(|table| table.row_height(row))
            .ok_or_else(|| PyIndexError::new_err("row index out of range"))?;
        length(py, Some(height))
    }

    /// Sets the row height and keeps the table frame height equal to the sum
    /// of its rows.
    #[setter]
    fn set_height(&self, py: Python<'_>, height: i64) -> PyResult<()> {
        validate_path(py, &self.presentation.borrow(py), &self.path, "row", "")?;
        let row =
            row_index(&self.path).ok_or_else(|| PyIndexError::new_err("row index is missing"))?;
        shape_mut_at(&mut self.presentation.borrow_mut(py).inner, &self.path)
            .and_then(rpptx::ShapeMut::into_table_mut)
            .ok_or_else(|| PyValueError::new_err("shape has no table"))?
            .set_row_height(row, rpptx::Emu(height))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pyclass(name = "Cell")]
pub struct PyCell {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyCell {
    fn read<T>(
        &self,
        py: Python<'_>,
        read: impl FnOnce(rpptx::TableCellRef<'_>) -> T,
    ) -> PyResult<T> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "cell", "")?;
        cell_ref_at(&presentation.inner, &self.path)
            .map(read)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }

    fn edit<T>(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rpptx::TableCellMut<'_>) -> T,
    ) -> PyResult<T> {
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(py, &presentation, &self.path, "cell", "")?;
        cell_mut_at(&mut presentation.inner, &self.path)
            .map(|mut cell| edit(&mut cell))
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }

    fn border(&self, py: Python<'_>, edge: rpptx::CellBorder) -> PyResult<Py<PyLineFormat>> {
        self.read(py, |_| ())?;
        Py::new(
            py,
            PyLineFormat::new(
                self.presentation.clone_ref(py),
                self.path.clone(),
                FillTarget::CellBorder(edge),
            ),
        )
    }

    fn margin(
        &self,
        py: Python<'_>,
        side: fn(&mut Margins) -> &mut Option<rpptx::Emu>,
    ) -> PyResult<Option<Py<PyAny>>> {
        let mut margins = self.read(py, |cell| cell.margins())?;
        length(py, *side(&mut margins))
    }

    /// Writes one margin and keeps the other three, leaving the package
    /// unchanged when the value equals the stored one.
    fn set_margin(
        &self,
        py: Python<'_>,
        side: fn(&mut Margins) -> &mut Option<rpptx::Emu>,
        value: Option<i64>,
    ) -> PyResult<()> {
        if value.is_some_and(|value| i32::try_from(value).is_err()) {
            return Err(PyValueError::new_err(
                "cell margin must fit a 32-bit EMU coordinate",
            ));
        }
        let current = self.read(py, |cell| cell.margins())?;
        let mut margins = current;
        *side(&mut margins) = value.map(rpptx::Emu);
        if margins == current {
            return Ok(());
        }
        let (left, right, top, bottom) = margins;
        self.edit(py, |cell| cell.set_margins(left, right, top, bottom))
    }
}

#[pymethods]
impl PyCell {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        validate_path(py, &self.presentation.borrow(py), &self.path, "cell", "")?;
        let row =
            row_index(&self.path).ok_or_else(|| PyIndexError::new_err("row index is missing"))?;
        let cell =
            cell_index(&self.path).ok_or_else(|| PyIndexError::new_err("cell index is missing"))?;
        shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .and_then(|shape| shape.table())
            .and_then(|table| table.cell(row, cell))
            .map(|cell| cell.text())
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        validate_path(py, &self.presentation.borrow(py), &self.path, "cell", "")?;
        let row =
            row_index(&self.path).ok_or_else(|| PyIndexError::new_err("row index is missing"))?;
        let cell =
            cell_index(&self.path).ok_or_else(|| PyIndexError::new_err("cell index is missing"))?;
        shape_mut_at(&mut self.presentation.borrow_mut(py).inner, &self.path)
            .and_then(rpptx::ShapeMut::into_table_mut)
            .and_then(|table| table.into_cell_mut(row, cell))
            .map(|mut cell| cell.set_text(value))
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }

    /// Merges the rectangle between this cell and `other_cell`, moving the
    /// text of the spanned cells into the top-left origin, as python-pptx does.
    fn merge(&self, py: Python<'_>, other_cell: &Bound<'_, PyAny>) -> PyResult<()> {
        let other = other_cell.extract::<PyRef<'_, PyCell>>()?;
        other.read(py, |_| ())?;
        if !other.presentation.is(&self.presentation)
            || !table_segments(&other.path).eq(table_segments(&self.path))
        {
            return Err(PyValueError::new_err("other_cell from different table"));
        }
        let (Some(row), Some(column)) = (row_index(&other.path), cell_index(&other.path)) else {
            return Err(PyIndexError::new_err("cell index is missing"));
        };
        self.edit(py, |cell| cell.merge_to(row, column))?
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    /// Splits this merge origin back into its grid cells.
    fn split(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, |cell| cell.split())?
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[getter]
    fn is_merge_origin(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |cell| cell.is_merge_origin())
    }

    #[getter]
    fn is_spanned(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |cell| cell.is_spanned())
    }

    #[getter]
    fn span_height(&self, py: Python<'_>) -> PyResult<u32> {
        self.read(py, |cell| cell.span_height())
    }

    #[getter]
    fn span_width(&self, py: Python<'_>) -> PyResult<u32> {
        self.read(py, |cell| cell.span_width())
    }

    #[getter]
    fn fill(&self, py: Python<'_>) -> PyResult<Py<PyFillFormat>> {
        self.read(py, |_| ())?;
        Py::new(
            py,
            PyFillFormat::new(
                self.presentation.clone_ref(py),
                self.path.clone(),
                FillTarget::TableCell,
            ),
        )
    }

    #[getter]
    fn margin_left(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.margin(py, |margins| &mut margins.0)
    }

    #[setter]
    fn set_margin_left(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.set_margin(py, |margins| &mut margins.0, value)
    }

    #[getter]
    fn margin_right(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.margin(py, |margins| &mut margins.1)
    }

    #[setter]
    fn set_margin_right(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.set_margin(py, |margins| &mut margins.1, value)
    }

    #[getter]
    fn margin_top(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.margin(py, |margins| &mut margins.2)
    }

    #[setter]
    fn set_margin_top(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.set_margin(py, |margins| &mut margins.2, value)
    }

    #[getter]
    fn margin_bottom(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.margin(py, |margins| &mut margins.3)
    }

    #[setter]
    fn set_margin_bottom(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.set_margin(py, |margins| &mut margins.3, value)
    }

    /// The `a:lnL` border as a live `LineFormat`.
    #[getter]
    fn border_left(&self, py: Python<'_>) -> PyResult<Py<PyLineFormat>> {
        self.border(py, rpptx::CellBorder::Left)
    }

    /// The `a:lnR` border as a live `LineFormat`.
    #[getter]
    fn border_right(&self, py: Python<'_>) -> PyResult<Py<PyLineFormat>> {
        self.border(py, rpptx::CellBorder::Right)
    }

    /// The `a:lnT` border as a live `LineFormat`.
    #[getter]
    fn border_top(&self, py: Python<'_>) -> PyResult<Py<PyLineFormat>> {
        self.border(py, rpptx::CellBorder::Top)
    }

    /// The `a:lnB` border as a live `LineFormat`.
    #[getter]
    fn border_bottom(&self, py: Python<'_>) -> PyResult<Py<PyLineFormat>> {
        self.border(py, rpptx::CellBorder::Bottom)
    }
}
