use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice, PyTuple};
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

pub(crate) fn border_style_from_name(value: &str) -> PyResult<rdocx::BorderStyle> {
    match value {
        "none" => Ok(rdocx::BorderStyle::None),
        "single" => Ok(rdocx::BorderStyle::Single),
        "thick" => Ok(rdocx::BorderStyle::Thick),
        "double" => Ok(rdocx::BorderStyle::Double),
        "dotted" => Ok(rdocx::BorderStyle::Dotted),
        "dashed" => Ok(rdocx::BorderStyle::Dashed),
        "dotDash" => Ok(rdocx::BorderStyle::DotDash),
        "wave" => Ok(rdocx::BorderStyle::Wave),
        _ => Err(PyValueError::new_err(
            "border style must be none, single, thick, double, dotted, dashed, dotDash or wave",
        )),
    }
}

fn table_border_edge(value: &str) -> PyResult<rdocx::TableBorderEdge> {
    match value {
        "top" => Ok(rdocx::TableBorderEdge::Top),
        "bottom" => Ok(rdocx::TableBorderEdge::Bottom),
        "left" => Ok(rdocx::TableBorderEdge::Left),
        "right" => Ok(rdocx::TableBorderEdge::Right),
        "insideH" => Ok(rdocx::TableBorderEdge::InsideHorizontal),
        "insideV" => Ok(rdocx::TableBorderEdge::InsideVertical),
        _ => Err(PyValueError::new_err(
            "border edge must be top, bottom, left, right, insideH or insideV",
        )),
    }
}

fn cell_border_edge(value: &str) -> PyResult<rdocx::CellBorderEdge> {
    match value {
        "top" => Ok(rdocx::CellBorderEdge::Top),
        "bottom" => Ok(rdocx::CellBorderEdge::Bottom),
        "left" => Ok(rdocx::CellBorderEdge::Left),
        "right" => Ok(rdocx::CellBorderEdge::Right),
        "insideH" => Ok(rdocx::CellBorderEdge::InsideHorizontal),
        "insideV" => Ok(rdocx::CellBorderEdge::InsideVertical),
        _ => Err(PyValueError::new_err(
            "border edge must be top, bottom, left, right, insideH or insideV",
        )),
    }
}

/// A border edge as `(style, size in eighths of a point, color)`.
type BorderSnapshot = (String, Option<u32>, Option<String>);

fn border_snapshot(border: rdocx::TableBorderRef<'_>) -> BorderSnapshot {
    (
        border.style().to_owned(),
        border.size_eighths_pt(),
        border.color().map(str::to_owned),
    )
}

/// Cell margins as `(top, right, bottom, left)`, each a `Length` or `None`.
type MarginSnapshot = (
    Option<Py<PyAny>>,
    Option<Py<PyAny>>,
    Option<Py<PyAny>>,
    Option<Py<PyAny>>,
);

fn margin_snapshot(py: Python<'_>, margins: rdocx::TableCellMargins) -> PyResult<MarginSnapshot> {
    let length =
        |value: Option<rdocx::Length>| value.map(|value| length_object(py, value)).transpose();
    Ok((
        length(margins.top)?,
        length(margins.right)?,
        length(margins.bottom)?,
        length(margins.left)?,
    ))
}

/// Read a `WD_ROW_HEIGHT_RULE` value, where `AT_LEAST` is 1 and `EXACTLY` is 2.
fn row_height_rule_is_exact(value: i32) -> PyResult<bool> {
    match value {
        1 => Ok(false),
        2 => Ok(true),
        _ => Err(PyValueError::new_err("unsupported row height rule")),
    }
}

fn row_height(length: rdocx::Length, exact: bool) -> rdocx::RowHeight {
    if exact {
        rdocx::RowHeight::Exact(length)
    } else {
        rdocx::RowHeight::AtLeast(length)
    }
}

fn row_height_parts(value: rdocx::RowHeight) -> (rdocx::Length, bool) {
    match value {
        rdocx::RowHeight::AtLeast(length) => (length, false),
        rdocx::RowHeight::Exact(length) => (length, true),
    }
}

fn vertical_merge_from_name(value: Option<&str>) -> PyResult<Option<rdocx::VMerge>> {
    match value {
        None => Ok(None),
        Some("restart") => Ok(Some(rdocx::VMerge::Restart)),
        Some("continue") => Ok(Some(rdocx::VMerge::Continue)),
        Some(_) => Err(PyValueError::new_err(
            "vertical merge must be restart, continue or None",
        )),
    }
}

fn vertical_merge_name(value: rdocx::VMerge) -> &'static str {
    match value {
        rdocx::VMerge::Restart => "restart",
        rdocx::VMerge::Continue => "continue",
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

    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<usize> {
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

    pub(crate) fn belongs_to(&self, py: Python<'_>, document: &Py<PyDocument>) -> bool {
        self.document.bind(py).is(document.bind(py))
    }

    /// Apply one checked native table edit without advancing the revision.
    fn edit<T>(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Table<'_>) -> rdocx::Result<T>,
    ) -> PyResult<T> {
        let index = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = document
            .inner
            .table_mut(index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        edit(&mut table).map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Resolve possibly negative row and cell indexes against this table.
    fn cell_coordinates(&self, py: Python<'_>, row: isize, col: isize) -> PyResult<(usize, usize)> {
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
        Ok((row, col))
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
        let (row, col) = self.cell_coordinates(py, row, col)?;
        let table_index = table_index(&self.path)?;
        let path = self.document.borrow(py).revisions.capture(smallvec![
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

    #[pyo3(signature = (style, *, size, color))]
    fn set_borders(&self, py: Python<'_>, style: &str, size: u32, color: &str) -> PyResult<()> {
        let style = border_style_from_name(style)?;
        self.edit(py, |table| {
            table.set_all_borders_checked(style, size, color)
        })
    }

    #[pyo3(signature = (edge, style, *, size, color))]
    fn set_border(
        &self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        size: u32,
        color: &str,
    ) -> PyResult<()> {
        let edge = table_border_edge(edge)?;
        let style = border_style_from_name(style)?;
        self.edit(py, |table| {
            table.set_border_checked(edge, style, size, color)
        })
    }

    fn border(&self, py: Python<'_>, edge: &str) -> PyResult<Option<BorderSnapshot>> {
        let edge = table_border_edge(edge)?;
        let index = self.validate(py)?;
        Ok(self
            .document
            .borrow(py)
            .inner
            .table(index)
            .and_then(|table| table.border(edge).map(border_snapshot)))
    }

    #[getter]
    fn cell_margins(&self, py: Python<'_>) -> PyResult<Option<MarginSnapshot>> {
        let index = self.validate(py)?;
        let margins = self
            .document
            .borrow(py)
            .inner
            .table(index)
            .and_then(|table| table.cell_margins());
        margins
            .map(|margins| margin_snapshot(py, margins))
            .transpose()
    }

    #[pyo3(signature = (*, top, right, bottom, left))]
    fn set_cell_margins(
        &self,
        py: Python<'_>,
        top: i64,
        right: i64,
        bottom: i64,
        left: i64,
    ) -> PyResult<()> {
        self.edit(py, |table| {
            table.set_cell_margins_checked(
                rdocx::Length::emu(top),
                rdocx::Length::emu(right),
                rdocx::Length::emu(bottom),
                rdocx::Length::emu(left),
            )
        })
    }

    #[getter]
    fn grid_widths<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let index = self.validate(py)?;
        let widths = self
            .document
            .borrow(py)
            .inner
            .table(index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
            .grid_widths();
        let widths = widths
            .into_iter()
            .map(|width| length_object(py, width))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, widths)
    }

    #[setter]
    fn set_grid_widths(&self, py: Python<'_>, value: Vec<i64>) -> PyResult<()> {
        let widths = value
            .into_iter()
            .map(rdocx::Length::emu)
            .collect::<Vec<_>>();
        self.edit(py, |table| table.set_grid_widths(&widths))
    }

    fn set_column_width(&self, py: Python<'_>, column: isize, width: i64) -> PyResult<()> {
        let index = self.validate(py)?;
        let columns = self
            .document
            .borrow(py)
            .inner
            .table(index)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?
            .grid_widths()
            .len();
        let column = normalize_index(column, columns, "column")?;
        let applied = self.edit(py, |table| {
            Ok(table.set_column_width(column, rdocx::Length::emu(width)))
        })?;
        if !applied {
            return Err(PyValueError::new_err(
                "column width must be nonnegative and every row must fit the table grid",
            ));
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, span))]
    fn set_cell_grid_span(
        &self,
        py: Python<'_>,
        row: isize,
        col: isize,
        span: Option<u32>,
    ) -> PyResult<()> {
        let (row, col) = self.cell_coordinates(py, row, col)?;
        let cells = |table: &mut rdocx::Table<'_>| table.row(row).map(|row| row.cell_count());
        let changed = self.edit(py, |table| {
            let before = cells(table);
            table.set_cell_grid_span_checked(row, col, span)?;
            Ok(cells(table) != before)
        })?;
        // Absorbed or restored cells shift the indexes after this one.
        if changed {
            self.document.borrow_mut(py).revisions.bump();
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, merge))]
    fn set_cell_vertical_merge(
        &self,
        py: Python<'_>,
        row: isize,
        col: isize,
        merge: Option<&str>,
    ) -> PyResult<()> {
        let merge = vertical_merge_from_name(merge)?;
        let (row, col) = self.cell_coordinates(py, row, col)?;
        self.edit(py, |table| table.set_cell_vertical_merge(row, col, merge))
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
            let source = normalize_index(index, row_count, "row")?;
            let insert_at = at.unwrap_or(source + 1);
            if insert_at > row_count {
                return Err(PyIndexError::new_err("row insertion index out of range"));
            }
            let inserted = document
                .inner
                .clone_table_row(table_index, source, insert_at)
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            document.revisions.bump();
            let mut segments = self.path.segs.clone();
            segments.push(PathSeg::Row(inserted));
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
        let row_index = normalize_index(index, row_count, "row")?;
        document
            .inner
            .remove_table_row(table_index, row_index)
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

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::RowRef<'_>) -> T) -> PyResult<T> {
        let (table, row) = self.validate(py)?;
        let document = self.document.borrow(py);
        let table = document
            .inner
            .table(table)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        let row = table
            .row(row)
            .ok_or_else(|| PyIndexError::new_err("row index out of range"))?;
        Ok(read(row))
    }

    /// Apply one checked native row edit that moves no content, so live
    /// handles stay valid.
    fn edit(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Row<'_>) -> rdocx::Result<()>,
    ) -> PyResult<()> {
        let (table, row) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = document
            .inner
            .table_mut(table)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        let mut row = table
            .row(row)
            .ok_or_else(|| PyIndexError::new_err("row index out of range"))?;
        edit(&mut row).map_err(|error| rdocx_to_pyerr(py, error))
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

    #[getter]
    fn height(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |row| row.height())?
            .map(|height| length_object(py, row_height_parts(height).0))
            .transpose()
    }

    // An exact height keeps its rule. Any other row gets a minimum height,
    // including one whose w:trHeight has an auto rule or no value, which the
    // native reader reports as no height.
    #[setter]
    fn set_height(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        let exact = self
            .read(py, |row| row.height())?
            .is_some_and(|height| row_height_parts(height).1);
        let height = row_height(rdocx::Length::emu(value), exact);
        self.edit(py, |row| row.set_height_checked(height))
    }

    #[getter]
    fn height_rule(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |row| row.height())?
            .map(|height| {
                let exact = row_height_parts(height).1;
                enum_object(py, "WD_ROW_HEIGHT_RULE", if exact { 2 } else { 1 })
            })
            .transpose()
    }

    #[setter]
    fn set_height_rule(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let exact = row_height_rule_is_exact(value)?;
        let length = self
            .read(py, |row| row.height())?
            .map(|height| row_height_parts(height).0)
            .ok_or_else(|| {
                PyValueError::new_err(
                    "row has no height for the rule to apply to, set height first",
                )
            })?;
        self.edit(py, |row| row.set_height_checked(row_height(length, exact)))
    }

    #[getter]
    fn cant_split(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.read(py, |row| row.cant_split_value())
    }

    #[setter]
    fn set_cant_split(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.edit(py, |row| {
            row.set_cant_split_value(value);
            Ok(())
        })
    }

    #[getter]
    fn is_header(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.read(py, |row| row.header_value())
    }

    #[setter]
    fn set_is_header(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.edit(py, |row| {
            row.set_header_value(value);
            Ok(())
        })
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

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::CellRef<'_>) -> T) -> PyResult<T> {
        let (table, row, cell) = self.validate(py)?;
        let document = self.document.borrow(py);
        let table = document
            .inner
            .table(table)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        let cell = table
            .cell(row, cell)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
        Ok(read(cell))
    }

    /// Apply one checked native cell edit that moves no content, so live
    /// handles stay valid.
    fn edit(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Cell<'_>) -> rdocx::Result<()>,
    ) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = document
            .inner
            .table_mut(table)
            .ok_or_else(|| PyIndexError::new_err("table index out of range"))?;
        let mut cell = table
            .cell(row, cell)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
        edit(&mut cell).map_err(|error| rdocx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyCell {
    /// Replace literal text in this cell and its supported nested descendants.
    #[pyo3(signature = (old, new, *, expect = None))]
    fn replace_text(
        &self,
        py: Python<'_>,
        old: &str,
        new: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        let cell = self.validate(py)?;
        self.document
            .borrow_mut(py)
            .scoped_replacement(py, |document| {
                document.try_replace_text_in_cell(cell, None, old, new, expect)
            })
    }

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
        let inner = &mut document.inner;
        py.detach(|| inner.try_set_cell_text(table, row, cell, value))
            .map_err(|error| rdocx_to_pyerr(py, error))?;
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

    #[getter]
    fn grid_span(&self, py: Python<'_>) -> PyResult<u32> {
        self.read(py, |cell| cell.grid_span().unwrap_or(1))
    }

    #[getter]
    fn vertical_merge(&self, py: Python<'_>) -> PyResult<Option<&'static str>> {
        self.read(py, |cell| cell.v_merge().copied().map(vertical_merge_name))
    }

    #[getter]
    fn shading(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.read(py, |cell| cell.shading_fill().map(str::to_owned))
    }

    #[setter]
    fn set_shading(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.edit(py, |cell| cell.set_shading_checked(value))
    }

    fn border(&self, py: Python<'_>, edge: &str) -> PyResult<Option<BorderSnapshot>> {
        let edge = cell_border_edge(edge)?;
        self.read(py, |cell| cell.border(edge).map(border_snapshot))
    }

    #[pyo3(signature = (edge, style, *, size, color))]
    fn set_border(
        &self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        size: u32,
        color: &str,
    ) -> PyResult<()> {
        let edge = cell_border_edge(edge)?;
        let style = border_style_from_name(style)?;
        self.edit(py, |cell| cell.set_border_checked(edge, style, size, color))
    }

    #[getter]
    fn margins(&self, py: Python<'_>) -> PyResult<Option<MarginSnapshot>> {
        self.read(py, |cell| cell.margins())?
            .map(|margins| margin_snapshot(py, margins))
            .transpose()
    }

    #[pyo3(signature = (*, top, right, bottom, left))]
    fn set_margins(
        &self,
        py: Python<'_>,
        top: i64,
        right: i64,
        bottom: i64,
        left: i64,
    ) -> PyResult<()> {
        self.edit(py, |cell| {
            cell.set_margins_checked(
                rdocx::Length::emu(top),
                rdocx::Length::emu(right),
                rdocx::Length::emu(bottom),
                rdocx::Length::emu(left),
            )
        })
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
