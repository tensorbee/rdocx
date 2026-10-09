use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{
    PyIndexError, PyKeyError, PyNotImplementedError, PyTypeError, PyUserWarning, PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice, PyTuple};
use smallvec::smallvec;

use crate::document::PyDocument;
use crate::paragraph::{PyParagraph, style_id_of_type};
use crate::{enum_object, length_object, normalize_index, rdocx_to_pyerr, stale_to_pyerr};

/// Where a table handle points: a table counted as `Document::table`
/// counts, then, for a nested table, the `(row, cell, nested table)` steps
/// into it.
#[derive(Clone)]
struct TableAddress {
    table: usize,
    nested: Vec<(usize, usize, usize)>,
}

impl TableAddress {
    /// Split a table, row or cell path into its table address and the row
    /// and cell indexes that follow it.
    ///
    /// A table path is `Body(table)`, and each nested table adds
    /// `Row, Cell, Body(nested table)` after the cell that holds it.
    fn parse(path: &ContentPath) -> PyResult<(Self, Option<usize>, Option<usize>)> {
        let mut segments = path.segs.iter().copied();
        let Some(PathSeg::Body(table)) = segments.next() else {
            return Err(PyIndexError::new_err("table index is missing"));
        };
        let mut address = Self {
            table,
            nested: Vec::new(),
        };
        let (mut row, mut cell) = (None, None);
        for segment in segments {
            match segment {
                PathSeg::Row(index) => row = Some(index),
                PathSeg::Cell(index) => cell = Some(index),
                PathSeg::Body(index) => {
                    let (Some(row), Some(cell)) = (row.take(), cell.take()) else {
                        return Err(PyIndexError::new_err("nested table path is incomplete"));
                    };
                    address.nested.push((row, cell, index));
                }
                _ => break,
            }
        }
        Ok((address, row, cell))
    }

    fn get<'d>(&self, document: &'d rdocx::Document) -> PyResult<rdocx::TableRef<'d>> {
        let mut table = document.table(self.table);
        for &(row, cell, index) in &self.nested {
            table = table.and_then(|table| table.nested_table(row, cell, index));
        }
        table.ok_or_else(|| PyIndexError::new_err("table index out of range"))
    }

    fn get_mut<'d>(&self, document: &'d mut rdocx::Document) -> PyResult<rdocx::Table<'d>> {
        let mut table = document.table_mut(self.table);
        for &(row, cell, index) in &self.nested {
            table = table.and_then(|table| table.into_nested_table(row, cell, index));
        }
        table.ok_or_else(|| PyIndexError::new_err("table index out of range"))
    }

    /// The table index of a top-level table, for the edits that run as a
    /// checked document transaction and address top-level tables only.
    fn top_level(&self, operation: &str) -> PyResult<usize> {
        if self.nested.is_empty() {
            Ok(self.table)
        } else {
            Err(PyNotImplementedError::new_err(format!(
                "{operation} works on top-level tables only. In a nested table, set \
                 text with cell.text, add rows with add_row and columns with insert_column"
            )))
        }
    }
}

/// What a revision bump inside one table moved, so that table, row and cell
/// handles it cannot have moved stay valid.
pub(crate) struct TableEdit {
    table: Vec<PathSeg>,
    scope: TableEditScope,
}

enum TableEditScope {
    /// Rows or columns appended at the end: no existing index moved.
    Appended,
    /// Rows, columns or cells moved: only the table itself keeps its path.
    Grid,
    /// One cell's content changed: paths below that cell moved.
    Cell(usize, usize),
}

impl TableEdit {
    fn keeps(&self, path: &[PathSeg]) -> bool {
        let below = |prefix: &[PathSeg]| path.len() > prefix.len() && path.starts_with(prefix);
        match self.scope {
            TableEditScope::Appended => true,
            TableEditScope::Grid => !below(&self.table),
            TableEditScope::Cell(row, cell) => {
                let mut prefix = self.table.clone();
                prefix.extend([PathSeg::Row(row), PathSeg::Cell(cell)]);
                !below(&prefix)
            }
        }
    }
}

/// Check a table, row or cell handle: it is valid at its own revision and
/// after every later bump that was a table edit leaving its path in place.
fn check_table_path(
    py: Python<'_>,
    document: &PyDocument,
    path: &ContentPath,
    kind: &str,
    hint: &str,
) -> PyResult<()> {
    let current = document.revisions.current();
    let kept = path.revision <= current
        && (path.revision + 1..=current).all(|revision| {
            document
                .table_edits
                .get(&revision)
                .is_some_and(|edit| edit.keeps(&path.segs))
        });
    if kept {
        return Ok(());
    }
    path.validate_revision(current, kind, hint)
        .map_err(|error| stale_to_pyerr(py, error))
}

/// Advance the revision for an edit inside the table at `table`.
fn bump_table(document: &mut PyDocument, table: &[PathSeg], scope: TableEditScope) {
    let revision = document.revisions.bump();
    document.table_edits.insert(
        revision,
        TableEdit {
            table: table.to_vec(),
            scope,
        },
    );
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

fn border_style_from_name(value: &str) -> PyResult<rdocx::BorderStyle> {
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

    /// Check the handle's revision and resolve where its table is.
    fn address(&self, py: Python<'_>) -> PyResult<TableAddress> {
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.path,
            "table",
            "Re-fetch it with doc.tables[i] or cell.tables[i].",
        )?;
        Ok(TableAddress::parse(&self.path)?.0)
    }

    /// The index of this top-level table among `doc.tables`.
    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let address = self.address(py)?;
        if !address.nested.is_empty() {
            return Err(PyValueError::new_err(
                "a nested table is not a direct body child",
            ));
        }
        Ok(address.table)
    }

    pub(crate) fn belongs_to(&self, py: Python<'_>, document: &Py<PyDocument>) -> bool {
        self.document.bind(py).is(document.bind(py))
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::TableRef<'_>) -> T) -> PyResult<T> {
        let address = self.address(py)?;
        let document = self.document.borrow(py);
        Ok(read(address.get(&document.inner)?))
    }

    /// Apply one checked native table edit without advancing the revision.
    fn edit<T>(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Table<'_>) -> rdocx::Result<T>,
    ) -> PyResult<T> {
        let address = self.address(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = address.get_mut(&mut document.inner)?;
        edit(&mut table).map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Apply one native edit that adds or moves rows or cells, then advance
    /// the revision. This table handle stays valid, and so do the row and
    /// cell handles the edit cannot have moved.
    fn edit_structure<T>(
        &self,
        py: Python<'_>,
        scope: TableEditScope,
        edit: impl FnOnce(&mut rdocx::Table<'_>) -> rdocx::Result<T>,
    ) -> PyResult<T> {
        let value = self.edit(py, edit)?;
        self.bump(py, scope);
        Ok(value)
    }

    fn bump(&self, py: Python<'_>, scope: TableEditScope) {
        bump_table(&mut self.document.borrow_mut(py), &self.path.segs, scope);
    }

    /// Capture a path below this table at the current revision.
    fn child_path(&self, py: Python<'_>, segments: &[PathSeg]) -> ContentPath {
        let mut path = self.path.segs.clone();
        path.extend_from_slice(segments);
        self.document.borrow(py).revisions.capture(path)
    }

    fn column(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyColumn>> {
        // The column sits below the table, so an edit that moves cells
        // retires it as it retires row and cell handles.
        let path = self.child_path(py, &[PathSeg::Row(usize::MAX)]);
        Py::new(
            py,
            PyColumn {
                document: self.document.clone_ref(py),
                path,
                index,
            },
        )
    }

    /// Resolve possibly negative row and cell indexes against this table.
    fn cell_coordinates(&self, py: Python<'_>, row: isize, col: isize) -> PyResult<(usize, usize)> {
        let (rows, cells) = self.read(py, |table| {
            let rows = table.row_count();
            let cells = (0..rows)
                .map(|row| table.row(row).map_or(0, |row| row.cell_count()))
                .collect::<Vec<_>>();
            (rows, cells)
        })?;
        let row = normalize_index(row, rows, "row")?;
        let col = normalize_index(col, cells[row], "cell")?;
        Ok((row, col))
    }

    fn look_flag(&self, py: Python<'_>, flag: fn(&rdocx::TableLook) -> bool) -> PyResult<bool> {
        self.read(py, |table| flag(&table_look(&table)))
    }

    fn set_look_flag(
        &self,
        py: Python<'_>,
        set: impl FnOnce(&mut rdocx::TableLook),
    ) -> PyResult<()> {
        let mut look = self.read(py, |table| table_look(&table))?;
        set(&mut look);
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }
}

/// The table's `w:tblLook` selection. An absent element selects no
/// conditional region and leaves banding on, which is what its attribute
/// defaults mean.
fn table_look(table: &rdocx::TableRef<'_>) -> rdocx::TableLook {
    table.look().unwrap_or(rdocx::TableLook {
        horizontal_banding: true,
        vertical_banding: true,
        ..Default::default()
    })
}

#[pymethods]
impl PyTable {
    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Py<PyRowCollection>> {
        self.address(py)?;
        Py::new(
            py,
            PyRowCollection::new(self.document.clone_ref(py), self.path.clone()),
        )
    }

    fn cell(&self, py: Python<'_>, row: isize, col: isize) -> PyResult<Py<PyCell>> {
        let (row, col) = self.cell_coordinates(py, row, col)?;
        let path = self.child_path(py, &[PathSeg::Row(row), PathSeg::Cell(col)]);
        Py::new(py, PyCell::new(self.document.clone_ref(py), path))
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.read(py, |table| table.style_id().map(str::to_owned))
    }

    /// Set the table style from its style ID or, as python-docx accepts, its
    /// name.
    #[setter]
    fn set_style(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        // Word draws a table whose style the package does not define with
        // the default table style. python-docx writes such a reference (its
        // template defines Word's built-in table styles), so an unknown value
        // is written as given and a UserWarning says it will not show.
        let style = match style_id_of_type(
            &self.document.borrow(py).inner,
            value,
            rdocx::StyleType::Table,
        ) {
            Ok(style) => style,
            Err(error) if error.is_instance_of::<PyKeyError>(py) => {
                PyErr::warn(
                    py,
                    &py.get_type::<PyUserWarning>(),
                    &std::ffi::CString::new(format!(
                        "no table style is named or identified by '{value}' in this document: \
                         Word and Google Docs draw the table with the default table style. \
                         Define it with Document.add_style(..., style_type=\"table\") or pick \
                         one of Document.styles"
                    ))?,
                    1,
                )?;
                value.to_owned()
            }
            Err(error) => return Err(error),
        };
        self.edit(py, |table| {
            table.set_style(&style);
            Ok(())
        })
    }

    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |table| table.alignment())?
            .and_then(alignment_to_int)
            .map(|value| enum_object(py, "WD_TABLE_ALIGNMENT", value))
            .transpose()
    }

    #[setter]
    fn set_alignment(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let value = alignment_from_int(value)?;
        self.edit(py, |table| {
            table.set_alignment(value);
            Ok(())
        })
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |table| table.width())?
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.edit(py, |table| {
            table.set_width(rdocx::Length::emu(value));
            Ok(())
        })
    }

    /// Whether Word may resize columns to fit their content, as python-docx
    /// `Table.autofit` reads it. `False` writes a fixed layout.
    #[getter]
    fn autofit(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |table| {
            table.layout() != Some(rdocx::TableLayout::Fixed)
        })
    }

    #[setter]
    fn set_autofit(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.edit(py, |table| {
            table.set_layout(if value {
                rdocx::TableLayout::AutoFit
            } else {
                rdocx::TableLayout::Fixed
            });
            Ok(())
        })
    }

    #[getter]
    fn first_row(&self, py: Python<'_>) -> PyResult<bool> {
        self.look_flag(py, |look| look.first_row)
    }

    #[setter]
    fn set_first_row(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_look_flag(py, |look| look.first_row = value)
    }

    #[getter]
    fn last_row(&self, py: Python<'_>) -> PyResult<bool> {
        self.look_flag(py, |look| look.last_row)
    }

    #[setter]
    fn set_last_row(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_look_flag(py, |look| look.last_row = value)
    }

    #[getter]
    fn first_col(&self, py: Python<'_>) -> PyResult<bool> {
        self.look_flag(py, |look| look.first_column)
    }

    #[setter]
    fn set_first_col(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_look_flag(py, |look| look.first_column = value)
    }

    #[getter]
    fn last_col(&self, py: Python<'_>) -> PyResult<bool> {
        self.look_flag(py, |look| look.last_column)
    }

    #[setter]
    fn set_last_col(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_look_flag(py, |look| look.last_column = value)
    }

    #[getter]
    fn horz_banding(&self, py: Python<'_>) -> PyResult<bool> {
        self.look_flag(py, |look| look.horizontal_banding)
    }

    #[setter]
    fn set_horz_banding(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_look_flag(py, |look| look.horizontal_banding = value)
    }

    #[getter]
    fn vert_banding(&self, py: Python<'_>) -> PyResult<bool> {
        self.look_flag(py, |look| look.vertical_banding)
    }

    #[setter]
    fn set_vert_banding(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_look_flag(py, |look| look.vertical_banding = value)
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
        self.read(py, |table| table.border(edge).map(border_snapshot))
    }

    #[getter]
    fn cell_margins(&self, py: Python<'_>) -> PyResult<Option<MarginSnapshot>> {
        self.read(py, |table| table.cell_margins())?
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
        let widths = self
            .read(py, |table| table.grid_widths())?
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
        let columns = self.read(py, |table| table.column_count())?;
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
            self.bump(py, TableEditScope::Grid);
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

    /// Append an empty row with the last row's row and cell formatting, as
    /// python-docx `Table.add_row` does. The table and its other row and
    /// cell handles stay valid.
    fn add_row(&self, py: Python<'_>) -> PyResult<Py<PyRow>> {
        let row = self.edit_structure(py, TableEditScope::Appended, |table| {
            table.add_row()?;
            Ok(table.row_count() - 1)
        })?;
        let path = self.child_path(py, &[PathSeg::Row(row)]);
        Py::new(py, PyRow::new(self.document.clone_ref(py), path))
    }

    /// Append a grid column, `width` wide or as wide as the last column, and
    /// return it, as python-docx `Table.add_column` does. The table and its
    /// row and cell handles stay valid.
    #[pyo3(signature = (width = None))]
    fn add_column(&self, py: Python<'_>, width: Option<i64>) -> PyResult<Py<PyColumn>> {
        let columns = self.read(py, |table| table.column_count())?;
        self.insert_column(py, columns as isize, width)?;
        self.column(py, columns)
    }

    /// The grid columns, as python-docx `Table.columns` lists them.
    #[getter]
    fn columns<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let columns = self.read(py, |table| table.column_count())?;
        let columns = (0..columns)
            .map(|index| self.column(py, index))
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, columns)
    }

    /// Insert a grid column before grid column `index`, or append one when
    /// `index` is the column count.
    #[pyo3(signature = (index, width = None))]
    fn insert_column(&self, py: Python<'_>, index: isize, width: Option<i64>) -> PyResult<()> {
        let columns = self.read(py, |table| table.column_count())?;
        let index = if index == columns as isize {
            columns
        } else {
            normalize_index(index, columns, "column")?
        };
        // An appended column moves no cell, an inserted one shifts the
        // cells after it.
        let scope = if index == columns {
            TableEditScope::Appended
        } else {
            TableEditScope::Grid
        };
        self.edit_structure(py, scope, |table| {
            table.insert_column(index, width.map(rdocx::Length::emu))
        })
    }

    /// Remove grid column `index` with the cells that cover only it.
    fn remove_column(&self, py: Python<'_>, index: isize) -> PyResult<()> {
        let table = self.address(py)?.top_level("remove_column")?;
        let columns = self.read(py, |table| table.column_count())?;
        let column = normalize_index(index, columns, "column")?;
        let mut document = self.document.borrow_mut(py);
        document
            .inner
            .remove_table_column(table, column)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        bump_table(&mut document, &self.path.segs, TableEditScope::Grid);
        Ok(())
    }

    #[pyo3(signature = (index, at = None))]
    fn clone_row(&self, py: Python<'_>, index: isize, at: Option<usize>) -> PyResult<Py<PyRow>> {
        let table_index = self.address(py)?.top_level("clone_row")?;
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
            bump_table(&mut document, &self.path.segs, TableEditScope::Grid);
            let mut segments = self.path.segs.clone();
            segments.push(PathSeg::Row(inserted));
            document.revisions.capture(segments)
        };
        Py::new(py, PyRow::new(self.document.clone_ref(py), path))
    }

    fn remove_row(&self, py: Python<'_>, index: isize) -> PyResult<()> {
        let table_index = self.address(py)?.top_level("remove_row")?;
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
        bump_table(&mut document, &self.path.segs, TableEditScope::Grid);
        Ok(())
    }
}

/// One grid column of a table, as python-docx `_Column` is.
#[pyclass(name = "Column")]
pub struct PyColumn {
    document: Py<PyDocument>,
    /// The table path followed by a placeholder row.
    path: ContentPath,
    index: usize,
}

impl PyColumn {
    fn table(&self, py: Python<'_>) -> PyResult<PyTable> {
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.path,
            "column",
            "Re-fetch it with table.columns[i].",
        )?;
        let segments = self.path.segs[..self.path.segs.len() - 1].to_vec();
        let path = ContentPath::new(segments.into_iter().collect(), self.path.revision);
        Ok(PyTable::new(self.document.clone_ref(py), path))
    }
}

#[pymethods]
impl PyColumn {
    #[getter]
    fn index(&self) -> usize {
        self.index
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let width = self
            .table(py)?
            .read(py, |table| table.grid_widths().get(self.index).copied())?;
        width.map(|width| length_object(py, width)).transpose()
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.table(py)?
            .set_column_width(py, self.index as isize, value)
    }

    /// The cell covering this grid column in each row, skipping rows that
    /// omit it. A merged cell appears in each column it spans.
    #[getter]
    fn cells<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let table = self.table(py)?;
        let positions = table.read(py, |table| {
            (0..table.row_count())
                .filter_map(|row| {
                    let cells = table.row(row)?;
                    let mut start = cells.grid_before().unwrap_or(0) as usize;
                    (0..cells.cell_count()).find_map(|cell| {
                        let span = cells.cell(cell)?.grid_span().unwrap_or(1).max(1) as usize;
                        let found = (start..start + span).contains(&self.index);
                        start += span;
                        found.then_some((row, cell))
                    })
                })
                .collect::<Vec<_>>()
        })?;
        let cells = positions
            .into_iter()
            .map(|(row, cell)| {
                let path = table.child_path(py, &[PathSeg::Row(row), PathSeg::Cell(cell)]);
                Py::new(py, PyCell::new(self.document.clone_ref(py), path))
            })
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, cells)
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
    fn validate(&self, py: Python<'_>) -> PyResult<TableAddress> {
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.table_path,
            "row collection",
            "Re-fetch it with table.rows.",
        )?;
        Ok(TableAddress::parse(&self.table_path)?.0)
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let address = self.validate(py)?;
        Ok(address.get(&self.document.borrow(py).inner)?.row_count())
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
    fn validate(&self, py: Python<'_>) -> PyResult<(TableAddress, usize)> {
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.path,
            "row",
            "Re-fetch it with table.rows[i].",
        )?;
        let (address, row, _) = TableAddress::parse(&self.path)?;
        Ok((
            address,
            row.ok_or_else(|| PyIndexError::new_err("row index is missing"))?,
        ))
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::RowRef<'_>) -> T) -> PyResult<T> {
        let (address, row) = self.validate(py)?;
        let document = self.document.borrow(py);
        let table = address.get(&document.inner)?;
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
        let (address, row) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = address.get_mut(&mut document.inner)?;
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
    fn validate(&self, py: Python<'_>) -> PyResult<(TableAddress, usize)> {
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.row_path,
            "cell collection",
            "Re-fetch it with row.cells.",
        )?;
        let (address, row, _) = TableAddress::parse(&self.row_path)?;
        Ok((
            address,
            row.ok_or_else(|| PyIndexError::new_err("row index is missing"))?,
        ))
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let (address, row) = self.validate(py)?;
        address
            .get(&self.document.borrow(py).inner)?
            .row(row)
            .map(|row| row.cell_count())
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
    fn validate(&self, py: Python<'_>) -> PyResult<(TableAddress, usize, usize)> {
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.path,
            "cell",
            "Re-fetch it with row.cells[i].",
        )?;
        match TableAddress::parse(&self.path)? {
            (address, Some(row), Some(cell)) => Ok((address, row, cell)),
            _ => Err(PyIndexError::new_err("cell path is incomplete")),
        }
    }

    /// The `(table, row, cell)` coordinates of a cell of a top-level table,
    /// for the edits that only address top-level tables.
    fn top_level(&self, py: Python<'_>, operation: &str) -> PyResult<(usize, usize, usize)> {
        let (address, row, cell) = self.validate(py)?;
        Ok((address.top_level(operation)?, row, cell))
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::CellRef<'_>) -> T) -> PyResult<T> {
        let (address, row, cell) = self.validate(py)?;
        let document = self.document.borrow(py);
        let table = address.get(&document.inner)?;
        let cell = table
            .cell(row, cell)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
        Ok(read(cell))
    }

    /// Apply one checked native edit to the table that holds this cell.
    fn edit_table<T>(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Table<'_>, usize, usize) -> rdocx::Result<T>,
    ) -> PyResult<T> {
        let (address, row, cell) = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut table = address.get_mut(&mut document.inner)?;
        if table.cell(row, cell).is_none() {
            return Err(PyIndexError::new_err("cell index out of range"));
        }
        edit(&mut table, row, cell).map_err(|error| rdocx_to_pyerr(py, error))
    }

    /// Apply one checked native cell edit that moves no content, so live
    /// handles stay valid.
    fn edit(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Cell<'_>) -> rdocx::Result<()>,
    ) -> PyResult<()> {
        self.edit_table(py, |table, row, cell| {
            edit(&mut table.cell(row, cell).expect("the cell was checked"))
        })
    }

    /// Advance the revision for an edit inside this cell, or across its
    /// table for `TableEditScope::Grid`. Handles outside what the edit
    /// moved stay valid.
    fn bump(&self, py: Python<'_>, grid: bool) -> PyResult<()> {
        let (_, row, cell) = self.validate(py)?;
        let table = &self.path.segs[..self.path.segs.len() - 2];
        let scope = if grid {
            TableEditScope::Grid
        } else {
            TableEditScope::Cell(row, cell)
        };
        bump_table(&mut self.document.borrow_mut(py), table, scope);
        Ok(())
    }

    /// A handle on the `index`-th table nested directly in this cell.
    fn nested_table(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyTable>> {
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Body(index));
        let path = self.document.borrow(py).revisions.capture(segments);
        Py::new(py, PyTable::new(self.document.clone_ref(py), path))
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
        let cell = self.top_level(py, "Cell.replace_text")?;
        self.document
            .borrow_mut(py)
            .scoped_replacement(py, |document| {
                document.try_replace_text_in_cell(cell, None, old, new, expect)
            })
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.read(py, |cell| cell.text())
    }
    #[setter]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let (address, row, cell) = self.validate(py)?;
        if address.nested.is_empty() {
            let mut document = self.document.borrow_mut(py);
            let inner = &mut document.inner;
            py.detach(|| inner.try_set_cell_text(address.table, row, cell, value))
                .map_err(|error| rdocx_to_pyerr(py, error))?;
        } else {
            self.edit(py, |cell| cell.try_set_text(value))?;
        }
        // Only handles inside this cell moved, so its siblings, its row and
        // its table stay valid, as in `table.add_row().cells` loops.
        self.bump(py, false)
    }
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> PyResult<Py<PyCellParagraphCollection>> {
        self.top_level(py, "Cell.paragraphs")?;
        Py::new(
            py,
            PyCellParagraphCollection::new(self.document.clone_ref(py), self.path.clone()),
        )
    }
    fn add_paragraph(&self, py: Python<'_>, text: &str) -> PyResult<Py<PyParagraph>> {
        let (table, row, cell) = self.top_level(py, "Cell.add_paragraph")?;
        self.bump(py, false)?;
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
            document.revisions.capture(smallvec![
                PathSeg::Body(table),
                PathSeg::Row(row),
                PathSeg::Cell(cell),
                PathSeg::Para(paragraph)
            ])
        };
        Py::new(py, PyParagraph::new(self.document.clone_ref(py), path))
    }

    /// The tables nested directly in this cell, in source order.
    #[getter]
    fn tables<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let count = self.read(py, |cell| cell.tables().count())?;
        let tables = (0..count)
            .map(|index| self.nested_table(py, index))
            .collect::<PyResult<Vec<_>>>()?;
        PyList::new(py, tables)
    }

    /// Add a nested table after this cell's content, followed by the empty
    /// paragraph Word requires at the end of a cell.
    ///
    /// As in python-docx, the table is as wide as the cell, its columns
    /// equal, when the cell has a fixed width.
    fn add_table(&self, py: Python<'_>, rows: usize, cols: usize) -> PyResult<Py<PyTable>> {
        let (index, width) = self.read(py, |cell| (cell.tables().count(), cell.width()))?;
        // A cell without a fixed width is as wide as the grid columns it
        // covers.
        let width = match width {
            Some(width) => Some(width),
            None => {
                let (address, row, cell) = self.validate(py)?;
                let document = self.document.borrow(py);
                let table = address.get(&document.inner)?;
                let grid = table.grid_widths();
                table.row(row).and_then(|cells| {
                    let mut start = cells.grid_before().unwrap_or(0) as usize;
                    for index in 0..cell {
                        start += cells.cell(index)?.grid_span().unwrap_or(1).max(1) as usize;
                    }
                    let span = cells.cell(cell)?.grid_span().unwrap_or(1).max(1) as usize;
                    let covered = grid.get(start..start + span)?;
                    Some(rdocx::Length::twips(
                        covered.iter().map(|width| width.to_twips()).sum(),
                    ))
                })
            }
        };
        self.edit_table(py, |table, row, cell| {
            let mut cell = table.cell(row, cell).expect("the cell was checked");
            let mut nested = cell.add_table_checked(rows, cols)?;
            if let Some(width) = width.filter(|width| width.to_twips() > 0) {
                let twips = width.to_twips();
                let columns = cols as i32;
                let widths = (0..columns)
                    .map(|column| {
                        rdocx::Length::twips(twips / columns + i32::from(column < twips % columns))
                    })
                    .collect::<Vec<_>>();
                if widths.iter().all(|width| width.to_twips() > 0) {
                    nested.set_grid_widths(&widths)?;
                }
            }
            Ok(())
        })?;
        self.bump(py, false)?;
        self.nested_table(py, index)
    }

    /// Split this merged cell into one cell per grid column and row it
    /// covers, undoing its horizontal and vertical merges, and return how
    /// many cells the merge became.
    fn split(&self, py: Python<'_>) -> PyResult<usize> {
        let cells = self.edit_table(py, |table, row, cell| table.split_cell(row, cell))?;
        self.bump(py, true)?;
        Ok(cells)
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |cell| cell.width())?
            .map(|value| length_object(py, value))
            .transpose()
    }
    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.edit(py, |cell| {
            cell.set_width(rdocx::Length::emu(value));
            Ok(())
        })
    }
    #[getter]
    fn vertical_alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |cell| cell.vertical_alignment())?
            .map(|value| enum_object(py, "WD_CELL_VERTICAL_ALIGNMENT", vertical_to_int(value)))
            .transpose()
    }
    #[setter]
    fn set_vertical_alignment(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let value = vertical_from_int(value)?;
        self.edit(py, |cell| {
            cell.set_vertical_alignment(value);
            Ok(())
        })
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
        check_table_path(
            py,
            &self.document.borrow(py),
            &self.cell_path,
            "cell paragraph collection",
            "Re-fetch it with cell.paragraphs.",
        )?;
        match TableAddress::parse(&self.cell_path)? {
            (address, Some(row), Some(cell)) => {
                Ok((address.top_level("Cell.paragraphs")?, row, cell))
            }
            _ => Err(PyIndexError::new_err("cell path is incomplete")),
        }
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
