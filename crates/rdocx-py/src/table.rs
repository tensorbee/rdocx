use oxml_py_support::{ContentPath, PathSeg, revision_to_pyerr};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice, PyTuple};
use smallvec::smallvec;

use crate::document::PyDocument;
use crate::paragraph::PyParagraph;
use crate::{
    color_hex, color_object, enum_object, length_object, normalize_index, rdocx_to_pyerr,
    stale_to_pyerr,
};

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

pub(crate) fn checked_table_style(
    py: Python<'_>,
    candidate: &mut rdocx::Document,
    value: &str,
) -> PyResult<String> {
    candidate
        .ensure_builtin_table_style(value)
        .map_err(|e| rdocx_to_pyerr(py, e))?;
    let style_id = {
        let styles = candidate.styles();
        let exact_id = styles.iter().any(|style| {
            style.style_id() == value && style.style_type() == rdocx::StyleType::Table
        });
        if !exact_id {
            let matching = styles
                .iter()
                .filter(|style| {
                    style.style_type() == rdocx::StyleType::Table
                        && style
                            .name()
                            .is_some_and(|name| name.eq_ignore_ascii_case(value))
                })
                .count();
            if matching > 1 {
                return Err(PyValueError::new_err("table style name is ambiguous"));
            }
        }
        crate::paragraph::style_id_of_type(candidate, value, rdocx::StyleType::Table)?
    };
    Ok(style_id)
}

#[derive(Clone)]
pub(crate) enum TableLocation {
    Body(usize),
    Related(rdocx::ContentLocation),
}
impl TableLocation {
    fn new(index: usize, related: &Option<rdocx::ContentLocation>) -> Self {
        related.as_ref().map_or(Self::Body(index), |location| {
            Self::Related(location.clone())
        })
    }
    pub(crate) fn native(
        &self,
        document: &rdocx::Document,
    ) -> rdocx::Result<rdocx::ContentLocation> {
        crate::document::note_fx260_word(document.content_count().max(1));
        match self {
            Self::Related(location) => Self::refreshed(document, location),
            Self::Body(index) => {
                let slot = document.content_index_of_table(*index).ok_or_else(|| {
                    rdocx::Error::Other("table has no checked direct physical owner".into())
                })?;
                document
                    .story_item_snapshots()?
                    .into_iter()
                    .find(|snapshot| {
                        snapshot.location().story().kind() == rdocx::StoryKind::Body
                            && snapshot.location().item_kind() == rdocx::StoryItemKind::Table
                            && snapshot.is_direct_child()
                            && snapshot.direct_body_index() == Some(slot)
                    })
                    .map(|snapshot| snapshot.location().clone())
                    .ok_or_else(|| {
                        rdocx::Error::Other("table has no checked physical source owner".into())
                    })
            }
        }
    }
    fn refreshed(
        document: &rdocx::Document,
        location: &rdocx::ContentLocation,
    ) -> rdocx::Result<rdocx::ContentLocation> {
        crate::document::note_fx260_word(document.content_count().max(1));
        let owner = location.story();
        let story = document
            .stories()?
            .into_iter()
            .find(|story| {
                story.kind() == owner.kind()
                    && story.part_name() == owner.part_name()
                    && story.owner_index() == owner.owner_index()
            })
            .ok_or_else(|| rdocx::Error::Other("rich table owner disappeared".to_owned()))?;
        Ok(rdocx::ContentLocation::new(
            story,
            location.item_kind(),
            location.index_path().to_vec(),
        ))
    }
    fn read<T>(
        &self,
        document: &rdocx::Document,
        read: impl FnOnce(rdocx::TableRef<'_>) -> T,
    ) -> rdocx::Result<T> {
        crate::document::note_fx260_word(document.content_count().max(1));
        match self {
            Self::Body(index) => document
                .table(*index)
                .map(read)
                .ok_or_else(|| rdocx::Error::Other("table index out of range".to_owned())),
            Self::Related(location) => {
                document.read_story_table(&Self::refreshed(document, location)?, read)
            }
        }
    }
    fn edit<T>(
        &self,
        document: &mut rdocx::Document,
        edit: impl FnOnce(&mut rdocx::Table<'_>) -> rdocx::Result<T>,
    ) -> rdocx::Result<T> {
        match self {
            Self::Body(index) => edit(
                &mut document
                    .table_mut(*index)
                    .ok_or_else(|| rdocx::Error::Other("table index out of range".to_owned()))?,
            ),
            Self::Related(location) => {
                document.edit_story_table(&Self::refreshed(document, location)?, edit)
            }
        }
    }
}

/// The route of a table handle's physical owner in the inventory of
/// `prepared`, which holds the route `candidate` resolves.
fn prepared_table_location(
    py: Python<'_>,
    document: &PyDocument,
    prepared: &crate::document::PreparedWordRevision,
    candidate: &rdocx::Document,
    index: usize,
    related: &Option<rdocx::ContentLocation>,
) -> PyResult<rdocx::ContentLocation> {
    let native = || TableLocation::new(index, related).native(candidate);
    match document.indexed_table_location(index, related, Some(prepared)) {
        Some(indexed) => {
            #[cfg(debug_assertions)]
            crate::document::assert_indexed_route(&indexed, native, "table");
            Ok(indexed)
        }
        None => native().map_err(|error| rdocx_to_pyerr(py, error)),
    }
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
type BorderSnapshot = (String, Option<u32>, Option<Py<PyAny>>);

fn table_color_hex(value: &Bound<'_, PyAny>, parameter: &str) -> PyResult<String> {
    if value
        .extract::<String>()
        .is_ok_and(|value| value.eq_ignore_ascii_case("auto"))
    {
        Ok("auto".to_owned())
    } else {
        color_hex(value, parameter)
    }
}

fn border_snapshot(py: Python<'_>, border: rdocx::TableBorderRef<'_>) -> PyResult<BorderSnapshot> {
    Ok((
        border.style().to_owned(),
        border.size_eighths_pt(),
        border
            .color()
            .map(|value| color_object(py, value))
            .transpose()?,
    ))
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
        Py::new(
            py,
            PyTable::new(self.document.clone_ref(py), path).bind_current(py, None)?,
        )
    }
}

#[pymethods]
impl PyTableCollection {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.document.borrow(py).cached_body_table_count()
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
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyTable {
    pub(crate) fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self {
            document,
            path,
            related: None,
            owner: None,
            furniture_scope: None,
        }
    }

    pub(crate) fn for_story(
        document: Py<PyDocument>,
        path: ContentPath,
        location: rdocx::ContentLocation,
    ) -> Self {
        Self {
            document,
            path,
            related: Some(location),
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
        let index = table_index(&self.path)?;
        {
            let document = self.document.borrow(py);
            match document
                .indexed_table_location(index, &self.related, None)
                .map(Ok)
                .unwrap_or_else(|| TableLocation::new(index, &self.related).native(&document.inner))
            {
                Ok(location) => {
                    let (path, owner) = document.capture_owner(
                        &crate::document::WordOwnerLocation::Table(location),
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

    /// Bind the table a prepared body splice inserts, before the native edit
    /// places it.
    pub(crate) fn bind_prepared_splice(
        mut self,
        document: &PyDocument,
        prepared: &crate::document::PreparedBodySplice,
    ) -> PyResult<Self> {
        self.furniture_scope = None;
        let (path, owner) =
            document.capture_prepared_splice_owner(prepared, self.path.segs.clone())?;
        self.path = path;
        self.owner = Some(owner);
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
        let index = table_index(&self.path)?;
        let location =
            prepared_table_location(py, document, prepared, candidate, index, &self.related)?;
        let (path, owner) = document.capture_prepared_owner(
            prepared,
            &crate::document::WordOwnerLocation::Table(location),
            self.path.segs.clone(),
            crate::document::WordView::Owner,
            self.furniture_scope.as_ref(),
        )?;
        self.path = path;
        self.owner = Some(owner);
        Ok(self)
    }

    fn current_related(&self, py: Python<'_>) -> PyResult<Option<rdocx::ContentLocation>> {
        if let (Some(owner), Some(_)) = (&self.owner, &self.related) {
            self.path
                .validate_scopes(
                    &self.document.borrow(py).revisions,
                    "table",
                    "Re-fetch it from its current physical owner.",
                )
                .map_err(|error| stale_to_pyerr(py, error))?;
            let document = self.document.borrow(py);
            return document
                .indexed_table_owner_location(owner)
                .map(Ok)
                .unwrap_or_else(|| document.current_table_owner(owner))
                .map(Some);
        }
        Ok(self.related.clone())
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::TableRef<'_>) -> T) -> PyResult<T> {
        let index = self.validate(py)?;
        TableLocation::new(index, &self.current_related(py)?)
            .read(&self.document.borrow(py).inner, read)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }
    pub(crate) fn direct_body_index(&self, py: Python<'_>) -> PyResult<usize> {
        let index = self.validate(py)?;
        let document = self.document.borrow(py);
        if let Some(location) = &self.current_related(py)? {
            let location = TableLocation::refreshed(&document.inner, location)
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            return document
                .inner
                .story_item_snapshots()
                .map_err(|error| rdocx_to_pyerr(py, error))?
                .into_iter()
                .find(|item| {
                    item.location() == &location
                        && item.is_direct_child()
                        && location.story().kind() == rdocx::StoryKind::Body
                })
                .and_then(|item| item.direct_body_index())
                .ok_or_else(|| PyValueError::new_err("content handle is not a direct body child"));
        }
        document
            .inner
            .content_index_of_table(index)
            .ok_or_else(|| PyValueError::new_err("content handle is not a direct body child"))
    }
    pub(crate) fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
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
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let result = TableLocation::new(index, &current_related.clone())
            .edit(&mut candidate, edit)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.adopt_property_edit(py, candidate, "Table property edit")?;
        Ok(result)
    }

    /// Resolve possibly negative row and cell indexes against this table.
    fn cell_coordinates(&self, py: Python<'_>, row: isize, col: isize) -> PyResult<(usize, usize)> {
        self.read(py, |table| {
            let row = normalize_index(row, table.row_count(), "row")?;
            let col = normalize_index(
                col,
                table.row(row).map(|row| row.cell_count()).unwrap_or(0),
                "cell",
            )?;
            Ok((row, col))
        })?
    }
}

#[pymethods]
impl PyTable {
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
            .ok_or_else(|| PyValueError::new_err("table XML owner route is unsupported"))?;
        let current = || -> PyResult<Vec<u8>> {
            crate::document::note_word_whole_source_read();
            let location = document.current_table_owner(owner)?;
            document
                .inner
                .table_element_xml(&location)
                .map_err(|error| rdocx_to_pyerr(py, error))
        };
        // A direct body table serializes alone, without the `w14`
        // declaration that only the rest of the body can bring.
        let bytes = match document
            .direct_body_table_slot(owner)
            .and_then(|slot| document.inner.direct_body_table_xml(slot).ok().flatten())
        {
            Some(bytes) => {
                #[cfg(debug_assertions)]
                crate::document::assert_direct_read(&bytes, current, "table");
                bytes
            }
            None => current()?,
        };
        String::from_utf8(bytes).map_err(|error| PyValueError::new_err(error.to_string()))
    }

    fn replace_xml(&self, py: Python<'_>, xml: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let owner = self
            .owner
            .as_ref()
            .ok_or_else(|| PyValueError::new_err("table XML owner route is unsupported"))?;
        let location = document.current_table_owner(owner)?;
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .replace_table_xml(&location, xml)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.publish_owner_candidate(
            py,
            candidate,
            "Table.replace_xml",
            crate::document::WordPublication::ElementXml(
                crate::document::WordOwnerLocation::Table(location),
            ),
        )
    }

    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Py<PyRowCollection>> {
        self.validate(py)?;
        let path = match &self.owner {
            Some(owner) => {
                let document = self.document.borrow(py);
                document.capture_current_view(
                    owner,
                    self.path.segs.clone(),
                    crate::document::WordView::Grid,
                    self.furniture_scope.as_ref(),
                )?
            }
            None => self.path.clone(),
        };
        Py::new(
            py,
            PyRowCollection::new(
                self.document.clone_ref(py),
                path,
                self.current_related(py)?,
                self.owner.clone(),
                self.furniture_scope.clone(),
            ),
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
        Py::new(
            py,
            PyCell::new(self.document.clone_ref(py), path, self.current_related(py)?)
                .bind_current(py, self.furniture_scope.clone())?,
        )
    }

    #[pyo3(signature = (width))]
    fn add_column(&self, py: Python<'_>, width: &Bound<'_, PyAny>) -> PyResult<()> {
        self.validate(py)?;
        let _ = width;
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "use table.insert_column(index, before=...) and table.set_column_width(column, width); construct an empty grid with Document.add_table(rows, cols)",
        ))
    }

    #[getter]
    fn columns(&self, py: Python<'_>) -> PyResult<()> {
        self.validate(py)?;
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "use table.grid_widths and table.set_column_width(column, width), or cell.width",
        ))
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.read(py, |table| table.style_id().map(str::to_owned))
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let index = self.validate(py)?;
        let related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let style_id = checked_table_style(py, &mut candidate, value)?;
        TableLocation::new(index, &related)
            .edit(&mut candidate, |table| {
                table.set_style(&style_id);
                Ok(())
            })
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        document.inner = candidate;
        Ok(())
    }

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
        self.read(py, |table| {
            table
                .look()
                .unwrap_or(rdocx::TableLook {
                    horizontal_banding: true,
                    vertical_banding: true,
                    ..rdocx::TableLook::default()
                })
                .first_row
        })
    }

    #[setter]
    fn set_first_row(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut look = self.read(py, |table| {
            table.look().unwrap_or(rdocx::TableLook {
                horizontal_banding: true,
                vertical_banding: true,
                ..rdocx::TableLook::default()
            })
        })?;
        look.first_row = value;
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }

    #[getter]
    fn last_row(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |table| {
            table
                .look()
                .unwrap_or(rdocx::TableLook {
                    horizontal_banding: true,
                    vertical_banding: true,
                    ..rdocx::TableLook::default()
                })
                .last_row
        })
    }

    #[setter]
    fn set_last_row(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut look = self.read(py, |table| {
            table.look().unwrap_or(rdocx::TableLook {
                horizontal_banding: true,
                vertical_banding: true,
                ..rdocx::TableLook::default()
            })
        })?;
        look.last_row = value;
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }

    #[getter]
    fn first_column(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |table| {
            table
                .look()
                .unwrap_or(rdocx::TableLook {
                    horizontal_banding: true,
                    vertical_banding: true,
                    ..rdocx::TableLook::default()
                })
                .first_column
        })
    }

    #[setter]
    fn set_first_column(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut look = self.read(py, |table| {
            table.look().unwrap_or(rdocx::TableLook {
                horizontal_banding: true,
                vertical_banding: true,
                ..rdocx::TableLook::default()
            })
        })?;
        look.first_column = value;
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }

    #[getter]
    fn last_column(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |table| {
            table
                .look()
                .unwrap_or(rdocx::TableLook {
                    horizontal_banding: true,
                    vertical_banding: true,
                    ..rdocx::TableLook::default()
                })
                .last_column
        })
    }

    #[setter]
    fn set_last_column(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut look = self.read(py, |table| {
            table.look().unwrap_or(rdocx::TableLook {
                horizontal_banding: true,
                vertical_banding: true,
                ..rdocx::TableLook::default()
            })
        })?;
        look.last_column = value;
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }

    #[getter]
    fn banded_rows(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |table| {
            table
                .look()
                .unwrap_or(rdocx::TableLook {
                    horizontal_banding: true,
                    vertical_banding: true,
                    ..rdocx::TableLook::default()
                })
                .horizontal_banding
        })
    }

    #[setter]
    fn set_banded_rows(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut look = self.read(py, |table| {
            table.look().unwrap_or(rdocx::TableLook {
                horizontal_banding: true,
                vertical_banding: true,
                ..rdocx::TableLook::default()
            })
        })?;
        look.horizontal_banding = value;
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }

    #[getter]
    fn banded_columns(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |table| {
            table
                .look()
                .unwrap_or(rdocx::TableLook {
                    horizontal_banding: true,
                    vertical_banding: true,
                    ..rdocx::TableLook::default()
                })
                .vertical_banding
        })
    }

    #[setter]
    fn set_banded_columns(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let mut look = self.read(py, |table| {
            table.look().unwrap_or(rdocx::TableLook {
                horizontal_banding: true,
                vertical_banding: true,
                ..rdocx::TableLook::default()
            })
        })?;
        look.vertical_banding = value;
        self.edit(py, |table| {
            table.set_look(look);
            Ok(())
        })
    }

    #[getter]
    fn indent(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |table| table.indent())?
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_indent(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.edit(py, |table| {
            table.set_indent_value(value.map(rdocx::Length::emu))
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

    #[pyo3(signature = (style, *, size, color))]
    fn set_borders(
        &self,
        py: Python<'_>,
        style: &str,
        size: u32,
        color: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let color = table_color_hex(color, "Table.set_borders(color)")?;
        let style = border_style_from_name(style)?;
        self.edit(py, |table| {
            table.set_all_borders_checked(style, size, &color)
        })
    }

    #[pyo3(signature = (edge, style, *, size, color))]
    fn set_border(
        &self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        size: u32,
        color: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let color = table_color_hex(color, "Table.set_border(color)")?;
        let edge = table_border_edge(edge)?;
        let style = border_style_from_name(style)?;
        self.edit(py, |table| {
            table.set_border_checked(edge, style, size, &color)
        })
    }

    fn border(&self, py: Python<'_>, edge: &str) -> PyResult<Option<BorderSnapshot>> {
        let edge = table_border_edge(edge)?;
        self.read(py, |table| {
            table.border(edge).map(|border| border_snapshot(py, border))
        })?
        .transpose()
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
        let columns = self.read(py, |table| table.grid_widths().len())?;
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
        let table_index = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let changed = {
            TableLocation::new(table_index, &current_related.clone())
                .edit(&mut candidate, |table| {
                    let before = table.row(row).map(|row| row.cell_count());
                    table.set_cell_grid_span_checked(row, col, span)?;
                    Ok(table.row(row).map(|row| row.cell_count()) != before)
                })
                .map_err(|error| rdocx_to_pyerr(py, error))?
        };
        if changed {
            let prepared = document.prepare_word_revision(
                py,
                &candidate,
                "Table.set_cell_grid_span",
                crate::document::WordPublication::Conservative,
            )?;
            document.commit_word_revision(prepared)?;
            document.inner = candidate;
        } else {
            document.adopt_property_edit(py, candidate, "Table.set_cell_grid_span")?;
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

    fn add_row(&self, py: Python<'_>) -> PyResult<Py<PyRow>> {
        let index = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let inserted = match &current_related.clone() {
            None => candidate.add_empty_table_row(index),
            Some(location) => {
                let location = TableLocation::refreshed(&candidate, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
                candidate.edit_story_table(&location, |table| {
                    let index = table.row_count();
                    table.add_row()?;
                    Ok(index)
                })
            }
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Table.add_row",
            crate::document::WordPublication::TableGrid(
                TableLocation::new(index, &current_related)
                    .native(&document.inner)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
            ),
        )?;
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Row(inserted));
        let path = document
            .revisions
            .capture_prepared(&prepared.revision, segments)
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyRow::new(self.document.clone_ref(py), path, current_related.clone()).bind_prepared(
                py,
                &document,
                &prepared,
                &candidate,
                self.furniture_scope.clone(),
            )?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    #[pyo3(signature = (index, *, before = true))]
    fn insert_column(&self, py: Python<'_>, index: usize, before: bool) -> PyResult<()> {
        let table_index = self.validate(py)?;
        if index >= self.read(py, |table| table.column_count())? {
            return Err(PyIndexError::new_err(
                "column insertion index is out of range",
            ));
        }
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        match &current_related.clone() {
            None => candidate.insert_table_column(table_index, index, before),
            Some(location) => {
                let location = TableLocation::refreshed(&candidate, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
                candidate
                    .edit_story_table(&location, |table| table.insert_column(index, before, None))
            }
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Table.insert_column",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }

    fn delete_column(&self, py: Python<'_>, index: usize) -> PyResult<()> {
        let table_index = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        match &current_related.clone() {
            None => candidate.delete_table_column(table_index, index),
            Some(location) => candidate.delete_story_table_column(
                &TableLocation::refreshed(&candidate, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
                index,
            ),
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Table.delete_column",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }

    #[pyo3(signature = (index, at = None))]
    fn clone_row(&self, py: Python<'_>, index: isize, at: Option<usize>) -> PyResult<Py<PyRow>> {
        let table_index = self.validate(py)?;
        let result = {
            let current_related = self.current_related(py)?;
            let mut document = self.document.borrow_mut(py);
            let mut candidate = document.inner.clone_for_staging();
            let row_count = TableLocation::new(table_index, &current_related.clone())
                .read(&candidate, |table| table.row_count())
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            let source = normalize_index(index, row_count, "row")?;
            let insert_at = at.unwrap_or(source + 1);
            if insert_at > row_count {
                return Err(PyIndexError::new_err("row insertion index out of range"));
            }
            let inserted = match &current_related.clone() {
                Some(location) => {
                    let location = TableLocation::refreshed(&candidate, location)
                        .map_err(|error| rdocx_to_pyerr(py, error))?;
                    candidate.clone_story_table_row(&location, source, insert_at)
                }
                None => candidate.clone_table_row(table_index, source, insert_at),
            }
            .map_err(|error| rdocx_to_pyerr(py, error))?;
            let prepared = document.prepare_word_revision(
                py,
                &candidate,
                "Table.clone_row",
                crate::document::WordPublication::Conservative,
            )?;
            let mut segments = self.path.segs.clone();
            segments.push(PathSeg::Row(inserted));
            let path = document
                .revisions
                .capture_prepared(&prepared.revision, segments)
                .map_err(oxml_py_support::revision_to_pyerr)?;
            let result = Py::new(
                py,
                PyRow::new(self.document.clone_ref(py), path, current_related.clone())
                    .bind_prepared(
                        py,
                        &document,
                        &prepared,
                        &candidate,
                        self.furniture_scope.clone(),
                    )?,
            )?;
            document.commit_word_revision(prepared)?;
            document.inner = candidate;
            result
        };
        Ok(result)
    }

    fn remove_row(&self, py: Python<'_>, index: isize) -> PyResult<()> {
        let table_index = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let row_count = TableLocation::new(table_index, &current_related.clone())
            .read(&candidate, |table| table.row_count())
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let row_index = normalize_index(index, row_count, "row")?;
        match &current_related.clone() {
            Some(location) => {
                let location = TableLocation::refreshed(&candidate, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
                candidate.remove_story_table_row(&location, row_index)
            }
            None => candidate.remove_table_row(table_index, row_index),
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Table.remove_row",
            crate::document::WordPublication::TableGrid(
                TableLocation::new(table_index, &current_related)
                    .native(&document.inner)
                    .map_err(|error| rdocx_to_pyerr(py, error))?,
            ),
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }
}

#[pyclass(name = "RowCollection")]
pub struct PyRowCollection {
    document: Py<PyDocument>,
    table_path: ContentPath,
    related: Option<rdocx::ContentLocation>,
    table_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyRowCollection {
    fn new(
        document: Py<PyDocument>,
        table_path: ContentPath,
        related: Option<rdocx::ContentLocation>,
        table_owner: Option<oxml_py_support::RevisionState>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> Self {
        Self {
            document,
            table_path,
            related,
            table_owner,
            furniture_scope,
        }
    }
    fn current_related(&self, py: Python<'_>) -> PyResult<Option<rdocx::ContentLocation>> {
        self.table_path
            .validate_scopes(
                &self.document.borrow(py).revisions,
                "table grid view",
                "Re-fetch it from its current physical table.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        if let (Some(owner), Some(_)) = (&self.table_owner, &self.related) {
            let document = self.document.borrow(py);
            return document
                .indexed_table_owner_location(owner)
                .map(Ok)
                .unwrap_or_else(|| document.current_table_owner(owner))
                .map(Some);
        }
        Ok(self.related.clone())
    }

    fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let document = self.document.borrow(py);
        self.table_path
            .validate_scopes(
                &document.revisions,
                "row collection",
                "Re-fetch it with table.rows.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        table_index(&self.table_path)
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let index = self.validate(py)?;
        if let Some(owner) = &self.table_owner
            && let Some(rows) = self.document.borrow(py).cached_table_rows(owner)
        {
            return Ok(rows.len());
        }
        TableLocation::new(index, &self.current_related(py)?)
            .read(&self.document.borrow(py).inner, |table| table.row_count())
            .map_err(|error| rdocx_to_pyerr(py, error))
    }
    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyRow>> {
        self.validate(py)?;
        let mut segments = self.table_path.segs.clone();
        segments.push(PathSeg::Row(index));
        let mut path = self.table_path.clone();
        path.segs = segments;
        Py::new(
            py,
            PyRow::new(self.document.clone_ref(py), path, self.current_related(py)?)
                .bind_current(py, self.furniture_scope.clone())?,
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
        self.validate(py)?;
        Py::new(
            py,
            PyRowIterator {
                document: self.document.clone_ref(py),
                table_path: self.table_path.clone(),
                related: self.current_related(py)?,
                table_owner: self.table_owner.clone(),
                furniture_scope: self.furniture_scope.clone(),
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
    related: Option<rdocx::ContentLocation>,
    table_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

#[pymethods]
impl PyRowIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyRow>>> {
        let collection = PyRowCollection::new(
            self.document.clone_ref(py),
            self.table_path.clone(),
            self.related.clone(),
            self.table_owner.clone(),
            self.furniture_scope.clone(),
        );
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
    related: Option<rdocx::ContentLocation>,
    table_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyRow {
    fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        related: Option<rdocx::ContentLocation>,
    ) -> Self {
        Self {
            document,
            path,
            related,
            table_owner: None,
            furniture_scope: None,
        }
    }
    fn bind_current(
        mut self,
        py: Python<'_>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> PyResult<Self> {
        self.furniture_scope = furniture_scope;
        let index = table_index(&self.path)?;
        {
            let document = self.document.borrow(py);
            match document
                .indexed_table_location(index, &self.related, None)
                .map(Ok)
                .unwrap_or_else(|| TableLocation::new(index, &self.related).native(&document.inner))
            {
                Ok(location) => {
                    let (path, owner) = document.capture_owner(
                        &crate::document::WordOwnerLocation::Table(location),
                        self.path.segs.clone(),
                        crate::document::WordView::Grid,
                        self.furniture_scope.as_ref(),
                    )?;
                    self.path = path;
                    self.table_owner = Some(owner);
                }
                Err(_) => {
                    self.path = document.capture_snapshot(self.path.segs.clone())?;
                }
            }
        }
        Ok(self)
    }

    fn bind_prepared(
        mut self,
        py: Python<'_>,
        document: &PyDocument,
        prepared: &crate::document::PreparedWordRevision,
        candidate: &rdocx::Document,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> PyResult<Self> {
        self.furniture_scope =
            document.prepared_furniture_scope(prepared, furniture_scope.as_ref())?;
        let location = prepared_table_location(
            py,
            document,
            prepared,
            candidate,
            table_index(&self.path)?,
            &self.related,
        )?;
        let (path, owner) = document.capture_prepared_owner(
            prepared,
            &crate::document::WordOwnerLocation::Table(location),
            self.path.segs.clone(),
            crate::document::WordView::Grid,
            self.furniture_scope.as_ref(),
        )?;
        self.path = path;
        self.table_owner = Some(owner);
        Ok(self)
    }

    fn current_related(&self, py: Python<'_>) -> PyResult<Option<rdocx::ContentLocation>> {
        self.path
            .validate_scopes(
                &self.document.borrow(py).revisions,
                "table grid view",
                "Re-fetch it from its current physical table.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        if let (Some(owner), Some(_)) = (&self.table_owner, &self.related) {
            let document = self.document.borrow(py);
            return document
                .indexed_table_owner_location(owner)
                .map(Ok)
                .unwrap_or_else(|| document.current_table_owner(owner))
                .map(Some);
        }
        Ok(self.related.clone())
    }

    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
                "row",
                "Re-fetch it with table.rows[i].",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok((table_index(&self.path)?, row_index(&self.path)?))
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rdocx::RowRef<'_>) -> T) -> PyResult<T> {
        let (table, row) = self.validate(py)?;
        TableLocation::new(table, &self.current_related(py)?)
            .read(&self.document.borrow(py).inner, |table| {
                let row = table
                    .row(row)
                    .ok_or_else(|| PyIndexError::new_err("row index out of range"))?;
                Ok(read(row))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
    }

    /// Apply one checked native row edit that moves no content, so live
    /// handles stay valid.
    fn edit(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Row<'_>) -> rdocx::Result<()>,
    ) -> PyResult<()> {
        let (table, row) = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        TableLocation::new(table, &current_related.clone())
            .edit(&mut candidate, |table| {
                let mut row = table
                    .row(row)
                    .ok_or_else(|| rdocx::Error::Other("row index out of range".to_owned()))?;
                edit(&mut row)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.adopt_property_edit(py, candidate, "Row property edit")?;
        Ok(())
    }
}

#[pymethods]
impl PyRow {
    #[getter]
    fn cells(&self, py: Python<'_>) -> PyResult<Py<PyCellCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyCellCollection::new(
                self.document.clone_ref(py),
                self.path.clone(),
                self.current_related(py)?,
                self.table_owner.clone(),
                self.furniture_scope.clone(),
            ),
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
    related: Option<rdocx::ContentLocation>,
    table_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyCellCollection {
    fn new(
        document: Py<PyDocument>,
        row_path: ContentPath,
        related: Option<rdocx::ContentLocation>,
        table_owner: Option<oxml_py_support::RevisionState>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> Self {
        Self {
            document,
            row_path,
            related,
            table_owner,
            furniture_scope,
        }
    }
    fn current_related(&self, py: Python<'_>) -> PyResult<Option<rdocx::ContentLocation>> {
        self.row_path
            .validate_scopes(
                &self.document.borrow(py).revisions,
                "table grid view",
                "Re-fetch it from its current physical table.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        if let (Some(owner), Some(_)) = (&self.table_owner, &self.related) {
            let document = self.document.borrow(py);
            return document
                .indexed_table_owner_location(owner)
                .map(Ok)
                .unwrap_or_else(|| document.current_table_owner(owner))
                .map(Some);
        }
        Ok(self.related.clone())
    }

    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize)> {
        let document = self.document.borrow(py);
        self.row_path
            .validate_scopes(
                &document.revisions,
                "cell collection",
                "Re-fetch it with row.cells.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        Ok((table_index(&self.row_path)?, row_index(&self.row_path)?))
    }
    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let (table, row) = self.validate(py)?;
        if let Some(owner) = &self.table_owner
            && let Some(rows) = self.document.borrow(py).cached_table_rows(owner)
        {
            return rows
                .get(row)
                .map(Vec::len)
                .ok_or_else(|| PyIndexError::new_err("row index out of range"));
        }
        TableLocation::new(table, &self.current_related(py)?)
            .read(&self.document.borrow(py).inner, |table| {
                table.row(row).map(|row| row.cell_count())
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("row index out of range"))
    }
    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyCell>> {
        self.validate(py)?;
        let mut segments = self.row_path.segs.clone();
        segments.push(PathSeg::Cell(index));
        let mut path = self.row_path.clone();
        path.segs = segments;
        Py::new(
            py,
            PyCell::new(self.document.clone_ref(py), path, self.current_related(py)?)
                .bind_current(py, self.furniture_scope.clone())?,
        )
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
                related: self.current_related(py)?,
                table_owner: self.table_owner.clone(),
                furniture_scope: self.furniture_scope.clone(),
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
    related: Option<rdocx::ContentLocation>,
    table_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

#[pymethods]
impl PyCellIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyCell>>> {
        let collection = PyCellCollection::new(
            self.document.clone_ref(py),
            self.row_path.clone(),
            self.related.clone(),
            self.table_owner.clone(),
            self.furniture_scope.clone(),
        );
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
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyCell {
    fn physical_story(
        &self,
        py: Python<'_>,
        document: &rdocx::Document,
        related: &Option<rdocx::ContentLocation>,
        table: usize,
        row: usize,
        cell: usize,
    ) -> PyResult<rdocx::StoryId> {
        let location = TableLocation::new(table, related)
            .native(document)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document
            .story_table_cell(&location, row, cell)
            .map_err(|error| rdocx_to_pyerr(py, error))
    }

    fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        related: Option<rdocx::ContentLocation>,
    ) -> Self {
        Self {
            document,
            path,
            related,
            owner: None,
            furniture_scope: None,
        }
    }
    fn bind_current(
        mut self,
        py: Python<'_>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> PyResult<Self> {
        self.furniture_scope = furniture_scope;
        let table = table_index(&self.path)?;
        let row = row_index(&self.path)?;
        let cell = cell_index(&self.path)?;
        {
            let document = self.document.borrow(py);
            match document
                .indexed_table_location(table, &self.related, None)
                .map(Ok)
                .unwrap_or_else(|| TableLocation::new(table, &self.related).native(&document.inner))
            {
                Ok(table) => {
                    let location = match document.indexed_cell_location(&table, row, cell) {
                        Some(location) => location,
                        None => {
                            // An inventory's body story has no fingerprint.
                            let native = TableLocation::Related(table.clone())
                                .native(&document.inner)
                                .map_err(|error| rdocx_to_pyerr(py, error))?;
                            let story = document
                                .inner
                                .story_table_cell(&native, row, cell)
                                .map_err(|error| rdocx_to_pyerr(py, error))?;
                            crate::document::WordOwnerLocation::Cell {
                                table,
                                row,
                                cell,
                                story,
                            }
                        }
                    };
                    let (path, owner) = document.capture_owner(
                        &location,
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

    fn current_related(&self, py: Python<'_>) -> PyResult<Option<rdocx::ContentLocation>> {
        self.path
            .validate_scopes(
                &self.document.borrow(py).revisions,
                "cell",
                "Re-fetch it from its current physical owner.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        if let (Some(owner), Some(_)) = (&self.owner, &self.related) {
            let document = self.document.borrow(py);
            if let Some(table) = document.indexed_cell_parent_location(owner) {
                return Ok(Some(table));
            }
            return document
                .current_cell_owner(owner)
                .map(|(table, _, _)| Some(table));
        }
        Ok(self.related.clone())
    }

    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
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
        TableLocation::new(table, &self.current_related(py)?)
            .read(&self.document.borrow(py).inner, |table| {
                let cell = table
                    .cell(row, cell)
                    .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?;
                Ok(read(cell))
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
    }

    /// Apply one checked native cell edit that moves no content, so live
    /// handles stay valid.
    fn edit(
        &self,
        py: Python<'_>,
        edit: impl FnOnce(&mut rdocx::Cell<'_>) -> rdocx::Result<()>,
    ) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        TableLocation::new(table, &current_related.clone())
            .edit(&mut candidate, |table| {
                let mut cell = table
                    .cell(row, cell)
                    .ok_or_else(|| rdocx::Error::Other("cell index out of range".to_owned()))?;
                edit(&mut cell)
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.adopt_property_edit(py, candidate, "Cell property edit")?;
        Ok(())
    }
}

#[pymethods]
impl PyCell {
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
            .ok_or_else(|| PyValueError::new_err("cell XML owner route is unsupported"))?;
        let current = || -> PyResult<Vec<u8>> {
            crate::document::note_word_whole_source_read();
            let (location, row, cell) = document.current_cell_owner(owner)?;
            document
                .inner
                .cell_element_xml(&location, row, cell)
                .map_err(|error| rdocx_to_pyerr(py, error))
        };
        // A cell of a direct body table serializes with that table alone.
        let bytes = match document
            .direct_body_cell_slot(owner)
            .and_then(|(slot, row, cell)| {
                document
                    .inner
                    .direct_body_cell_xml(slot, row, cell)
                    .ok()
                    .flatten()
            }) {
            Some(bytes) => {
                #[cfg(debug_assertions)]
                crate::document::assert_direct_read(&bytes, current, "cell");
                bytes
            }
            None => current()?,
        };
        String::from_utf8(bytes).map_err(|error| PyValueError::new_err(error.to_string()))
    }

    fn replace_xml(&self, py: Python<'_>, xml: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        let owner = self
            .owner
            .as_ref()
            .ok_or_else(|| PyValueError::new_err("cell XML owner route is unsupported"))?;
        let (location, row, cell) = document.current_cell_owner(owner)?;
        let story = document
            .inner
            .story_table_cell(&location, row, cell)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let mut candidate = document.inner.clone_for_staging();
        candidate
            .replace_cell_xml(&location, row, cell, xml)
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        document.publish_owner_candidate(
            py,
            candidate,
            "Cell.replace_xml",
            crate::document::WordPublication::ElementXml(
                crate::document::WordOwnerLocation::Cell {
                    table: location,
                    row,
                    cell,
                    story,
                },
            ),
        )
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
        let (table, row, cell) = self.validate(py)?;
        let (data, filename) = crate::document::picture_input(data, filename)?;
        let width = crate::document::checked_section_length(width, "picture width")?
            .map(rdocx::Length::emu);
        let height = crate::document::checked_section_length(height, "picture height")?
            .map(rdocx::Length::emu);
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let story = self.physical_story(py, &document.inner, &current_related, table, row, cell)?;
        let mut candidate = document.inner.clone_for_staging();
        let location = candidate
            .insert_picture_to_story_with_options(
                &story,
                None,
                &data,
                &filename,
                width,
                height,
                options.map(|o| o.inner.clone()),
            )
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let picture = candidate
            .pictures()
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .find(|p| p.location == location)
            .ok_or_else(|| PyIndexError::new_err("inserted cell picture was not inventoried"))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Cell.add_picture",
            crate::document::WordPublication::Conservative,
        )?;
        let result = Py::new(
            py,
            crate::document::PyPicture {
                inner: picture,
                binding: document
                    .revisions
                    .capture_prepared(&prepared.revision, smallvec![])
                    .map_err(revision_to_pyerr)?,
            },
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    /// Replace literal text in this cell and its supported nested descendants.
    #[pyo3(signature = (old, new, *, expect = None))]
    fn replace_text(
        &self,
        py: Python<'_>,
        old: &str,
        new: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        let (table, row, cell) = self.validate(py)?;
        let related = match &self.current_related(py)? {
            Some(location) => {
                let document = self.document.borrow(py);
                let location = TableLocation::refreshed(&document.inner, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
                Some(
                    document
                        .inner
                        .story_table_cell(&location, row, cell)
                        .map_err(|error| rdocx_to_pyerr(py, error))?,
                )
            }
            None => None,
        };
        self.document
            .borrow_mut(py)
            .scoped_replacement(py, "Cell.replace_text", |document| match related.as_ref() {
                Some(story) => document.try_replace_text_in_story_cell(story, old, new, expect),
                None => {
                    document.try_replace_text_in_cell((table, row, cell), None, old, new, expect)
                }
            })
    }

    fn merge(
        slf: &Bound<'_, Self>,
        py: Python<'_>,
        other: &Bound<'_, Self>,
    ) -> PyResult<Py<PyCell>> {
        let this = slf.borrow();
        let other = other.borrow();
        if this.document.as_ptr() != other.document.as_ptr() {
            return Err(PyValueError::new_err(
                "merge cell belongs to a different document",
            ));
        }
        let (table, row, cell) = this.validate(py)?;
        let (other_table, other_row, other_cell) = other.validate(py)?;
        let related = this.current_related(py)?;
        let other_related = other.current_related(py)?;
        let mut document = this.document.borrow_mut(py);
        let location = TableLocation::new(table, &related)
            .native(&document.inner)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let other_location = TableLocation::new(other_table, &other_related)
            .native(&document.inner)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        if location != other_location {
            return Err(PyValueError::new_err(
                "merge cells belong to different physical tables",
            ));
        }
        if (row, cell) == (other_row, other_cell) {
            return Ok(slf.clone().unbind());
        }
        let mut candidate = document.inner.clone_for_staging();
        let merged = candidate
            .edit_story_table(&location, |table| {
                table.merge_cells((row, cell), (other_row, other_cell))
            })
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let candidate_location = TableLocation::new(table, &related)
            .native(&candidate)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let story = candidate
            .story_table_cell(&candidate_location, merged.0, merged.1)
            .map_err(|e| rdocx_to_pyerr(py, e))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Cell.merge",
            crate::document::WordPublication::TableGrid(location),
        )?;
        let segments = smallvec![
            PathSeg::Body(table),
            PathSeg::Row(merged.0),
            PathSeg::Cell(merged.1)
        ];
        let (path, owner) = document.capture_prepared_owner(
            &prepared,
            &crate::document::WordOwnerLocation::Cell {
                table: candidate_location,
                row: merged.0,
                cell: merged.1,
                story,
            },
            segments,
            crate::document::WordView::Owner,
            this.furniture_scope.as_ref(),
        )?;
        let mut result = PyCell::new(this.document.clone_ref(py), path, related);
        result.owner = Some(owner);
        result.furniture_scope = this.furniture_scope.clone();
        let result = Py::new(py, result)?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }

    fn split(&self, py: Python<'_>) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let noop = self.read(py, |cell| {
            cell.grid_span().unwrap_or(1) == 1
                && cell.v_merge().is_none()
                && !cell.has_horizontal_merge()
        })?;
        if noop {
            return Ok(());
        }
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        match &current_related.clone() {
            Some(location) => {
                let location = TableLocation::refreshed(&candidate, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
                candidate.edit_story_table(&location, |table| {
                    table.split_cell(row, cell)?;
                    Ok(())
                })
            }
            None => {
                let logical_column = candidate
                    .table(table)
                    .and_then(|table| {
                        let row = table.row(row)?;
                        (0..cell).try_fold(
                            row.grid_before().unwrap_or(0) as usize,
                            |column, index| {
                                column
                                    .checked_add(row.cell(index)?.grid_span().unwrap_or(1) as usize)
                            },
                        )
                    })
                    .ok_or_else(|| {
                        PyValueError::new_err("cell has no checked logical grid position")
                    })?;
                candidate
                    .split_table_cell(table, row, logical_column)
                    .map(|_| ())
            }
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Cell.split",
            crate::document::WordPublication::Conservative,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.read(py, |cell| cell.text())
    }
    #[getter]
    fn tables<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let (table, row, cell) = self.validate(py)?;
        let cached = self
            .owner
            .as_ref()
            .and_then(|owner| self.document.borrow(py).cached_cell_tables(owner));
        let locations = match cached {
            Some(locations) => locations,
            None => {
                let current_related = self.current_related(py)?;
                let document = self.document.borrow(py);
                let story =
                    self.physical_story(py, &document.inner, &current_related, table, row, cell)?;
                document
                    .inner
                    .story_items(&story)
                    .map_err(|error| rdocx_to_pyerr(py, error))?
                    .into_iter()
                    .filter(|item| item.kind() == rdocx::StoryItemKind::Table)
                    .map(|item| item.location().clone())
                    .collect()
            }
        };
        let mut tables = Vec::new();
        for location in locations {
            let path = self
                .document
                .borrow(py)
                .revisions
                .capture(smallvec![PathSeg::Body(location.index_path()[0])]);
            tables.push(Py::new(
                py,
                PyTable::for_story(self.document.clone_ref(py), path, location)
                    .bind_current(py, self.furniture_scope.clone())?,
            )?);
        }
        PyTuple::new(py, tables)
    }

    fn add_table(&self, py: Python<'_>, rows: usize, cols: usize) -> PyResult<Py<PyTable>> {
        let (table, row, cell) = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        let story = self.physical_story(py, &candidate, &current_related, table, row, cell)?;
        let before = candidate
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .filter(|item| item.kind() == rdocx::StoryItemKind::Table)
            .count();
        TableLocation::new(table, &current_related.clone())
            .edit(&mut candidate, |table| {
                table
                    .cell(row, cell)
                    .ok_or_else(|| rdocx::Error::Other("cell index out of range".to_owned()))?
                    .add_table_checked(rows, cols)?;
                Ok(())
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?;
        let story = self.physical_story(py, &candidate, &current_related, table, row, cell)?;
        let tables = candidate
            .story_items(&story)
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .into_iter()
            .filter(|item| item.kind() == rdocx::StoryItemKind::Table)
            .map(|item| item.location().clone())
            .collect::<Vec<_>>();
        if tables.len() != before + 1 {
            return Err(PyValueError::new_err(
                "nested table insertion did not retain physical owner correspondence",
            ));
        }
        let location = tables
            .last()
            .cloned()
            .ok_or_else(|| PyValueError::new_err("nested table was not published"))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Cell.add_table",
            crate::document::WordPublication::Conservative,
        )?;
        let path = document
            .revisions
            .capture_prepared(
                &prepared.revision,
                smallvec![PathSeg::Body(location.index_path()[0])],
            )
            .map_err(oxml_py_support::revision_to_pyerr)?;
        let result = Py::new(
            py,
            PyTable::for_story(self.document.clone_ref(py), path, location).bind_prepared(
                py,
                &document,
                &prepared,
                &candidate,
                self.furniture_scope.clone(),
            )?,
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(result)
    }
    #[setter]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let (table, row, cell) = self.validate(py)?;
        let current_related = self.current_related(py)?;
        let mut document = self.document.borrow_mut(py);
        let mut candidate = document.inner.clone_for_staging();
        match &current_related.clone() {
            Some(_) => {
                TableLocation::new(table, &current_related.clone()).edit(&mut candidate, |table| {
                    table
                        .cell(row, cell)
                        .ok_or_else(|| rdocx::Error::Other("cell index out of range".to_owned()))?
                        .try_set_text(value)
                })
            }
            None => candidate.try_set_cell_text(table, row, cell, value),
        }
        .map_err(|error| rdocx_to_pyerr(py, error))?;
        let prepared = document.prepare_word_revision(
            py,
            &candidate,
            "Cell.text",
            TableLocation::new(table, &current_related)
                .native(&document.inner)
                .map(|location| crate::document::WordPublication::CellText(location, row, cell))
                .unwrap_or(crate::document::WordPublication::Conservative),
        )?;
        document.commit_word_revision(prepared)?;
        document.inner = candidate;
        Ok(())
    }
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> PyResult<Py<PyCellParagraphCollection>> {
        self.validate(py)?;
        let path = match &self.owner {
            Some(owner) => {
                let document = self.document.borrow(py);
                document.capture_current_view(
                    owner,
                    self.path.segs.clone(),
                    crate::document::WordView::Content,
                    self.furniture_scope.as_ref(),
                )?
            }
            None => self.path.clone(),
        };
        Py::new(
            py,
            PyCellParagraphCollection::new(
                self.document.clone_ref(py),
                path,
                self.current_related(py)?,
                self.owner.clone(),
                self.furniture_scope.clone(),
            ),
        )
    }
    fn add_paragraph(&self, py: Python<'_>, text: &str) -> PyResult<Py<PyParagraph>> {
        let (table, row, cell) = self.validate(py)?;
        let result = {
            let current_related = self.current_related(py)?;
            let mut document = self.document.borrow_mut(py);
            let mut candidate = document.inner.clone_for_staging();
            let paragraph = TableLocation::new(table, &current_related.clone())
                .edit(&mut candidate, |table| {
                    let mut cell = table
                        .cell(row, cell)
                        .ok_or_else(|| rdocx::Error::Other("cell index out of range".to_owned()))?;
                    let paragraph = cell.paragraph_count();
                    cell.add_paragraph(text);
                    Ok(paragraph)
                })
                .map_err(|error| rdocx_to_pyerr(py, error))?;
            let prepared = document.prepare_word_revision(
                py,
                &candidate,
                "Cell.add_paragraph",
                crate::document::WordPublication::Conservative,
            )?;
            let path = document
                .revisions
                .capture_prepared(
                    &prepared.revision,
                    smallvec![
                        PathSeg::Body(table),
                        PathSeg::Row(row),
                        PathSeg::Cell(cell),
                        PathSeg::Para(paragraph)
                    ],
                )
                .map_err(oxml_py_support::revision_to_pyerr)?;
            let paragraph = match &current_related.clone() {
                Some(location) => {
                    let location = TableLocation::refreshed(&candidate, location)
                        .map_err(|error| rdocx_to_pyerr(py, error))?;
                    let location = candidate
                        .story_table_cell_paragraph_location(&location, row, cell, paragraph)
                        .map_err(|error| rdocx_to_pyerr(py, error))?;
                    PyParagraph::for_story(self.document.clone_ref(py), path, location)
                        .bind_prepared(
                            py,
                            &document,
                            &prepared,
                            &candidate,
                            self.furniture_scope.clone(),
                        )?
                }
                None => PyParagraph::new(self.document.clone_ref(py), path).bind_prepared(
                    py,
                    &document,
                    &prepared,
                    &candidate,
                    self.furniture_scope.clone(),
                )?,
            };
            let result = Py::new(py, paragraph)?;
            document.commit_word_revision(prepared)?;
            document.inner = candidate;
            result
        };
        Ok(result)
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
    fn shading(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.read(py, |cell| cell.shading_fill().map(str::to_owned))?
            .as_deref()
            .map(|value| color_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_shading(&self, py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let value = table_color_hex(value, "Cell.shading")?;
        self.edit(py, |cell| cell.set_shading_checked(&value))
    }

    fn border(&self, py: Python<'_>, edge: &str) -> PyResult<Option<BorderSnapshot>> {
        let edge = cell_border_edge(edge)?;
        self.read(py, |cell| {
            cell.border(edge).map(|border| border_snapshot(py, border))
        })?
        .transpose()
    }

    #[pyo3(signature = (edge, style, *, size, color))]
    fn set_border(
        &self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        size: u32,
        color: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let color = table_color_hex(color, "Cell.set_border(color)")?;
        let edge = cell_border_edge(edge)?;
        let style = border_style_from_name(style)?;
        self.edit(py, |cell| {
            cell.set_border_checked(edge, style, size, &color)
        })
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
    related: Option<rdocx::ContentLocation>,
    cell_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

impl PyCellParagraphCollection {
    fn new(
        document: Py<PyDocument>,
        cell_path: ContentPath,
        related: Option<rdocx::ContentLocation>,
        cell_owner: Option<oxml_py_support::RevisionState>,
        furniture_scope: Option<oxml_py_support::RevisionState>,
    ) -> Self {
        Self {
            document,
            cell_path,
            related,
            cell_owner,
            furniture_scope,
        }
    }
    fn current_related(&self, py: Python<'_>) -> PyResult<Option<rdocx::ContentLocation>> {
        self.cell_path
            .validate_scopes(
                &self.document.borrow(py).revisions,
                "cell paragraph collection",
                "Re-fetch it from its current physical cell.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        if let (Some(owner), Some(_)) = (&self.cell_owner, &self.related) {
            let document = self.document.borrow(py);
            if let Some(table) = document.indexed_cell_parent_location(owner) {
                return Ok(Some(table));
            }
            return document
                .current_cell_owner(owner)
                .map(|(table, _, _)| Some(table));
        }
        Ok(self.related.clone())
    }

    fn validate(&self, py: Python<'_>) -> PyResult<(usize, usize, usize)> {
        let document = self.document.borrow(py);
        self.cell_path
            .validate_scopes(
                &document.revisions,
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
        if let Some(owner) = &self.cell_owner
            && let Some(count) = self.document.borrow(py).cached_cell_paragraph_count(owner)
        {
            return Ok(count);
        }
        TableLocation::new(table, &self.current_related(py)?)
            .read(&self.document.borrow(py).inner, |table| {
                table.cell(row, cell).map(|cell| cell.paragraph_count())
            })
            .map_err(|error| rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))
    }
    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyParagraph>> {
        self.validate(py)?;
        let mut segments = self.cell_path.segs.clone();
        segments.push(PathSeg::Para(index));
        let path = self.document.borrow(py).revisions.capture(segments);
        let cached = self
            .cell_owner
            .as_ref()
            .and_then(|owner| self.document.borrow(py).cached_cell_paragraph(owner, index));
        if let Some(location) = cached {
            let paragraph = if self.related.is_some() {
                PyParagraph::for_story(self.document.clone_ref(py), path, location)
            } else {
                PyParagraph::new(self.document.clone_ref(py), path)
            };
            return Py::new(
                py,
                paragraph.bind_current(py, self.furniture_scope.clone())?,
            );
        }
        let related = match &self.current_related(py)? {
            Some(location) => {
                let (_, row, cell) = self.validate(py)?;
                let document = self.document.borrow(py);
                let location = TableLocation::refreshed(&document.inner, location)
                    .map_err(|error| rdocx_to_pyerr(py, error))?;
                Some(
                    document
                        .inner
                        .story_table_cell_paragraph_location(&location, row, cell, index)
                        .map_err(|error| rdocx_to_pyerr(py, error))?,
                )
            }
            None => None,
        };
        let paragraph = match related {
            Some(location) => PyParagraph::for_story(self.document.clone_ref(py), path, location)
                .bind_current(py, self.furniture_scope.clone())?,
            None => PyParagraph::new(self.document.clone_ref(py), path)
                .bind_current(py, self.furniture_scope.clone())?,
        };
        Py::new(py, paragraph)
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
                related: self.current_related(py)?,
                cell_owner: self.cell_owner.clone(),
                furniture_scope: self.furniture_scope.clone(),
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
    related: Option<rdocx::ContentLocation>,
    cell_owner: Option<oxml_py_support::RevisionState>,
    furniture_scope: Option<oxml_py_support::RevisionState>,
}

#[pymethods]
impl PyCellParagraphIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyParagraph>>> {
        let collection = PyCellParagraphCollection::new(
            self.document.clone_ref(py),
            self.cell_path.clone(),
            self.related.clone(),
            self.cell_owner.clone(),
            self.furniture_scope.clone(),
        );
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}
