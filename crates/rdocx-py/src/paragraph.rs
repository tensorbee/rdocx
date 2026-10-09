use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyKeyError, PyRuntimeError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList, PySlice};
use smallvec::smallvec;

use crate::document::PyDocument;
use crate::run::{PyRun, PyRunCollection};
use crate::{normalize_index, stale_to_pyerr};

/// Where a hyperlink run points.
enum LinkTarget<'a> {
    /// A web address, written as an external relationship.
    Url(&'a str),
    /// A bookmark of the same document, written as `w:anchor`.
    Anchor(String),
}

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

/// Find the style a value names among `styles` of `style_type`, or among
/// every style when it is `None`: a style ID or, as python-docx accepts, a
/// style name.
///
/// The ID is tried first, so a value read from `Paragraph.style` always
/// finds the same style. A name then matches exactly before it matches
/// regardless of case, which is how "Heading 1" finds Word's "heading 1".
fn find_style<'s, 'a>(
    styles: &'s [rdocx::Style<'a>],
    value: &str,
    style_type: Option<rdocx::StyleType>,
) -> Option<&'s rdocx::Style<'a>> {
    let lowered = value.to_lowercase();
    let candidates = || {
        styles
            .iter()
            .filter(move |style| style_type.is_none_or(|wanted| style.style_type() == wanted))
    };
    candidates()
        .find(|style| style.style_id() == value)
        .or_else(|| candidates().find(|style| style.name() == Some(value)))
        .or_else(|| {
            candidates().find(|style| {
                style
                    .name()
                    .is_some_and(|name| name.to_lowercase() == lowered)
            })
        })
}

fn no_style_error(value: &str) -> PyErr {
    PyKeyError::new_err(format!("no style with ID or name '{value}'"))
}

/// Resolve a style of `style_type`, given by ID or name, to the ID the
/// package defines.
///
/// Each step of [`find_style`] looks at styles of that type only, so a
/// character style cannot hide a paragraph style of that name. As in
/// python-docx, a value naming no style raises `KeyError` and one naming only
/// a style of another type raises `ValueError`.
pub(crate) fn style_id_of_type(
    document: &rdocx::Document,
    value: &str,
    style_type: rdocx::StyleType,
) -> PyResult<String> {
    let styles = document.styles();
    if let Some(style) = find_style(&styles, value, Some(style_type)) {
        return Ok(style.style_id().to_owned());
    }
    match find_style(&styles, value, None) {
        Some(style) => Err(PyValueError::new_err(format!(
            "style '{value}' is a {} style, not a {} style",
            style.style_type().to_str(),
            style_type.to_str()
        ))),
        None => Err(no_style_error(value)),
    }
}

/// The type and ID of the style of any type a value names, as [`find_style`]
/// finds it, or `KeyError`.
pub(crate) fn defined_style(
    document: &rdocx::Document,
    value: &str,
) -> PyResult<(rdocx::StyleType, String)> {
    let styles = document.styles();
    find_style(&styles, value, None)
        .map(|style| (style.style_type(), style.style_id().to_owned()))
        .ok_or_else(|| no_style_error(value))
}

#[pyclass(name = "Paragraph")]
pub struct PyParagraph {
    pub(crate) document: Py<PyDocument>,
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

    pub(crate) fn belongs_to(&self, py: Python<'_>, document: &Py<PyDocument>) -> bool {
        self.document.bind(py).is(document.bind(py))
    }

    /// Check that a bookmark exists, since Word ignores a link to a
    /// missing one.
    fn bookmark_anchor(&self, py: Python<'_>, name: &str) -> PyResult<String> {
        let document = self.document.borrow(py);
        if document
            .inner
            .bookmarks()
            .iter()
            .any(|bookmark| bookmark.name() == Some(name))
        {
            Ok(name.to_owned())
        } else {
            Err(PyKeyError::new_err(format!(
                "no bookmark is named '{name}', add one with Document.add_bookmark or pass \
                 the heading paragraph itself as anchor=paragraph"
            )))
        }
    }

    /// The bookmark around a body heading paragraph, added when it has none.
    fn heading_anchor(&self, py: Python<'_>, heading: &PyParagraph) -> PyResult<String> {
        if !heading.belongs_to(py, &self.document) {
            return Err(PyValueError::new_err(
                "the heading paragraph belongs to a different document",
            ));
        }
        let ParagraphLocation::Body(paragraph) = heading.validate(py)? else {
            return Err(PyValueError::new_err(
                "a link target paragraph must be a body paragraph, bookmark a cell paragraph \
                 with Document.add_bookmark and pass its name",
            ));
        };
        let mut document = self.document.borrow_mut(py);
        let content = document
            .inner
            .content_index_of_paragraph(paragraph)
            .ok_or_else(|| {
                PyValueError::new_err(
                    "the heading paragraph is not a direct body child, bookmark it with \
                     Document.add_bookmark and pass its name",
                )
            })?;
        // Bookmark markers sit between runs, so no content moves.
        document
            .inner
            .heading_bookmark(content)
            .map_err(|error| crate::rdocx_to_pyerr(py, error))
    }

    /// Append one run, or one hyperlink run when `link` is given, and
    /// return its handle.
    fn append_run(
        &self,
        py: Python<'_>,
        text: &str,
        link: Option<LinkTarget<'_>>,
        tooltip: Option<&str>,
    ) -> PyResult<Py<PyRun>> {
        let location = self.validate(py)?;
        let (path, run_path) = {
            let mut document = self.document.borrow_mut(py);
            let relationship_id = match link {
                Some(LinkTarget::Url(url)) => {
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
                Some(LinkTarget::Anchor(_)) | None => None,
            };
            let append = |paragraph: &mut rdocx::Paragraph<'_>| {
                let run_index = paragraph.run_count();
                match (&relationship_id, &link) {
                    (Some(relationship_id), _) => {
                        paragraph.add_hyperlink_with_tooltip(text, relationship_id, tooltip);
                    }
                    (None, Some(LinkTarget::Anchor(anchor))) => {
                        paragraph.add_internal_hyperlink(text, anchor, tooltip);
                    }
                    (None, _) => {
                        paragraph.add_run(text);
                    }
                }
                let run_path = paragraph
                    .run_path(run_index)
                    .expect("the appended run has an accepted source path");
                (run_index, run_path)
            };
            let (run_index, run_path) = match location {
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
            (document.revisions.capture(segments), run_path)
        };
        Py::new(py, PyRun::new(self.document.clone_ref(py), path, run_path))
    }
}

#[pymethods]
impl PyParagraph {
    /// Replace literal text within this paragraph, retaining run formatting.
    #[pyo3(signature = (old, new, *, expect = None))]
    fn replace_text(
        &self,
        py: Python<'_>,
        old: &str,
        new: &str,
        expect: Option<usize>,
    ) -> PyResult<usize> {
        let location = self.validate(py)?;
        let mut document = self.document.borrow_mut(py);
        match location {
            ParagraphLocation::Body(index) => {
                let location = document
                    .inner
                    .paragraph_story_location(index)
                    .map_err(|error| crate::rdocx_to_pyerr(py, error))?
                    .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
                document.scoped_replacement(py, |document| {
                    document.try_replace_text_at(&location, old, new, expect)
                })
            }
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => document.scoped_replacement(py, |document| {
                document.try_replace_text_in_cell(
                    (table, row, cell),
                    Some(paragraph),
                    old,
                    new,
                    expect,
                )
            }),
        }
    }

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

    // Replace the content with one run holding `text`, keeping the paragraph
    // properties, comments and bookmarks. `None` is empty text. A success
    // advances the revision, so every earlier handle, this one included, is
    // stale.
    #[setter]
    fn set_text(&self, py: Python<'_>, text: Option<&str>) -> PyResult<()> {
        let location = self.validate(py)?;
        let text = text.unwrap_or_default();
        let mut document = self.document.borrow_mut(py);
        let result = match location {
            ParagraphLocation::Body(index) => document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?
                .set_text(text),
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
                    .set_text(text)
            }
        };
        result.map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        document.revisions.bump();
        Ok(())
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
        self.append_run(py, text, None, None)
    }

    /// Append a hyperlink run: to a web address with `url`, or inside the
    /// document with `anchor`, a bookmark name or a body heading paragraph.
    ///
    /// A heading without a bookmark around it gets one, as Google Docs' link
    /// to a heading does. A `url` that starts with `#` names a bookmark.
    #[pyo3(signature = (text, url = None, *, anchor = None, tooltip = None))]
    fn add_hyperlink(
        &self,
        py: Python<'_>,
        text: &str,
        url: Option<&str>,
        anchor: Option<&Bound<'_, PyAny>>,
        tooltip: Option<&str>,
    ) -> PyResult<Py<PyRun>> {
        let link = match (url, anchor) {
            (Some(_), Some(_)) => {
                return Err(PyTypeError::new_err(
                    "pass url for a web link or anchor for a link inside the document, not both",
                ));
            }
            (None, None) => {
                return Err(PyTypeError::new_err(
                    "add_hyperlink needs url=\"https://...\" or anchor=<bookmark name or heading paragraph>",
                ));
            }
            (Some(url), None) => match url.strip_prefix('#') {
                Some(name) => LinkTarget::Anchor(self.bookmark_anchor(py, name)?),
                None => LinkTarget::Url(url),
            },
            (None, Some(anchor)) => {
                if let Ok(heading) = anchor.cast::<PyParagraph>() {
                    LinkTarget::Anchor(self.heading_anchor(py, &heading.borrow())?)
                } else {
                    let name = anchor.extract::<String>().map_err(|_| {
                        PyTypeError::new_err("anchor is a bookmark name or a heading Paragraph")
                    })?;
                    LinkTarget::Anchor(self.bookmark_anchor(py, &name)?)
                }
            }
        };
        self.append_run(py, text, Some(link), tooltip)
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

    #[getter]
    fn style(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let location = self.validate(py)?;
        Ok(crate::formatting::paragraph_snapshot(py, &self.document, location)?.style_id)
    }

    #[setter]
    fn set_style(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        let location = self.validate(py)?;
        let value = value
            .map(|value| {
                style_id_of_type(
                    &self.document.borrow(py).inner,
                    &value,
                    rdocx::StyleType::Paragraph,
                )
            })
            .transpose()?;
        crate::formatting::apply_paragraph_update(
            py,
            &self.document,
            location,
            crate::formatting::ParagraphUpdate::Style(value),
        )
    }

    #[getter]
    fn numbering(&self, py: Python<'_>) -> PyResult<Option<(u32, u32)>> {
        let location = self.validate(py)?;
        Ok(crate::formatting::paragraph_snapshot(py, &self.document, location)?.numbering)
    }

    #[setter]
    fn set_numbering(&self, py: Python<'_>, value: Option<(u32, u32)>) -> PyResult<()> {
        let location = self.validate(py)?;
        if value.is_some_and(|(_, level)| level > 8) {
            return Err(PyValueError::new_err(
                "numbering level must be between 0 and 8",
            ));
        }
        crate::formatting::apply_paragraph_update(
            py,
            &self.document,
            location,
            crate::formatting::ParagraphUpdate::Numbering(value),
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
