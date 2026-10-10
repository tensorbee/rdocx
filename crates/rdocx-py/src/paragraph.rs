use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{
    PyIndexError, PyKeyError, PyNotImplementedError, PyRuntimeError, PyTypeError, PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyBytes, PyList, PySlice};
use smallvec::smallvec;

use crate::document::{EditScope, PyDocument, raw_xml_argument, raw_xml_error};
use crate::normalize_index;
use crate::run::{PyRun, PyRunCollection};

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

/// The raw XML address of the paragraph at `location`.
pub(crate) fn xml_paragraph(location: ParagraphLocation) -> rdocx::XmlParagraph {
    match location {
        ParagraphLocation::Body(index) => rdocx::XmlParagraph::Body(index),
        ParagraphLocation::Cell {
            table,
            row,
            cell,
            paragraph,
        } => rdocx::XmlParagraph::Cell {
            table,
            row,
            cell,
            paragraph,
        },
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

/// The list styles of python-docx's default template, by name, style ID,
/// whether they are bulleted, and list level.
const BUILTIN_LIST_STYLES: [(&str, &str, bool, u32); 6] = [
    ("List Bullet", "ListBullet", true, 0),
    ("List Bullet 2", "ListBullet2", true, 1),
    ("List Bullet 3", "ListBullet3", true, 2),
    ("List Number", "ListNumber", false, 0),
    ("List Number 2", "ListNumber2", false, 1),
    ("List Number 3", "ListNumber3", false, 2),
];

/// Add the python-docx list style `value` names when the document lacks it.
///
/// A first draft writes `add_paragraph(text, style='List Bullet')`, and
/// python-docx's template defines that style. A document without it gets
/// the style, linked to the list definition its siblings share (one for the
/// bulleted styles, one for the numbered ones), so the paragraph is a list
/// item in Word, LibreOffice and Google Docs. Any other value is left to
/// the style lookup.
pub(crate) fn ensure_builtin_style(
    py: Python<'_>,
    document: &mut rdocx::Document,
    value: &str,
) -> PyResult<()> {
    let Some(&(name, style_id, bullet, level)) = BUILTIN_LIST_STYLES
        .iter()
        .find(|(name, id, ..)| name.eq_ignore_ascii_case(value) || *id == value)
    else {
        return Ok(());
    };
    if find_style(&document.styles(), value, None).is_some() {
        return Ok(());
    }
    let shared = BUILTIN_LIST_STYLES
        .iter()
        .filter(|(.., sibling_bullet, _)| *sibling_bullet == bullet)
        .find_map(|(_, sibling, ..)| {
            document
                .style(sibling)?
                .paragraph_properties()?
                .num_id
                .filter(|num_id| *num_id != 0)
        });
    let num_id = match shared {
        Some(num_id) => num_id,
        None => document.add_list_definition(&[if bullet {
            rdocx::ListLevel::bullet()
        } else {
            rdocx::ListLevel::decimal()
        }]),
    };
    let mut builder = rdocx::StyleBuilder::paragraph(style_id, name);
    if document.style("Normal").is_some() {
        builder = builder.based_on("Normal");
    }
    document
        .add_style(builder)
        .and_then(|()| document.link_style_to_numbering(style_id, num_id, level))
        .map_err(|error| crate::rdocx_to_pyerr(py, error))
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
        paragraph_location(&self.resolved(py)?)
    }

    /// This handle's path at the current revision, moved by the edits made
    /// since it was created.
    pub(crate) fn resolved(&self, py: Python<'_>) -> PyResult<ContentPath> {
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
        self.document
            .borrow(py)
            .resolve_path(py, &self.path, "paragraph", &recovery_hint)
    }

    pub(crate) fn belongs_to(&self, py: Python<'_>, document: &Py<PyDocument>) -> bool {
        self.document.bind(py).is(document.bind(py))
    }

    /// Append one run, or one external hyperlink run when `url` is given, and
    /// return its handle.
    fn append_run(&self, py: Python<'_>, text: &str, url: Option<&str>) -> PyResult<Py<PyRun>> {
        let resolved = self.resolved(py)?;
        let location = paragraph_location(&resolved)?;
        let (path, run_path) = {
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
            let call = match url {
                Some(_) => "Paragraph.add_hyperlink",
                None => "Paragraph.add_run",
            };
            document.bump_edit(call, EditScope::Appended);
            let mut segments = resolved.segs;
            segments.push(PathSeg::Run(run_index));
            (document.revisions.capture(segments), run_path)
        };
        Py::new(py, PyRun::new(self.document.clone_ref(py), path, run_path))
    }
}

#[pymethods]
impl PyParagraph {
    fn __getattr__(&self, name: &str) -> PyResult<Py<PyAny>> {
        Err(crate::missing_attribute("Paragraph", name))
    }

    fn __setattr__(slf: &Bound<'_, Self>, name: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        crate::set_attribute(slf.as_any(), "Paragraph", name, value)
    }

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
                document.scoped_replacement(py, "Paragraph.replace_text", |document| {
                    document.try_replace_text_at(&location, old, new, expect)
                })
            }
            ParagraphLocation::Cell {
                table,
                row,
                cell,
                paragraph,
            } => document.scoped_replacement(py, "Paragraph.replace_text", |document| {
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
    // retires the run handles of this paragraph, and keeps every other handle.
    #[setter]
    fn set_text(&self, py: Python<'_>, text: Option<&str>) -> PyResult<()> {
        let resolved = self.resolved(py)?;
        let location = paragraph_location(&resolved)?;
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
        document.bump_edit("Paragraph.text", EditScope::Below(resolved.segs));
        Ok(())
    }

    #[getter]
    fn runs(&self, py: Python<'_>) -> PyResult<Py<PyRunCollection>> {
        let resolved = self.resolved(py)?;
        Py::new(
            py,
            PyRunCollection::new(self.document.clone_ref(py), resolved),
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

    /// This paragraph's `w:p` element as standalone XML bytes.
    #[getter]
    fn xml<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let target = rdocx::XmlTarget::Paragraph(xml_paragraph(self.validate(py)?));
        let xml = self
            .document
            .borrow(py)
            .inner
            .element_xml(&target)
            .map_err(|error| raw_xml_error(py, error))?;
        Ok(PyBytes::new(py, &xml))
    }

    /// Replace this paragraph with one `w:p` element given as XML. The
    /// handle stays valid and its run handles retire.
    fn replace_xml(&self, py: Python<'_>, xml: &Bound<'_, PyAny>) -> PyResult<()> {
        let xml = raw_xml_argument(xml)?;
        let resolved = self.resolved(py)?;
        let target = rdocx::XmlTarget::Paragraph(xml_paragraph(paragraph_location(&resolved)?));
        let mut document = self.document.borrow_mut(py);
        document
            .inner
            .replace_element_xml(&target, &xml)
            .map_err(|error| raw_xml_error(py, error))?;
        document.bump_edit("Paragraph.replace_xml", EditScope::Below(resolved.segs));
        Ok(())
    }

    /// Insert a paragraph before this body paragraph and return it, as
    /// python-docx does. This handle and every other one stay valid.
    #[pyo3(signature = (text = "", style = None))]
    fn insert_paragraph_before(
        &self,
        py: Python<'_>,
        text: &str,
        style: Option<String>,
    ) -> PyResult<Py<PyParagraph>> {
        let ParagraphLocation::Body(index) = self.validate(py)? else {
            return Err(PyNotImplementedError::new_err(
                "insert_paragraph_before works on body paragraphs, use cell.add_paragraph(text) in a table cell",
            ));
        };
        let content = self
            .document
            .borrow(py)
            .inner
            .content_index_of_paragraph(index)
            .ok_or_else(|| {
                PyValueError::new_err(
                    "this paragraph sits inside a content control, insert next to the control with document.insert_paragraph(index, text)",
                )
            })?;
        let paragraph =
            PyDocument::insert_paragraph(self.document.clone_ref(py), py, content, text)?;
        if style.is_some() {
            paragraph.borrow(py).set_style(py, style)?;
        }
        Ok(paragraph)
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
        let resolved = self.resolved(py)?;
        Py::new(
            py,
            crate::formatting::PyParagraphFormat::new(self.document.clone_ref(py), resolved),
        )
    }

    #[getter]
    fn style(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let location = self.validate(py)?;
        Ok(crate::formatting::paragraph_snapshot(py, &self.document, location)?.style_id)
    }

    #[setter]
    pub(crate) fn set_style(&self, py: Python<'_>, value: Option<String>) -> PyResult<()> {
        let location = self.validate(py)?;
        if let Some(value) = &value {
            ensure_builtin_style(py, &mut self.document.borrow_mut(py).inner, value)?;
        }
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
