//! Python bindings for the rpptx facade.

mod dml;
mod layout;
mod presentation;
mod shape;
mod slide;
mod table;
mod text;

use pyo3::exceptions::{PyAttributeError, PyIndexError, PyRuntimeError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyType};

use oxml_py_support::{ContentPath, PathSeg, StaleElementError};
use presentation::{PyComment, PyCommentAuthor, PyCommentReply, PyPresentation, PyValidationIssue};

pub(crate) fn normalize_index(index: isize, len: usize, kind: &str) -> PyResult<usize> {
    let normalized = if index < 0 {
        len as isize + index
    } else {
        index
    };
    if normalized < 0 || normalized >= len as isize {
        return Err(PyIndexError::new_err(format!("{kind} index out of range")));
    }
    Ok(normalized as usize)
}

fn public_exception_type<'py>(py: Python<'py>, class_name: &str) -> PyResult<Bound<'py, PyType>> {
    py.import("rpptx")
        .and_then(|module| module.getattr(class_name))
        .and_then(|class| class.cast_into::<PyType>().map_err(Into::into))
}

fn public_error(py: Python<'_>, class_name: &str, message: String) -> PyErr {
    match public_exception_type(py, class_name) {
        Ok(class) => PyErr::from_type(class, (message,)),
        Err(_) => PyRuntimeError::new_err(message),
    }
}

/// A counted replacement that matched a different number of times than the
/// caller expected, worded like `rpptx replace --expect`.
pub(crate) fn replacement_count_to_pyerr(
    py: Python<'_>,
    placeholder: &str,
    expected: usize,
    found: usize,
) -> PyErr {
    let message = format!("expected {expected} replacement(s) of \"{placeholder}\", found {found}");
    match public_exception_type(py, "ReplacementCountError") {
        Ok(class) => PyErr::from_type(class, (message, expected, found)),
        Err(_) => PyRuntimeError::new_err(message),
    }
}

/// python-pptx names that rpptx spells differently or does not have, by
/// class and attribute, and the rpptx way to do the same.
const DIVERGENCES: &[(&str, &str, &str)] = &[
    (
        "SlideCollection",
        "_sldIdLst",
        "use prs.slides.remove(slide) to delete a slide and prs.slides.move(old_index, new_index) to reorder slides",
    ),
    (
        "Shape",
        "insert_picture",
        "add the picture with slide.shapes.add_picture(image, placeholder.left, placeholder.top, placeholder.width, placeholder.height), then remove the placeholder with slide.shapes.remove(placeholder)",
    ),
    (
        "Shape",
        "placeholder_format",
        "find a placeholder by its idx with slide.placeholders[idx], and read shape.name to tell placeholders apart",
    ),
    (
        "Shape",
        "is_placeholder",
        "find placeholders with slide.placeholders[idx] or slide.shapes.placeholders",
    ),
    (
        "Cell",
        "text_frame",
        "set cell.text, or edit the cell's a:txBody in the table shape's XML with shape.replace_xml(xml)",
    ),
    (
        "Cell",
        "vertical_anchor",
        "edit the cell's a:tcPr anchor attribute in the table shape's XML with shape.replace_xml(xml)",
    ),
    (
        "NotesTextFrame",
        "paragraphs",
        "notes are read and written as plain text here: use notes_text_frame.text or slide.notes_text",
    ),
    (
        "NotesTextFrame",
        "add_paragraph",
        "append a line to notes_text_frame.text, for example text_frame.text += '\\nNext point'",
    ),
    (
        "Presentation",
        "slide_master",
        "edit the layouts a deck uses through prs.slide_layouts[i].xml and replace_xml(xml)",
    ),
];

/// The lxml handles python-pptx scripts reach for, which have no rpptx
/// counterpart: raw XML is read and replaced instead.
const LXML_NAMES: &[&str] = &[
    "_element", "element", "_sp", "_txBody", "_sld", "_r", "_p", "_tbl", "_tc",
];

/// The handles that read and replace their own XML.
const XML_OWNERS: &[&str] = &["Shape", "TextFrame", "Slide", "SlideLayout"];

/// An `AttributeError` for `class.name` that names the rpptx way when
/// python-pptx spells it differently.
pub(crate) fn missing_attribute(class: &str, name: &str) -> PyErr {
    let hint = DIVERGENCES
        .iter()
        .find(|(owner, attribute, _)| *owner == class && *attribute == name)
        .map(|(.., hint)| (*hint).to_owned())
        .or_else(|| {
            LXML_NAMES.contains(&name).then(|| {
                if XML_OWNERS.contains(&class) {
                    let handle = match class {
                        "TextFrame" => "text_frame".to_owned(),
                        "SlideLayout" => "slide_layout".to_owned(),
                        other => other.to_lowercase(),
                    };
                    format!(
                        "rpptx has no lxml element: read {handle}.xml and write {handle}.replace_xml(xml)"
                    )
                } else {
                    "rpptx has no lxml element: read shape.xml or shape.text_frame.xml and write them back with replace_xml(xml)".to_owned()
                }
            })
        });
    match hint {
        Some(hint) => PyAttributeError::new_err(format!(
            "'{class}' object has no attribute '{name}': {hint}"
        )),
        None => PyAttributeError::new_err(format!("'{class}' object has no attribute '{name}'")),
    }
}

/// Set `name` through the property the class defines, or raise
/// [`missing_attribute`]: a pyclass without `__dict__` would otherwise say
/// only that the attribute does not exist.
pub(crate) fn set_attribute(
    slf: &Bound<'_, PyAny>,
    class: &str,
    name: &str,
    value: &Bound<'_, PyAny>,
) -> PyResult<()> {
    match slf.get_type().getattr(name) {
        Ok(descriptor) if descriptor.hasattr("__set__")? => {
            descriptor.call_method1("__set__", (slf, value))?;
            Ok(())
        }
        _ => Err(missing_attribute(class, name)),
    }
}

/// The bytes of a `replace_xml` argument: `str`, `bytes` or `bytearray`.
pub(crate) fn raw_xml_argument(xml: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    if let Ok(text) = xml.extract::<String>() {
        return Ok(text.into_bytes());
    }
    xml.extract::<Vec<u8>>()
        .map_err(|_| PyTypeError::new_err("xml must be str or bytes"))
}

/// A refused raw XML replacement is a `ValueError` that says why.
pub(crate) fn raw_xml_error(py: Python<'_>, error: rpptx::Error) -> PyErr {
    match error {
        rpptx::Error::InvalidShapeMutation {
            operation: "replace_xml",
            message,
        } => PyValueError::new_err(message),
        error => rpptx_to_pyerr(py, error),
    }
}

pub(crate) fn stale_to_pyerr(py: Python<'_>, error: StaleElementError) -> PyErr {
    public_error(py, "StaleElementError", error.to_string())
}

pub(crate) fn recovery_hint(path: &ContentPath, suffix: &str) -> String {
    let mut public_path = String::from("prs");
    let mut pending_row = None;
    for segment in &path.segs {
        match segment {
            PathSeg::Slide(index) => public_path.push_str(&format!(".slides[{index}]")),
            PathSeg::Shape(index) => public_path.push_str(&format!(".shapes[{index}]")),
            PathSeg::Body(index) => public_path.push_str(&format!(".body[{index}]")),
            PathSeg::Row(index) => pending_row = Some(*index),
            PathSeg::Cell(index) => {
                if let Some(row) = pending_row.take() {
                    public_path.push_str(&format!(".table.cell({row}, {index})"));
                } else {
                    public_path.push_str(&format!(".table.columns[{index}]"));
                }
            }
            PathSeg::Para(index) => {
                public_path.push_str(&format!(".text_frame.paragraphs[{index}]"));
            }
            PathSeg::Run(index) => public_path.push_str(&format!(".runs[{index}]")),
        }
    }
    if let Some(row) = pending_row {
        public_path.push_str(&format!(".table.rows[{row}]"));
    }
    public_path.push_str(suffix);
    format!("Re-fetch it with {public_path}.")
}

pub(crate) fn validate_path(
    py: Python<'_>,
    presentation: &PyPresentation,
    path: &ContentPath,
    kind: &str,
    suffix: &str,
) -> PyResult<()> {
    path.validate_revision(
        presentation.revisions.current(),
        kind,
        &recovery_hint(path, suffix),
    )
    .map_err(|error| stale_to_pyerr(py, error))
}

pub(crate) fn rpptx_to_pyerr(py: Python<'_>, error: rpptx::Error) -> PyErr {
    let class_name = match &error {
        rpptx::Error::MalformedPart { .. } => "XmlError",
        rpptx::Error::Package(_)
        | rpptx::Error::MissingMainDocument
        | rpptx::Error::MissingRelationship { .. }
        | rpptx::Error::WrongRelationshipType { .. }
        | rpptx::Error::ExternalRelationship { .. }
        | rpptx::Error::MissingPart { .. }
        | rpptx::Error::CorePropertiesPartCollision { .. }
        | rpptx::Error::DuplicateNotesSlides { .. } => "PackageError",
        _ => "RpptxError",
    };
    public_error(py, class_name, error.to_string())
}

pub(crate) fn rpptx_value_to_pyerr(py: Python<'_>, message: String) -> PyErr {
    public_error(py, "RpptxError", message)
}

#[pymodule]
fn _rpptx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyPresentation>()?;
    module.add_class::<PyCommentAuthor>()?;
    module.add_class::<PyComment>()?;
    module.add_class::<PyCommentReply>()?;
    module.add_class::<PyValidationIssue>()?;
    slide::register(module)?;
    shape::register(module)?;
    dml::register(module)?;
    text::register(module)?;
    table::register(module)?;
    layout::register(module)?;
    Ok(())
}
