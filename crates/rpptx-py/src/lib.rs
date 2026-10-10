//! Python bindings for the rpptx facade.

mod dml;
mod layout;
mod presentation;
mod shape;
mod slide;
mod table;
mod text;

use pyo3::exceptions::{PyIndexError, PyRuntimeError};
use pyo3::prelude::*;
use pyo3::types::PyType;

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
            // Presentation handles never capture a story segment.
            PathSeg::Story(_) => {}
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
