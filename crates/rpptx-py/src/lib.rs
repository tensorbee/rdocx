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

use oxml_py_support::{ContentPath, PathSeg, RevisionCounter};
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

/// How far a structural edit reaches, from the widest scope to the narrowest.
///
/// A handle belongs to the scope of the last step of its path: a slide to
/// `Slides`, a shape to `Shapes`, a table row or cell to `Tables`, a paragraph
/// to `Paragraphs`, and a run to `Runs`. An edit invalidates the handles of
/// its scope and of every narrower one, so replacing a text frame's text
/// invalidates its paragraph and run handles but keeps its shape and slide
/// handles, as python-pptx keeps them. Collections without a path, such as
/// `prs.slides` and `prs.slide_layouts`, read the presentation live and stay
/// valid.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd)]
pub(crate) enum Scope {
    Slides = 0,
    Shapes = 1,
    Tables = 2,
    Paragraphs = 3,
    Runs = 4,
}

impl Scope {
    const fn name(self) -> &'static str {
        match self {
            Self::Slides => "slide",
            Self::Shapes => "shape",
            Self::Tables => "table",
            Self::Paragraphs => "paragraph",
            Self::Runs => "run",
        }
    }

    fn of(path: &ContentPath) -> Option<Self> {
        Some(match path.segs.last()? {
            PathSeg::Slide(_) => Self::Slides,
            PathSeg::Shape(_) => Self::Shapes,
            PathSeg::Row(_) | PathSeg::Cell(_) => Self::Tables,
            PathSeg::Body(_) | PathSeg::Para(_) => Self::Paragraphs,
            PathSeg::Run(_) => Self::Runs,
        })
    }
}

/// One revision counter per [`Scope`], with the call that last advanced it.
#[derive(Debug)]
pub(crate) struct HandleRevisions {
    counters: [RevisionCounter; 5],
    causes: [&'static str; 5],
}

impl HandleRevisions {
    pub(crate) fn new() -> Self {
        Self {
            counters: [RevisionCounter::new(); 5],
            causes: [""; 5],
        }
    }

    /// Invalidates the handles of `scope` and of every narrower scope,
    /// recording `cause`, the public call, for the stale-handle message.
    pub(crate) fn invalidate(&mut self, scope: Scope, cause: &'static str) {
        for level in scope as usize..self.counters.len() {
            self.counters[level].bump();
            self.causes[level] = cause;
        }
    }

    /// Captures `segs` at the revision of the scope its last step belongs to.
    pub(crate) fn capture(&self, segs: smallvec::SmallVec<[PathSeg; 5]>) -> ContentPath {
        let path = ContentPath::new(segs, 0);
        let revision = Scope::of(&path).map_or(0, |scope| self.counters[scope as usize].current());
        ContentPath::new(path.segs, revision)
    }
}

pub(crate) fn validate_path(
    py: Python<'_>,
    presentation: &PyPresentation,
    path: &ContentPath,
    kind: &str,
    suffix: &str,
) -> PyResult<()> {
    let Some(scope) = Scope::of(path) else {
        return Ok(());
    };
    let revisions = &presentation.revisions;
    let current = revisions.counters[scope as usize].current();
    if path.revision == current {
        return Ok(());
    }
    Err(public_error(
        py,
        "StaleElementError",
        format!(
            "{kind} handle was created at {name} revision {captured}, but the {name} revision \
             is now {current} because {cause} renumbered what it points to. {hint}",
            name = scope.name(),
            captured = path.revision,
            cause = revisions.causes[scope as usize],
            hint = recovery_hint(path, suffix),
        ),
    ))
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
