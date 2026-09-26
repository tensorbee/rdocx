//! Fill, line, and colour formats, mirroring python-pptx `pptx.dml`.
//!
//! Shape fills, line fills, and slide backgrounds share one fill model, so
//! the three formats read and write through a single target.

use oxml_py_support::ContentPath;
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;

use crate::presentation::PyPresentation;
use crate::shape::{length, shape_mut_at, shape_ref_at, slide_index};
use crate::{rpptx_to_pyerr, validate_path};

const MAX_LINE_WIDTH_EMU: i64 = 20_116_800;

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyFillFormat>()?;
    module.add_class::<PyLineFormat>()?;
    module.add_class::<PyColorFormat>()?;
    Ok(())
}

/// The DrawingML fill one format object reads and writes.
#[derive(Clone, Copy, Eq, PartialEq)]
pub(crate) enum FillTarget {
    Shape,
    Line,
    Background,
}

impl FillTarget {
    fn suffix(self) -> &'static str {
        match self {
            Self::Shape => ".fill",
            Self::Line => ".line.fill",
            Self::Background => ".background.fill",
        }
    }
}

fn current_fill(
    presentation: &rpptx::Presentation,
    path: &ContentPath,
    target: FillTarget,
) -> PyResult<Option<rpptx::Fill>> {
    let missing = || PyIndexError::new_err("shape index out of range");
    Ok(match target {
        FillTarget::Shape => shape_ref_at(presentation, path)
            .ok_or_else(missing)?
            .fill()
            .cloned(),
        FillTarget::Line => shape_ref_at(presentation, path)
            .ok_or_else(missing)?
            .line()
            .and_then(|line| line.fill.clone()),
        FillTarget::Background => presentation
            .slide(slide_index(path)?)
            .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?
            .background_fill()
            .cloned(),
    })
}

fn write_fill(
    py: Python<'_>,
    presentation: &mut rpptx::Presentation,
    path: &ContentPath,
    target: FillTarget,
    fill: rpptx::Fill,
) -> PyResult<()> {
    let missing = || PyIndexError::new_err("shape index out of range");
    let result = match target {
        FillTarget::Shape => shape_mut_at(presentation, path)
            .ok_or_else(missing)?
            .set_fill(fill),
        FillTarget::Line => {
            let mut line = shape_ref_at(presentation, path)
                .ok_or_else(missing)?
                .line()
                .cloned()
                .unwrap_or_default();
            line.fill = Some(fill);
            shape_mut_at(presentation, path)
                .ok_or_else(missing)?
                .set_line(line)
        }
        FillTarget::Background => {
            let index = slide_index(path)?;
            let direct = presentation
                .slide(index)
                .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?
                .background_fill()
                .is_some();
            let mut slide = presentation
                .slide_mut(index)
                .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?;
            if !direct {
                slide.remove_background();
            }
            slide.set_background(fill)
        }
    };
    result.map_err(|error| rpptx_to_pyerr(py, error))
}

fn fill_type_name(fill: Option<&rpptx::Fill>) -> &'static str {
    match fill {
        None => "_NoneFill",
        Some(rpptx::Fill::NoFill(_)) => "_NoFill",
        Some(rpptx::Fill::Solid(_)) => "_SolidFill",
        Some(rpptx::Fill::Gradient(_)) => "_GradFill",
        Some(rpptx::Fill::Pattern(_)) => "_PattFill",
        Some(rpptx::Fill::Blip(_)) => "_BlipFill",
    }
}

fn no_foreground(fill: Option<&rpptx::Fill>) -> PyErr {
    PyTypeError::new_err(format!(
        "fill type {} has no foreground color, call .solid() or .patterned() first",
        fill_type_name(fill)
    ))
}

fn with_rgb(color: Option<rpptx::ColorChoice>, rgb: rpptx::RgbColor) -> rpptx::ColorChoice {
    match color {
        Some(rpptx::ColorChoice::Srgb {
            transforms,
            raw_children,
            ..
        }) => rpptx::ColorChoice::Srgb {
            value: rgb,
            transforms,
            raw_children,
        },
        _ => rpptx::ColorChoice::srgb(rgb),
    }
}

/// A live view of one shape fill, line fill, or slide background fill.
#[pyclass(name = "FillFormat")]
pub struct PyFillFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    target: FillTarget,
}

impl PyFillFormat {
    pub(crate) fn new(
        presentation: Py<PyPresentation>,
        path: ContentPath,
        target: FillTarget,
    ) -> Self {
        Self {
            presentation,
            path,
            target,
        }
    }

    fn fill(&self, py: Python<'_>) -> PyResult<Option<rpptx::Fill>> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "fill", self.target.suffix())?;
        current_fill(&presentation.inner, &self.path, self.target)
    }

    fn set(&self, py: Python<'_>, fill: rpptx::Fill) -> PyResult<()> {
        let mut presentation = self.presentation.borrow_mut(py);
        write_fill(py, &mut presentation.inner, &self.path, self.target, fill)
    }

    /// Only a shape fill can take its group's fill. A line fill or a slide
    /// background has no `a:grpFill` choice.
    fn has_group_fill(&self, py: Python<'_>) -> PyResult<bool> {
        if !matches!(self.target, FillTarget::Shape) {
            return Ok(false);
        }
        let presentation = self.presentation.borrow(py);
        Ok(shape_ref_at(&presentation.inner, &self.path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .has_group_fill())
    }
}

#[pymethods]
impl PyFillFormat {
    #[getter(r#type)]
    fn fill_type(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let value = match self.fill(py)? {
            None if self.has_group_fill(py)? => 101,
            None => return Ok(None),
            Some(rpptx::Fill::Solid(_)) => 1,
            Some(rpptx::Fill::Pattern(_)) => 2,
            Some(rpptx::Fill::Gradient(_)) => 3,
            Some(rpptx::Fill::NoFill(_)) => 5,
            Some(rpptx::Fill::Blip(_)) => 6,
        };
        py.import("rpptx.enum.dml")?
            .getattr("MSO_FILL_TYPE")?
            .call1((value,))
            .map(|member| Some(member.unbind()))
    }

    /// Sets a solid fill, keeping an existing solid fill and its colour.
    fn solid(&self, py: Python<'_>) -> PyResult<()> {
        if matches!(self.fill(py)?, Some(rpptx::Fill::Solid(_))) {
            return Ok(());
        }
        self.set(py, rpptx::Fill::Solid(rpptx::SolidFill::default()))
    }

    /// Sets an explicit no-fill, so what lies behind shows through.
    fn background(&self, py: Python<'_>) -> PyResult<()> {
        self.fill(py)?;
        self.set(py, rpptx::Fill::NoFill(rpptx::NoFill::default()))
    }

    #[getter]
    fn fore_color(&self, py: Python<'_>) -> PyResult<Py<PyColorFormat>> {
        let fill = self.fill(py)?;
        if !matches!(fill, Some(rpptx::Fill::Solid(_) | rpptx::Fill::Pattern(_))) {
            return Err(no_foreground(fill.as_ref()));
        }
        Py::new(
            py,
            PyColorFormat {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                target: self.target,
                solidify: false,
            },
        )
    }
}

/// A live view of the foreground colour of one fill.
#[pyclass(name = "ColorFormat")]
pub struct PyColorFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    target: FillTarget,
    solidify: bool,
}

#[pymethods]
impl PyColorFormat {
    #[getter]
    fn rgb(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "color", self.target.suffix())?;
        let color = match current_fill(&presentation.inner, &self.path, self.target)? {
            Some(rpptx::Fill::Solid(fill)) => fill.color,
            Some(rpptx::Fill::Pattern(fill)) => fill.foreground,
            _ => None,
        };
        drop(presentation);
        let Some(rpptx::ColorChoice::Srgb { value, .. }) = color else {
            return Ok(None);
        };
        let [red, green, blue] = value.components();
        py.import("rpptx.dml.color")?
            .getattr("RGBColor")?
            .call1((red, green, blue))
            .map(|color| Some(color.unbind()))
    }

    #[setter]
    fn set_rgb(&self, py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let rgb_class = py.import("rpptx.dml.color")?.getattr("RGBColor")?;
        if !value.is_instance(&rgb_class)? {
            return Err(PyValueError::new_err(
                "assigned value must be type RGBColor",
            ));
        }
        let (red, green, blue) = value.extract::<(u8, u8, u8)>()?;
        let rgb = rpptx::RgbColor::new(red, green, blue);
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(py, &presentation, &self.path, "color", self.target.suffix())?;
        let fill = match current_fill(&presentation.inner, &self.path, self.target)? {
            Some(rpptx::Fill::Solid(mut fill)) => {
                fill.color = Some(with_rgb(fill.color.take(), rgb));
                rpptx::Fill::Solid(fill)
            }
            // A line colour makes any other line fill solid, pattern included.
            _ if self.solidify => {
                let mut fill = rpptx::SolidFill::default();
                fill.color = Some(rpptx::ColorChoice::srgb(rgb));
                rpptx::Fill::Solid(fill)
            }
            Some(rpptx::Fill::Pattern(mut fill)) => {
                fill.foreground = Some(with_rgb(fill.foreground.take(), rgb));
                rpptx::Fill::Pattern(fill)
            }
            other => return Err(no_foreground(other.as_ref())),
        };
        write_fill(py, &mut presentation.inner, &self.path, self.target, fill)
    }
}

/// A live view of the outline of one shape, picture, or connector.
#[pyclass(name = "LineFormat")]
pub struct PyLineFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyLineFormat {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "line",
            ".line",
        )
    }
}

#[pymethods]
impl PyLineFormat {
    /// Returns the line colour. Setting its `rgb` makes the line fill solid.
    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Py<PyColorFormat>> {
        self.validate(py)?;
        Py::new(
            py,
            PyColorFormat {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                target: FillTarget::Line,
                solidify: true,
            },
        )
    }

    #[getter]
    fn fill(&self, py: Python<'_>) -> PyResult<Py<PyFillFormat>> {
        self.validate(py)?;
        Py::new(
            py,
            PyFillFormat::new(
                self.presentation.clone_ref(py),
                self.path.clone(),
                FillTarget::Line,
            ),
        )
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let width = shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .line()
            .and_then(|line| line.width)
            .unwrap_or(0);
        length(py, Some(rpptx::Emu(i64::from(width))))
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.validate(py)?;
        let width = value.unwrap_or(0);
        if !(0..=MAX_LINE_WIDTH_EMU).contains(&width) {
            return Err(PyValueError::new_err(format!(
                "line width must be between 0 and {MAX_LINE_WIDTH_EMU} EMU, got {width}"
            )));
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let missing = || PyIndexError::new_err("shape index out of range");
        let mut line = shape_ref_at(&presentation.inner, &self.path)
            .ok_or_else(missing)?
            .line()
            .cloned()
            .unwrap_or_default();
        // Zero is the schema default, so it is written as an absent width,
        // as python-pptx does.
        line.width = (width != 0).then_some(width as u32);
        shape_mut_at(&mut presentation.inner, &self.path)
            .ok_or_else(missing)?
            .set_line(line)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}
