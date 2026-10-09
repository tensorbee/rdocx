//! Fill, line, shadow, and colour formats, mirroring python-pptx `pptx.dml`.
//!
//! Shape fills, line fills, slide backgrounds, table cell fills, and cell
//! border fills share one fill model, so the formats read and write through a
//! single target.

use oxml_py_support::ContentPath;
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;

use crate::presentation::PyPresentation;
use crate::shape::{
    ANGLE_UNITS_PER_DEGREE, ANGLE_UNITS_PER_TURN, MAX_COORDINATE, image_bytes, length, part_of,
    shape_mut_at, shape_ref_at,
};
use crate::table::{cell_mut_at, cell_ref_at};
use crate::{rpptx_to_pyerr, validate_path};

const MAX_LINE_WIDTH_EMU: i64 = 20_116_800;

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyFillFormat>()?;
    module.add_class::<PyLineFormat>()?;
    module.add_class::<PyLineEndFormat>()?;
    module.add_class::<PyColorFormat>()?;
    module.add_class::<PyShadowFormat>()?;
    module.add_class::<PyGradientStops>()?;
    module.add_class::<PyGradientStop>()?;
    Ok(())
}

/// The DrawingML fill one format object reads and writes.
///
/// `Line` and `CellBorder` also name the line a `LineFormat` reads and writes.
#[derive(Clone, Copy, Eq, PartialEq)]
pub(crate) enum FillTarget {
    Shape,
    Line,
    Background,
    TableCell,
    CellBorder(rpptx::CellBorder),
}

impl FillTarget {
    fn suffix(self) -> &'static str {
        match self {
            Self::Shape | Self::TableCell => ".fill",
            Self::Line => ".line.fill",
            Self::Background => ".background.fill",
            Self::CellBorder(rpptx::CellBorder::Left) => ".border_left.fill",
            Self::CellBorder(rpptx::CellBorder::Right) => ".border_right.fill",
            Self::CellBorder(rpptx::CellBorder::Top) => ".border_top.fill",
            Self::CellBorder(rpptx::CellBorder::Bottom) => ".border_bottom.fill",
        }
    }

    fn line_suffix(self) -> &'static str {
        self.suffix()
            .strip_suffix(".fill")
            .expect("every fill suffix ends in .fill")
    }
}

/// Reads the shape line or cell border a line target names.
fn current_line(
    presentation: &rpptx::Presentation,
    path: &ContentPath,
    target: FillTarget,
) -> PyResult<Option<rpptx::CT_LineProperties>> {
    Ok(match target {
        FillTarget::CellBorder(edge) => cell_ref_at(presentation, path)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
            .border(edge)
            .cloned(),
        _ => shape_ref_at(presentation, path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .line()
            .cloned(),
    })
}

/// Replaces the shape line or cell border a line target names.
fn write_line(
    py: Python<'_>,
    presentation: &mut rpptx::Presentation,
    path: &ContentPath,
    target: FillTarget,
    line: rpptx::CT_LineProperties,
) -> PyResult<()> {
    match target {
        FillTarget::CellBorder(edge) => {
            cell_mut_at(presentation, path)
                .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
                .set_border(edge, Some(line));
            Ok(())
        }
        _ => shape_mut_at(presentation, path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .set_line(line)
            .map_err(|error| rpptx_to_pyerr(py, error)),
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
        FillTarget::Line | FillTarget::CellBorder(_) => {
            current_line(presentation, path, target)?.and_then(|line| line.fill)
        }
        FillTarget::Background => presentation
            .part_background_fill(part_of(path)?)
            .map_err(|error| PyIndexError::new_err(error.to_string()))?
            .cloned(),
        FillTarget::TableCell => cell_ref_at(presentation, path)
            .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
            .fill()
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
        FillTarget::Line | FillTarget::CellBorder(_) => {
            let mut line = current_line(presentation, path, target)?.unwrap_or_default();
            line.fill = Some(fill);
            return write_line(py, presentation, path, target, line);
        }
        FillTarget::Background => presentation.set_part_background(part_of(path)?, fill),
        FillTarget::TableCell => {
            cell_mut_at(presentation, path)
                .ok_or_else(|| PyIndexError::new_err("cell index out of range"))?
                .set_fill(Some(fill));
            Ok(())
        }
    };
    result.map_err(|error| rpptx_to_pyerr(py, error))
}

/// Changes the direct line of a shape or cell border. Without one, `a:ln` is created only
/// when `create` is set, so a removal never adds an empty line.
fn update_line(
    py: Python<'_>,
    presentation: &mut rpptx::Presentation,
    path: &ContentPath,
    target: FillTarget,
    create: bool,
    change: impl FnOnce(&mut rpptx::CT_LineProperties),
) -> PyResult<()> {
    let line = current_line(presentation, path, target)?;
    let mut line = match line {
        Some(line) => line,
        None if create => rpptx::CT_LineProperties::default(),
        None => return Ok(()),
    };
    change(&mut line);
    write_line(py, presentation, path, target, line)
}

fn dml_enum(py: Python<'_>, name: &str, value: i32) -> PyResult<Py<PyAny>> {
    py.import("rpptx.enum.dml")?
        .getattr(name)?
        .call1((value,))
        .map(Bound::unbind)
}

/// `MSO_LINE_DASH_STYLE` values, with the python-pptx mapping for the nine
/// members it shares.
fn dash_value(dash: rpptx::ST_PresetLineDashVal) -> i32 {
    use rpptx::ST_PresetLineDashVal as Dash;
    match dash {
        Dash::Solid => 1,
        Dash::SystemDash => 2,
        Dash::SystemDot => 3,
        Dash::Dash => 4,
        Dash::DashDot => 5,
        Dash::LargeDashDotDot => 6,
        Dash::LargeDash => 7,
        Dash::LargeDashDot => 8,
        Dash::SystemDashDot => 12,
        Dash::Dot => 13,
        Dash::SystemDashDotDot => 14,
    }
}

fn dash_from_value(value: i32) -> PyResult<rpptx::ST_PresetLineDashVal> {
    rpptx::ST_PresetLineDashVal::ALL
        .into_iter()
        .find(|dash| dash_value(*dash) == value)
        .ok_or_else(|| {
            PyValueError::new_err(
                "dash style must be an MSO_LINE_DASH_STYLE member other than DASH_STYLE_MIXED",
            )
        })
}

/// `MSO_ARROWHEAD_STYLE` values, as `MsoArrowheadStyle` numbers them.
fn end_type_value(kind: rpptx::LineEndType) -> i32 {
    match kind {
        rpptx::LineEndType::None => 1,
        rpptx::LineEndType::Triangle => 2,
        rpptx::LineEndType::Arrow => 3,
        rpptx::LineEndType::Stealth => 4,
        rpptx::LineEndType::Diamond => 5,
        rpptx::LineEndType::Oval => 6,
    }
}

fn end_type_from_value(value: i32) -> PyResult<rpptx::LineEndType> {
    Ok(match value {
        1 => rpptx::LineEndType::None,
        2 => rpptx::LineEndType::Triangle,
        3 => rpptx::LineEndType::Arrow,
        4 => rpptx::LineEndType::Stealth,
        5 => rpptx::LineEndType::Diamond,
        6 => rpptx::LineEndType::Oval,
        _ => {
            return Err(PyValueError::new_err(
                "line end type must be an MSO_ARROWHEAD_STYLE member",
            ));
        }
    })
}

/// `MSO_ARROWHEAD_WIDTH` and `MSO_ARROWHEAD_LENGTH` share these values.
fn end_size_value(size: rpptx::LineEndSize) -> i32 {
    match size {
        rpptx::LineEndSize::Small => 1,
        rpptx::LineEndSize::Medium => 2,
        rpptx::LineEndSize::Large => 3,
    }
}

fn end_size_from_value(value: i32, name: &str) -> PyResult<rpptx::LineEndSize> {
    Ok(match value {
        1 => rpptx::LineEndSize::Small,
        2 => rpptx::LineEndSize::Medium,
        3 => rpptx::LineEndSize::Large,
        _ => {
            return Err(PyValueError::new_err(format!(
                "line end size must be an {name} member"
            )));
        }
    })
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

/// Reads a colour argument: an `RGBColor` or any `(r, g, b)` triple of 0
/// to 255 integers, or a six-digit hex string with or without `#`.
pub(crate) fn color_argument(
    value: &Bound<'_, PyAny>,
    parameter: &str,
) -> PyResult<rpptx::RgbColor> {
    if let Ok(text) = value.extract::<String>() {
        let digits = text.strip_prefix('#').unwrap_or(&text);
        return rpptx::RgbColor::parse(digits).map_err(|_| {
            PyValueError::new_err(format!(
                "{parameter} must be six hexadecimal digits such as \"#1A237E\", got {text:?}"
            ))
        });
    }
    let (red, green, blue) = value.extract::<(i64, i64, i64)>().map_err(|_| {
        PyTypeError::new_err(format!(
            "{parameter} must be an RGBColor, a hex string such as \"#1A237E\" or an (r, g, b) triple"
        ))
    })?;
    let channel = |value: i64| {
        u8::try_from(value).map_err(|_| {
            PyValueError::new_err(format!("{parameter} channels must be from 0 to 255"))
        })
    };
    Ok(rpptx::RgbColor::new(
        channel(red)?,
        channel(green)?,
        channel(blue)?,
    ))
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

    fn gradient_fill(&self, py: Python<'_>) -> PyResult<rpptx::GradientFill> {
        current_gradient(self.fill(py)?)
    }
}

/// python-pptx's default gradient: two `accent1` stops on a linear axis.
const DEFAULT_GRADIENT: &str = concat!(
    r#"<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1"><a:gsLst>"#,
    r#"<a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>"#,
    r#"<a:gs pos="100000"><a:schemeClr val="accent1"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>"#,
    r#"</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>"#
);

fn current_gradient(fill: Option<rpptx::Fill>) -> PyResult<rpptx::GradientFill> {
    match fill {
        Some(rpptx::Fill::Gradient(gradient)) => Ok(gradient),
        other => Err(PyTypeError::new_err(format!(
            "fill type {} is not a gradient, call .gradient() first",
            fill_type_name(other.as_ref())
        ))),
    }
}

/// The stops of one gradient fill, like python-pptx `GradientStops`.
#[pyclass(name = "GradientStops")]
pub struct PyGradientStops {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    target: FillTarget,
}

#[pymethods]
impl PyGradientStops {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "gradient stops",
            self.target.suffix(),
        )?;
        Ok(
            current_gradient(current_fill(&presentation.inner, &self.path, self.target)?)?
                .stops
                .len(),
        )
    }

    fn __getitem__(&self, py: Python<'_>, index: isize) -> PyResult<Py<PyGradientStop>> {
        let index = crate::normalize_index(index, self.__len__(py)?, "gradient stop")?;
        Py::new(
            py,
            PyGradientStop {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                target: self.target,
                index,
            },
        )
    }
}

/// One gradient stop, like python-pptx `_GradientStop`.
#[pyclass(name = "GradientStop")]
pub struct PyGradientStop {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    target: FillTarget,
    index: usize,
}

#[pymethods]
impl PyGradientStop {
    /// The stop's place along the gradient, from 0.0 to 1.0.
    #[getter]
    fn position(&self, py: Python<'_>) -> PyResult<f64> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "gradient stop",
            self.target.suffix(),
        )?;
        let gradient =
            current_gradient(current_fill(&presentation.inner, &self.path, self.target)?)?;
        gradient
            .stops
            .get(self.index)
            .map(|stop| f64::from(stop.position.0) / 100_000.0)
            .ok_or_else(|| PyIndexError::new_err("gradient stop index out of range"))
    }

    #[setter]
    fn set_position(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        if !(0.0..=1.0).contains(&value) {
            return Err(PyValueError::new_err(
                "gradient stop position must be from 0.0 to 1.0",
            ));
        }
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "gradient stop",
            self.target.suffix(),
        )?;
        let mut gradient =
            current_gradient(current_fill(&presentation.inner, &self.path, self.target)?)?;
        let stop = gradient
            .stops
            .get_mut(self.index)
            .ok_or_else(|| PyIndexError::new_err("gradient stop index out of range"))?;
        stop.position = rpptx::Percent1000((value * 100_000.0).round() as i32);
        write_fill(
            py,
            &mut presentation.inner,
            &self.path,
            self.target,
            rpptx::Fill::Gradient(gradient),
        )
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Py<PyColorFormat>> {
        Py::new(
            py,
            PyColorFormat {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                source: ColorSource::GradientStop {
                    target: self.target,
                    index: self.index,
                },
            },
        )
    }
}

#[pymethods]
impl PyFillFormat {
    #[getter(r#type)]
    fn fill_type(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let value = match self.fill(py)? {
            None => return Ok(None),
            Some(rpptx::Fill::Solid(_)) => 1,
            Some(rpptx::Fill::Pattern(_)) => 2,
            Some(rpptx::Fill::Gradient(_)) => 3,
            Some(rpptx::Fill::NoFill(_)) => 5,
            Some(rpptx::Fill::Blip(_)) => 6,
        };
        dml_enum(py, "MSO_FILL_TYPE", value).map(Some)
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

    // Gradient and picture fills (#312).

    /// Sets the python-pptx default two-stop `accent1` linear gradient,
    /// keeping a gradient the fill already has.
    fn gradient(&self, py: Python<'_>) -> PyResult<()> {
        if matches!(self.fill(py)?, Some(rpptx::Fill::Gradient(_))) {
            return Ok(());
        }
        let fill = rpptx::Fill::from_xml(DEFAULT_GRADIENT.as_bytes())
            .map_err(|error| PyValueError::new_err(error.to_string()))?;
        self.set(py, fill)
    }

    /// The linear gradient's angle in degrees, counter-clockwise from left
    /// to right, as python-pptx reads it.
    #[getter]
    fn gradient_angle(&self, py: Python<'_>) -> PyResult<f64> {
        let gradient = self.gradient_fill(py)?;
        let Some(rpptx::GradientGeometry::Linear(linear)) = &gradient.geometry else {
            return Err(PyValueError::new_err("not a linear gradient"));
        };
        let clockwise = f64::from(linear.angle.0) / ANGLE_UNITS_PER_DEGREE;
        Ok(if clockwise == 0.0 {
            0.0
        } else {
            360.0 - clockwise
        })
    }

    #[setter]
    fn set_gradient_angle(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        if !value.is_finite() {
            return Err(PyValueError::new_err(
                "gradient angle must be a finite number",
            ));
        }
        let mut gradient = self.gradient_fill(py)?;
        let Some(rpptx::GradientGeometry::Linear(linear)) = &mut gradient.geometry else {
            return Err(PyValueError::new_err("not a linear gradient"));
        };
        let clockwise = (360.0 - value).rem_euclid(360.0);
        linear.angle = rpptx::Angle(
            ((clockwise * ANGLE_UNITS_PER_DEGREE).round() as i64 % ANGLE_UNITS_PER_TURN) as i32,
        );
        self.set(py, rpptx::Fill::Gradient(gradient))
    }

    /// The gradient stops, like python-pptx `GradientStops`.
    #[getter]
    fn gradient_stops(&self, py: Python<'_>) -> PyResult<Py<PyGradientStops>> {
        self.gradient_fill(py)?;
        Py::new(
            py,
            PyGradientStops {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                target: self.target,
            },
        )
    }

    /// Sets a picture fill that stretches an image over a slide, layout or
    /// master background (rpptx extension). The image part and its
    /// relationship are added to that part.
    fn picture(&self, py: Python<'_>, image_file: &Bound<'_, PyAny>) -> PyResult<()> {
        if self.target != FillTarget::Background {
            return Err(PyTypeError::new_err(
                "picture fills are supported on backgrounds only, add a picture shape with shapes.add_picture",
            ));
        }
        self.fill(py)?;
        let (bytes, filename) = image_bytes(image_file)?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .set_part_picture_background(part_of(&self.path)?, &bytes, &filename)
            .map_err(|error| rpptx_to_pyerr(py, error))
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
                source: ColorSource::Fill {
                    target: self.target,
                    solidify: false,
                },
            },
        )
    }
}

/// The colour one colour format reads and writes.
#[derive(Clone, Copy)]
enum ColorSource {
    /// The foreground of a fill. `solidify` makes a set colour a solid fill.
    Fill { target: FillTarget, solidify: bool },
    /// The colour of the outer shadow.
    Shadow,
    /// The colour of one gradient stop.
    GradientStop { target: FillTarget, index: usize },
}

impl ColorSource {
    fn suffix(self) -> &'static str {
        match self {
            Self::Fill { target, .. } | Self::GradientStop { target, .. } => target.suffix(),
            Self::Shadow => ".shadow.color",
        }
    }
}

/// A live view of the foreground colour of one fill, or of a shadow colour.
#[pyclass(name = "ColorFormat")]
pub struct PyColorFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    source: ColorSource,
}

#[pymethods]
impl PyColorFormat {
    #[getter]
    fn rgb(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "color", self.source.suffix())?;
        let color = match self.source {
            ColorSource::Fill { target, .. } => {
                match current_fill(&presentation.inner, &self.path, target)? {
                    Some(rpptx::Fill::Solid(fill)) => fill.color,
                    Some(rpptx::Fill::Pattern(fill)) => fill.foreground,
                    _ => None,
                }
            }
            ColorSource::Shadow => {
                current_shadow(&presentation.inner, &self.path)?.and_then(|shadow| shadow.color)
            }
            ColorSource::GradientStop { target, index } => {
                current_gradient(current_fill(&presentation.inner, &self.path, target)?)?
                    .stops
                    .get(index)
                    .and_then(|stop| stop.color.clone())
            }
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
        validate_path(py, &presentation, &self.path, "color", self.source.suffix())?;
        let (target, solidify) = match self.source {
            ColorSource::Fill { target, solidify } => (target, solidify),
            ColorSource::Shadow => {
                return edit_shadow(py, &mut presentation.inner, &self.path, |shadow| {
                    let color = shadow_rgb(shadow.color.take(), rgb);
                    shadow.replace_color(color);
                });
            }
            ColorSource::GradientStop { target, index } => {
                let mut gradient =
                    current_gradient(current_fill(&presentation.inner, &self.path, target)?)?;
                let stop = gradient
                    .stops
                    .get_mut(index)
                    .ok_or_else(|| PyIndexError::new_err("gradient stop index out of range"))?;
                stop.color = Some(with_rgb(stop.color.take(), rgb));
                return write_fill(
                    py,
                    &mut presentation.inner,
                    &self.path,
                    target,
                    rpptx::Fill::Gradient(gradient),
                );
            }
        };
        let fill = match current_fill(&presentation.inner, &self.path, target)? {
            Some(rpptx::Fill::Solid(mut fill)) => {
                fill.color = Some(with_rgb(fill.color.take(), rgb));
                rpptx::Fill::Solid(fill)
            }
            Some(rpptx::Fill::Pattern(mut fill)) => {
                fill.foreground = Some(with_rgb(fill.foreground.take(), rgb));
                rpptx::Fill::Pattern(fill)
            }
            _ if solidify => {
                let mut fill = rpptx::SolidFill::default();
                fill.color = Some(rpptx::ColorChoice::srgb(rgb));
                rpptx::Fill::Solid(fill)
            }
            other => return Err(no_foreground(other.as_ref())),
        };
        write_fill(py, &mut presentation.inner, &self.path, target, fill)
    }
}

/// A live view of the outline of one shape, picture, or connector, or of one
/// table cell border.
#[pyclass(name = "LineFormat")]
pub struct PyLineFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    target: FillTarget,
}

impl PyLineFormat {
    /// Creates a view of a shape line, with `FillTarget::Line`, or of a cell
    /// border, with `FillTarget::CellBorder`.
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

    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "line",
            self.target.line_suffix(),
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
                source: ColorSource::Fill {
                    target: self.target,
                    solidify: true,
                },
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
                self.target,
            ),
        )
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let width = current_line(&self.presentation.borrow(py).inner, &self.path, self.target)?
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
        let mut line =
            current_line(&presentation.inner, &self.path, self.target)?.unwrap_or_default();
        line.width = Some(width as u32);
        write_line(py, &mut presentation.inner, &self.path, self.target, line)
    }

    /// Returns the preset dash, or `None` without one or with a custom dash.
    #[getter]
    fn dash_style(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let dash = current_line(&self.presentation.borrow(py).inner, &self.path, self.target)?
            .as_ref()
            .and_then(|line| match &line.dash {
                Some(rpptx::LineDash::Preset(dash)) => Some(dash.value),
                _ => None,
            });
        dash.map(|dash| dml_enum(py, "MSO_LINE_DASH_STYLE", dash_value(dash)))
            .transpose()
    }

    /// Writes `a:prstDash`. `None` removes a preset or custom dash.
    #[setter]
    fn set_dash_style(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.validate(py)?;
        let dash = value.map(dash_from_value).transpose()?;
        let mut presentation = self.presentation.borrow_mut(py);
        update_line(
            py,
            &mut presentation.inner,
            &self.path,
            self.target,
            dash.is_some(),
            |line| {
                line.dash = match (dash, line.dash.take()) {
                    (None, _) => None,
                    (Some(value), Some(rpptx::LineDash::Preset(mut preset))) => {
                        preset.value = value;
                        Some(rpptx::LineDash::Preset(preset))
                    }
                    (Some(value), _) => {
                        Some(rpptx::LineDash::Preset(rpptx::PresetDash::new(value)))
                    }
                };
            },
        )
    }

    /// Returns the `a:headEnd`, the end at the first point of the line.
    #[getter]
    fn head_end(&self, py: Python<'_>) -> PyResult<Py<PyLineEndFormat>> {
        self.line_end(py, false)
    }

    /// Returns the `a:tailEnd`, the end at the last point of the line.
    #[getter]
    fn tail_end(&self, py: Python<'_>) -> PyResult<Py<PyLineEndFormat>> {
        self.line_end(py, true)
    }
}

impl PyLineFormat {
    fn line_end(&self, py: Python<'_>, tail: bool) -> PyResult<Py<PyLineEndFormat>> {
        self.validate(py)?;
        Py::new(
            py,
            PyLineEndFormat {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                target: self.target,
                tail,
            },
        )
    }
}

/// A live view of one end of a line, its `a:headEnd` or `a:tailEnd`.
///
/// Each property is `None` when its attribute is absent. Assigning `None`
/// removes the attribute, and an end left with no attribute is removed. An
/// `a:ln` left empty is kept, as python-pptx keeps it.
#[pyclass(name = "LineEndFormat")]
pub struct PyLineEndFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    target: FillTarget,
    tail: bool,
}

impl PyLineEndFormat {
    fn end(&self, py: Python<'_>) -> PyResult<Option<rpptx::LineEnd>> {
        let presentation = self.presentation.borrow(py);
        let suffix = if self.tail {
            ".line.tail_end"
        } else {
            ".line.head_end"
        };
        validate_path(py, &presentation, &self.path, "line end", suffix)?;
        let line = current_line(&presentation.inner, &self.path, self.target)?;
        Ok(line.as_ref().and_then(|line| {
            if self.tail {
                line.tail_end.clone()
            } else {
                line.head_end.clone()
            }
        }))
    }

    fn update(
        &self,
        py: Python<'_>,
        create: bool,
        change: impl FnOnce(&mut rpptx::LineEnd),
    ) -> PyResult<()> {
        self.end(py)?;
        let tail = self.tail;
        let mut presentation = self.presentation.borrow_mut(py);
        update_line(
            py,
            &mut presentation.inner,
            &self.path,
            self.target,
            create,
            |line| {
                let slot = if tail {
                    &mut line.tail_end
                } else {
                    &mut line.head_end
                };
                if slot.is_none() && !create {
                    return;
                }
                let mut end = slot.take().unwrap_or_default();
                change(&mut end);
                *slot = (end != rpptx::LineEnd::default()).then_some(end);
            },
        )
    }
}

#[pymethods]
impl PyLineEndFormat {
    #[getter(r#type)]
    fn end_type(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.end(py)?
            .and_then(|end| end.kind)
            .map(|kind| dml_enum(py, "MSO_ARROWHEAD_STYLE", end_type_value(kind)))
            .transpose()
    }

    /// Writes `type`. `NONE` writes `type="none"`, as PowerPoint scripting does.
    #[setter(r#type)]
    fn set_end_type(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        let kind = value.map(end_type_from_value).transpose()?;
        self.update(py, kind.is_some(), |end| end.kind = kind)
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.end(py)?
            .and_then(|end| end.width)
            .map(|size| dml_enum(py, "MSO_ARROWHEAD_WIDTH", end_size_value(size)))
            .transpose()
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        let size = value
            .map(|value| end_size_from_value(value, "MSO_ARROWHEAD_WIDTH"))
            .transpose()?;
        self.update(py, size.is_some(), |end| end.width = size)
    }

    #[getter]
    fn length(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.end(py)?
            .and_then(|end| end.length)
            .map(|size| dml_enum(py, "MSO_ARROWHEAD_LENGTH", end_size_value(size)))
            .transpose()
    }

    #[setter]
    fn set_length(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        let size = value
            .map(|value| end_size_from_value(value, "MSO_ARROWHEAD_LENGTH"))
            .transpose()?;
        self.update(py, size.is_some(), |end| end.length = size)
    }
}

const DEFAULT_SHADOW_BLUR_RADIUS: i64 = 50_800;
const DEFAULT_SHADOW_DISTANCE: i64 = 38_100;
const DEFAULT_SHADOW_DIRECTION: i32 = 2_700_000;
const DEFAULT_SHADOW_ALPHA: i32 = 40_000;
const OPAQUE_ALPHA: f64 = 100_000.0;
const RECT_ALIGNMENTS: &str = "tl, t, tr, l, ctr, r, bl, b, br";

fn current_effects(
    presentation: &rpptx::Presentation,
    path: &ContentPath,
) -> PyResult<Option<rpptx::CT_EffectList>> {
    Ok(shape_ref_at(presentation, path)
        .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
        .effects()
        .cloned())
}

fn current_shadow(
    presentation: &rpptx::Presentation,
    path: &ContentPath,
) -> PyResult<Option<rpptx::CT_OuterShadowEffect>> {
    Ok(current_effects(presentation, path)?.and_then(|effects| effects.outer_shadow))
}

fn write_effects(
    py: Python<'_>,
    presentation: &mut rpptx::Presentation,
    path: &ContentPath,
    effects: Option<rpptx::CT_EffectList>,
) -> PyResult<()> {
    shape_mut_at(presentation, path)
        .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
        .set_effects(effects)
        .map_err(|error| rpptx_to_pyerr(py, error))
}

/// The outer shadow PowerPoint for Mac writes for its Offset Diagonal
/// Bottom Right preset, `msoShadow21`: preset black at 40% opacity, with a
/// 4 pt blur, 3 pt away at 45 degrees.
fn default_outer_shadow() -> rpptx::CT_OuterShadowEffect {
    let mut shadow = rpptx::CT_OuterShadowEffect::default();
    shadow.blur_radius = Some(DEFAULT_SHADOW_BLUR_RADIUS);
    shadow.distance = Some(DEFAULT_SHADOW_DISTANCE);
    shadow.direction = Some(rpptx::Angle(DEFAULT_SHADOW_DIRECTION));
    shadow.alignment = Some(rpptx::RectAlignment::TopLeft);
    shadow.rotate_with_shape = Some(false);
    shadow.color = Some(rpptx::ColorChoice::Preset {
        value: "black".to_owned(),
        transforms: vec![rpptx::ColorTransform::Alpha(rpptx::Percent1000(
            DEFAULT_SHADOW_ALPHA,
        ))],
        raw_children: Default::default(),
    });
    shadow
}

/// Changes the outer shadow, first adding PowerPoint's default one when the
/// shape has none. Sibling effects in the list are kept.
fn edit_shadow(
    py: Python<'_>,
    presentation: &mut rpptx::Presentation,
    path: &ContentPath,
    change: impl FnOnce(&mut rpptx::CT_OuterShadowEffect),
) -> PyResult<()> {
    let mut effects = current_effects(presentation, path)?.unwrap_or_default();
    change(
        effects
            .outer_shadow
            .get_or_insert_with(default_outer_shadow),
    );
    write_effects(py, presentation, path, Some(effects))
}

/// Gives a shadow colour a new sRGB value. The opacity survives a change of
/// colour kind, which a theme or preset colour's other transforms do not.
fn shadow_rgb(color: Option<rpptx::ColorChoice>, rgb: rpptx::RgbColor) -> rpptx::ColorChoice {
    if let Some(color @ rpptx::ColorChoice::Srgb { .. }) = color {
        return with_rgb(Some(color), rgb);
    }
    rpptx::ColorChoice::Srgb {
        value: rgb,
        transforms: color
            .iter()
            .flat_map(rpptx::ColorChoice::transforms)
            .filter(|transform| matches!(transform, rpptx::ColorTransform::Alpha(_)))
            .copied()
            .collect(),
        raw_children: Default::default(),
    }
}

fn color_transforms_mut(color: &mut rpptx::ColorChoice) -> &mut Vec<rpptx::ColorTransform> {
    match color {
        rpptx::ColorChoice::Srgb { transforms, .. }
        | rpptx::ColorChoice::Scheme { transforms, .. }
        | rpptx::ColorChoice::System { transforms, .. }
        | rpptx::ColorChoice::Preset { transforms, .. } => transforms,
    }
}

fn check_shadow_length(name: &str, value: i64) -> PyResult<i64> {
    if (0..=MAX_COORDINATE).contains(&value) {
        Ok(value)
    } else {
        Err(PyValueError::new_err(format!(
            "shadow {name} must be between 0 and {MAX_COORDINATE} EMU, got {value}"
        )))
    }
}

/// A live view of the shadow of one shape, picture, connector, or group.
///
/// `inherit` is python-pptx's: whether the shape has no `a:effectLst` of
/// its own and so takes its theme effect. The other properties read and
/// write the shape's own `a:outerShdw`. They read `None` when the shape has
/// none, and a property the element omits reads as its schema default.
/// Setting one on a shape without an outer shadow first adds PowerPoint's
/// Offset Diagonal Bottom Right shadow, preset black at 40% opacity with a
/// 4 pt blur, 3 pt away at 45 degrees.
#[pyclass(name = "ShadowFormat")]
pub struct PyShadowFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyShadowFormat {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn read<T>(
        &self,
        py: Python<'_>,
        read: impl FnOnce(Option<rpptx::CT_EffectList>) -> T,
    ) -> PyResult<T> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "shadow", ".shadow")?;
        Ok(read(current_effects(&presentation.inner, &self.path)?))
    }

    fn read_shadow<T>(
        &self,
        py: Python<'_>,
        read: impl FnOnce(&rpptx::CT_OuterShadowEffect) -> T,
    ) -> PyResult<Option<T>> {
        self.read(py, |effects| {
            effects.and_then(|effects| effects.outer_shadow.as_ref().map(read))
        })
    }

    fn write(&self, py: Python<'_>, effects: Option<rpptx::CT_EffectList>) -> PyResult<()> {
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(py, &presentation, &self.path, "shadow", ".shadow")?;
        write_effects(py, &mut presentation.inner, &self.path, effects)
    }

    fn edit(
        &self,
        py: Python<'_>,
        change: impl FnOnce(&mut rpptx::CT_OuterShadowEffect),
    ) -> PyResult<()> {
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(py, &presentation, &self.path, "shadow", ".shadow")?;
        edit_shadow(py, &mut presentation.inner, &self.path, change)
    }
}

#[pymethods]
impl PyShadowFormat {
    /// Whether the shape takes its effects from the theme.
    ///
    /// Setting `True` removes the shape's `a:effectLst`, and with it every
    /// effect of its own. Setting `False` adds an empty one when absent, so
    /// no effect shows, as python-pptx does.
    #[getter]
    fn inherit(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |effects| effects.is_none())
    }

    #[setter]
    fn set_inherit(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        let effects = self.read(py, |effects| effects)?;
        match (value, effects) {
            (true, None) | (false, Some(_)) => Ok(()),
            (true, Some(_)) => self.write(py, None),
            (false, None) => self.write(py, Some(rpptx::CT_EffectList::default())),
        }
    }

    /// Whether the shape has an outer shadow of its own.
    ///
    /// A theme shadow reached through `inherit` is not reported. Setting
    /// `False` removes the outer shadow and keeps the effect list, so the
    /// theme shadow does not return either. On a shape without an outer
    /// shadow of its own it changes nothing.
    #[getter]
    fn visible(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.read_shadow(py, |_| ())?.is_some())
    }

    #[setter]
    fn set_visible(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        if value {
            return self.edit(py, |_| {});
        }
        let Some(mut effects) = self.read(py, |effects| effects)? else {
            return Ok(());
        };
        if effects.outer_shadow.take().is_none() {
            return Ok(());
        }
        self.write(py, Some(effects))
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Py<PyColorFormat>> {
        self.read(py, |_| ())?;
        Py::new(
            py,
            PyColorFormat {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                source: ColorSource::Shadow,
            },
        )
    }

    /// The shadow opacity from 0.0, transparent, to 1.0, opaque.
    ///
    /// A colour rpptx does not model, `a:scrgbClr` or `a:hslClr`, reads
    /// `None` and refuses a new opacity until `color.rgb` replaces it.
    #[getter]
    fn alpha(&self, py: Python<'_>) -> PyResult<Option<f64>> {
        let alpha = self.read_shadow(py, |shadow| {
            if shadow.has_unmodelled_color() {
                return None;
            }
            Some(
                shadow
                    .color
                    .iter()
                    .flat_map(rpptx::ColorChoice::transforms)
                    .rev()
                    .find_map(|transform| match transform {
                        rpptx::ColorTransform::Alpha(value) => {
                            Some(f64::from(value.0) / OPAQUE_ALPHA)
                        }
                        _ => None,
                    })
                    .unwrap_or(1.0),
            )
        })?;
        Ok(alpha.flatten())
    }

    #[setter]
    fn set_alpha(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        if !(0.0..=1.0).contains(&value) {
            return Err(PyValueError::new_err(format!(
                "shadow alpha must be between 0.0 and 1.0, got {value}"
            )));
        }
        if self.read_shadow(py, rpptx::CT_OuterShadowEffect::has_unmodelled_color)? == Some(true) {
            return Err(PyValueError::new_err(
                "shadow colour is an a:scrgbClr or a:hslClr, which rpptx does not model, \
                 set color.rgb first",
            ));
        }
        let alpha = (value * OPAQUE_ALPHA).round_ties_even() as i32;
        self.edit(py, |shadow| {
            let color = shadow
                .color
                .get_or_insert_with(|| rpptx::ColorChoice::srgb(rpptx::RgbColor::new(0, 0, 0)));
            let transforms = color_transforms_mut(color);
            transforms.retain(|transform| !matches!(transform, rpptx::ColorTransform::Alpha(_)));
            if alpha < OPAQUE_ALPHA as i32 {
                transforms.push(rpptx::ColorTransform::Alpha(rpptx::Percent1000(alpha)));
            }
        })
    }

    #[getter]
    fn blur_radius(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let blur = self.read_shadow(py, |shadow| shadow.blur_radius.unwrap_or(0))?;
        length(py, blur.map(rpptx::Emu))
    }

    #[setter]
    fn set_blur_radius(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        let value = check_shadow_length("blur_radius", value)?;
        self.edit(py, |shadow| shadow.blur_radius = Some(value))
    }

    #[getter]
    fn distance(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let distance = self.read_shadow(py, |shadow| shadow.distance.unwrap_or(0))?;
        length(py, distance.map(rpptx::Emu))
    }

    #[setter]
    fn set_distance(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        let value = check_shadow_length("distance", value)?;
        self.edit(py, |shadow| shadow.distance = Some(value))
    }

    /// The direction the shadow is cast in, in degrees clockwise from the
    /// positive x axis, normalized to `0.0 <= direction < 360.0`.
    #[getter]
    fn direction(&self, py: Python<'_>) -> PyResult<Option<f64>> {
        self.read_shadow(py, |shadow| {
            f64::from(shadow.direction.unwrap_or_default().0) / ANGLE_UNITS_PER_DEGREE
        })
    }

    #[setter]
    fn set_direction(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        if !value.is_finite() {
            return Err(PyValueError::new_err(
                "shadow direction must be a finite number of degrees",
            ));
        }
        let units = ((value * ANGLE_UNITS_PER_DEGREE).round_ties_even() as i64)
            .rem_euclid(ANGLE_UNITS_PER_TURN);
        let angle = rpptx::Angle(i32::try_from(units).expect("a normalized angle fits in i32"));
        self.edit(py, |shadow| shadow.direction = Some(angle))
    }

    /// The `ST_RectAlignment` token the shadow is anchored at: `tl`, `t`,
    /// `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br`.
    #[getter]
    fn align(&self, py: Python<'_>) -> PyResult<Option<&'static str>> {
        self.read_shadow(py, |shadow| {
            shadow
                .alignment
                .unwrap_or(rpptx::RectAlignment::Bottom)
                .as_str()
        })
    }

    #[setter]
    fn set_align(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let alignment = rpptx::RectAlignment::parse(value).ok_or_else(|| {
            PyValueError::new_err(format!(
                "shadow align must be one of {RECT_ALIGNMENTS}, got {value:?}"
            ))
        })?;
        self.edit(py, |shadow| shadow.alignment = Some(alignment))
    }

    #[getter]
    fn rotate_with_shape(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.read_shadow(py, |shadow| shadow.rotate_with_shape.unwrap_or(true))
    }

    #[setter]
    fn set_rotate_with_shape(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.edit(py, |shadow| shadow.rotate_with_shape = Some(value))
    }
}
