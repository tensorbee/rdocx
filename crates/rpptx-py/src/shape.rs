use std::path::PathBuf;

use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyNotImplementedError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyByteArray, PyBytes, PyIterator, PyList, PySlice, PyString, PyTuple};
use smallvec::smallvec;

use crate::dml::{FillTarget, PyFillFormat, PyLineFormat, PyShadowFormat};
use crate::normalize_index;
use crate::presentation::PyPresentation;
use crate::rpptx_to_pyerr;
use crate::slide::PySlide;
use crate::table::PyTable;
use crate::text::PyTextFrame;
use crate::validate_path;

const MIN_COORDINATE: i64 = -27_273_042_329_600;
pub(crate) const MAX_COORDINATE: i64 = 27_273_042_316_900;
pub(crate) const ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
pub(crate) const ANGLE_UNITS_PER_TURN: i64 = 21_600_000;
/// `a:srcRect` stores a crop inset in thousandths of a percent.
const CROP_UNITS_PER_FRACTION: f64 = 100_000.0;

/// Left, top, right, and bottom picture crop insets, as the facade reports them.
type Crop = (
    rpptx::Percent1000,
    rpptx::Percent1000,
    rpptx::Percent1000,
    rpptx::Percent1000,
);

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyShape>()?;
    module.add_class::<PyShapeClickAction>()?;
    module.add_class::<PyShapeHyperlink>()?;
    module.add_class::<PyShapeCollection>()?;
    module.add_class::<PyPlaceholderCollection>()?;
    module.add_class::<PyImage>()?;
    module.add_class::<PyAdjustmentCollection>()?;
    module.add_class::<PyPlaceholderFormat>()?;
    Ok(())
}

pub(crate) fn length(py: Python<'_>, value: Option<rpptx::Emu>) -> PyResult<Option<Py<PyAny>>> {
    value
        .map(|value| {
            py.import("rpptx")?
                .getattr("Length")?
                .call1((value.0,))
                .map(Bound::unbind)
        })
        .transpose()
}

/// Reads the bytes of a path, a bytes-like object, or a binary file-like object.
///
/// A file-like object is rewound first when it can seek, as python-pptx does.
pub(crate) fn image_bytes(image_file: &Bound<'_, PyAny>) -> PyResult<(Vec<u8>, String)> {
    fn blob(data: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
        match data.cast::<PyBytes>() {
            Ok(bytes) => Ok(bytes.as_bytes().to_vec()),
            Err(_) => data.extract::<Vec<u8>>(),
        }
    }
    if image_file.is_instance_of::<PyBytes>() || image_file.is_instance_of::<PyByteArray>() {
        return Ok((blob(image_file)?, "image".to_owned()));
    }
    if image_file.hasattr("read")? {
        if image_file.hasattr("seek")? {
            image_file.call_method1("seek", (0,))?;
        }
        return Ok((blob(&image_file.call_method0("read")?)?, "image".to_owned()));
    }
    let path = image_file.extract::<PathBuf>()?;
    let bytes = std::fs::read(&path)?;
    let filename = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("image")
        .to_owned();
    Ok((bytes, filename))
}

/// Reads an `MSO_SHAPE` member, its integer value, or a DrawingML preset name.
fn preset_name(py: Python<'_>, shape_type: &Bound<'_, PyAny>) -> PyResult<String> {
    if shape_type.is_instance_of::<PyString>() {
        return shape_type.extract::<String>();
    }
    let value = shape_type.extract::<i64>()?;
    py.import("rpptx.enum.shapes")?
        .getattr("MSO_SHAPE")?
        .call1((value,))
        .map_err(|_| PyValueError::new_err("unsupported MSO_SHAPE value"))?
        .getattr("xml_value")?
        .extract::<String>()
}

fn check_coordinate(name: &str, value: i64, minimum: i64) -> PyResult<()> {
    if (minimum..=MAX_COORDINATE).contains(&value) {
        Ok(())
    } else {
        Err(PyValueError::new_err(format!(
            "{name} must be between {minimum} and {MAX_COORDINATE} EMU, got {value}"
        )))
    }
}

/// The slide, layout or master a content path starts in.
pub(crate) fn part_of(path: &ContentPath) -> PyResult<rpptx::PartRef> {
    path.segs
        .iter()
        .find_map(|segment| match segment {
            PathSeg::Slide(index) => Some(rpptx::PartRef::Slide(*index)),
            PathSeg::Layout(index) => Some(rpptx::PartRef::Layout(*index)),
            PathSeg::Master(index) => Some(rpptx::PartRef::Master(*index)),
            _ => None,
        })
        .ok_or_else(|| PyIndexError::new_err("slide index is missing"))
}

/// The slide a content path starts in. Operations that only slides
/// support, such as `click_action.target_slide`, replacing or cropping a
/// picture's image, effective geometry and media, raise `ValueError` on
/// layout and master shapes. `click_action.hyperlink` works on all three.
pub(crate) fn slide_index(path: &ContentPath) -> PyResult<usize> {
    match part_of(path)? {
        rpptx::PartRef::Slide(index) => Ok(index),
        _ => Err(PyValueError::new_err(
            "this works on slide shapes only, not on slide layout or slide master shapes",
        )),
    }
}

fn shape_indices(path: &ContentPath) -> impl Iterator<Item = usize> + '_ {
    path.segs.iter().filter_map(|segment| match segment {
        PathSeg::Shape(index) => Some(*index),
        _ => None,
    })
}

pub(crate) fn shape_ref_at<'a>(
    presentation: &'a rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeRef<'a>> {
    let part = part_of(path).ok()?;
    let mut indices = shape_indices(path);
    let mut shape = presentation.part_shape(part, indices.next()?)?;
    for index in indices {
        shape = shape.child(index)?;
    }
    Some(shape)
}

pub(crate) fn shape_mut_at<'a>(
    presentation: &'a mut rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeMut<'a>> {
    let part = part_of(path).ok()?;
    let mut indices = shape_indices(path);
    let mut shape = presentation.part_shape_mut(part, indices.next()?)?;
    for index in indices {
        shape = shape.into_child_mut(index)?;
    }
    Some(shape)
}

#[pyclass(name = "Shape")]
pub struct PyShape {
    pub(crate) presentation: Py<PyPresentation>,
    pub(crate) path: ContentPath,
}

impl PyShape {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "shape", "")
    }

    fn edit(
        &self,
        py: Python<'_>,
        change: impl FnOnce(&mut rpptx::ShapeMut<'_>) -> rpptx::Result<()>,
    ) -> PyResult<()> {
        self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut shape = shape_mut_at(&mut presentation.inner, &self.path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?;
        change(&mut shape).map_err(|error| rpptx_to_pyerr(py, error))
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rpptx::ShapeRef<'_>) -> T) -> PyResult<T> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        shape_ref_at(&presentation.inner, &self.path)
            .map(read)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))
    }

    fn require_shape_properties(&self, py: Python<'_>, property: &str) -> PyResult<()> {
        let kind = self.read(py, |shape| shape.kind())?;
        if matches!(
            kind,
            rpptx::ShapeKind::Shape | rpptx::ShapeKind::Picture | rpptx::ShapeKind::Connector
        ) {
            Ok(())
        } else {
            Err(PyValueError::new_err(format!("shape has no {property}")))
        }
    }

    fn crop(&self, py: Python<'_>) -> PyResult<Crop> {
        self.read(py, |shape| shape.crop())?
            .ok_or_else(|| PyValueError::new_err("shape is not a picture"))
    }

    fn crop_edge(
        &self,
        py: Python<'_>,
        edge: fn(&mut Crop) -> &mut rpptx::Percent1000,
    ) -> PyResult<f64> {
        let mut crop = self.crop(py)?;
        Ok(f64::from(edge(&mut crop).0) / CROP_UNITS_PER_FRACTION)
    }

    /// Writes one crop inset as python-pptx does, rounding half to even, and
    /// leaves the picture unchanged when the value equals the stored one.
    fn set_crop_edge(
        &self,
        py: Python<'_>,
        edge: fn(&mut Crop) -> &mut rpptx::Percent1000,
        value: f64,
    ) -> PyResult<()> {
        let units = (value * CROP_UNITS_PER_FRACTION).round_ties_even();
        if !(f64::from(i32::MIN)..=f64::from(i32::MAX)).contains(&units) {
            return Err(PyValueError::new_err(format!(
                "crop must be a finite fraction from -21474.83648 to 21474.83647, got {value}"
            )));
        }
        let current = self.crop(py)?;
        let mut crop = current;
        *edge(&mut crop) = rpptx::Percent1000(units as i32);
        if crop == current {
            return Ok(());
        }
        let (left, top, right, bottom) = crop;
        self.edit(py, |shape| shape.set_crop(left, top, right, bottom))
    }

    /// Copies a placeholder's inherited transform onto it before a geometry or rotation edit.
    fn materialize_geometry(&self, py: Python<'_>) -> PyResult<()> {
        self.validate(py)?;
        let shape_path = shape_indices(&self.path).collect::<Vec<_>>();
        self.presentation
            .borrow_mut(py)
            .inner
            .materialize_geometry(slide_index(&self.path)?, &shape_path)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    fn picture_id(&self, py: Python<'_>) -> PyResult<(usize, u32)> {
        let (kind, id) = self.read(py, |shape| (shape.kind(), shape.non_visual_id()))?;
        match (kind, id) {
            (rpptx::ShapeKind::Picture, Some(id)) => Ok((slide_index(&self.path)?, id)),
            _ => Err(PyValueError::new_err("shape is not a picture")),
        }
    }
}

#[pymethods]
impl PyShape {
    #[getter]
    fn click_action(&self, py: Python<'_>) -> PyResult<Py<PyShapeClickAction>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeClickAction {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    #[getter]
    fn left(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.position())
            .map(|(left, _)| left);
        drop(presentation);
        length(py, value)
    }

    #[getter]
    fn top(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.position())
            .map(|(_, top)| top);
        drop(presentation);
        length(py, value)
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.size())
            .map(|(width, _)| width);
        drop(presentation);
        length(py, value)
    }

    #[getter]
    fn height(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.size())
            .map(|(_, height)| height);
        drop(presentation);
        length(py, value)
    }

    /// Left, top, width and height as rendering places the shape.
    ///
    /// A placeholder without its own transform reports the one it inherits
    /// from its layout and master, where `left` and the other properties
    /// report `None`.
    fn effective_geometry<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        self.validate(py)?;
        let shape_path = shape_indices(&self.path).collect::<Vec<_>>();
        let geometry = self
            .presentation
            .borrow(py)
            .inner
            .effective_geometry(slide_index(&self.path)?, &shape_path)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let Some((left, top, width, height)) = geometry else {
            return Ok(None);
        };
        let length = py.import("rpptx")?.getattr("Length")?;
        let values = [left, top, width, height]
            .into_iter()
            .map(|value| length.call1((value.0,)))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, values).map(Some)
    }

    #[getter]
    fn shape_id(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        self.validate(py)?;
        Ok(
            shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
                .and_then(|shape| shape.non_visual_id()),
        )
    }

    #[setter]
    fn set_left(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("left", value, MIN_COORDINATE)?;
        self.materialize_geometry(py)?;
        let top = self.read(py, |shape| {
            shape.position().map_or(rpptx::Emu(0), |(_, top)| top)
        })?;
        self.edit(py, |shape| shape.set_position(rpptx::Emu(value), top))
    }

    #[setter]
    fn set_top(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("top", value, MIN_COORDINATE)?;
        self.materialize_geometry(py)?;
        let left = self.read(py, |shape| {
            shape.position().map_or(rpptx::Emu(0), |(left, _)| left)
        })?;
        self.edit(py, |shape| shape.set_position(left, rpptx::Emu(value)))
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("width", value, 0)?;
        self.materialize_geometry(py)?;
        let height = self.read(py, |shape| {
            shape.size().map_or(rpptx::Emu(0), |(_, height)| height)
        })?;
        self.edit(py, |shape| shape.set_size(rpptx::Emu(value), height))
    }

    #[setter]
    fn set_height(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("height", value, 0)?;
        self.materialize_geometry(py)?;
        let width = self.read(py, |shape| {
            shape.size().map_or(rpptx::Emu(0), |(width, _)| width)
        })?;
        self.edit(py, |shape| shape.set_size(width, rpptx::Emu(value)))
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.validate(py)?;
        Ok(
            shape_ref_at(&self.presentation.borrow(py).inner, &self.path)
                .and_then(|shape| shape.non_visual_name()),
        )
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.edit(py, |shape| shape.set_name(value))
    }

    /// Clockwise rotation in degrees, normalized to `0.0 <= rotation < 360.0`.
    #[getter]
    fn rotation(&self, py: Python<'_>) -> PyResult<f64> {
        let rotation = self.read(py, |shape| shape.rotation())?;
        Ok(rotation.map_or(0.0, |angle| {
            i64::from(angle.0).rem_euclid(ANGLE_UNITS_PER_TURN) as f64 / ANGLE_UNITS_PER_DEGREE
        }))
    }

    #[setter]
    fn set_rotation(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        if !value.is_finite() {
            return Err(PyValueError::new_err(
                "rotation must be a finite number of degrees",
            ));
        }
        let units = ((value * ANGLE_UNITS_PER_DEGREE).round_ties_even() as i64)
            .rem_euclid(ANGLE_UNITS_PER_TURN);
        let angle = rpptx::Angle(i32::try_from(units).expect("a normalized angle fits in i32"));
        self.materialize_geometry(py)?;
        self.edit(py, |shape| shape.set_rotation(angle))
    }

    #[getter]
    fn shape_type(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let Some(shape_type) = self.read(py, |shape| shape.shape_type())? else {
            return Ok(None);
        };
        let value = match shape_type {
            rpptx::ShapeType::AutoShape => 1,
            rpptx::ShapeType::Chart => 3,
            rpptx::ShapeType::Freeform => 5,
            rpptx::ShapeType::Group => 6,
            rpptx::ShapeType::EmbeddedOleObject => 7,
            rpptx::ShapeType::Line => 9,
            rpptx::ShapeType::LinkedOleObject => 10,
            rpptx::ShapeType::Picture => 13,
            rpptx::ShapeType::Placeholder => 14,
            rpptx::ShapeType::Media => 16,
            rpptx::ShapeType::TextBox => 17,
            rpptx::ShapeType::Table => 19,
        };
        py.import("rpptx.enum.shapes")?
            .getattr("MSO_SHAPE_TYPE")?
            .call1((value,))
            .map(|member| Some(member.unbind()))
    }

    /// The `MSO_SHAPE` member of an auto shape's preset geometry.
    ///
    /// A picture reports its mask, or `None` without a preset. Any other
    /// shape raises `ValueError`, as python-pptx does.
    #[getter]
    fn auto_shape_type(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let (kind, preset) = self.read(py, |shape| {
            (shape.kind(), shape.auto_shape_type().map(str::to_owned))
        })?;
        match (kind, preset) {
            (_, Some(preset)) => py
                .import("rpptx.enum.shapes")?
                .getattr("MSO_SHAPE")?
                .call_method1("from_xml", (preset,))
                .map(|member| Some(member.unbind())),
            (rpptx::ShapeKind::Picture, None) => Ok(None),
            _ => Err(PyValueError::new_err("shape is not an auto shape")),
        }
    }

    /// Replaces the preset geometry and resets its adjustments to defaults.
    #[setter]
    fn set_auto_shape_type(&self, py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let refused = self.read(py, |shape| match shape.kind() {
            rpptx::ShapeKind::Shape => shape.shape_type() == Some(rpptx::ShapeType::TextBox),
            rpptx::ShapeKind::Picture => false,
            _ => true,
        })?;
        if refused {
            return Err(PyValueError::new_err("shape is not an auto shape"));
        }
        let preset = preset_name(py, value)?;
        self.edit(py, |shape| shape.set_auto_shape_type(&preset))
    }

    #[getter]
    fn adjustments(&self, py: Python<'_>) -> PyResult<Py<PyAdjustmentCollection>> {
        if self.read(py, |shape| shape.kind())? != rpptx::ShapeKind::Shape {
            return Err(PyValueError::new_err("shape has no adjustments"));
        }
        Py::new(
            py,
            PyAdjustmentCollection {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    #[getter]
    fn fill(&self, py: Python<'_>) -> PyResult<Py<PyFillFormat>> {
        self.require_shape_properties(py, "fill")?;
        Py::new(
            py,
            PyFillFormat::new(
                self.presentation.clone_ref(py),
                self.path.clone(),
                FillTarget::Shape,
            ),
        )
    }

    #[getter]
    fn line(&self, py: Python<'_>) -> PyResult<Py<PyLineFormat>> {
        self.require_shape_properties(py, "line")?;
        Py::new(
            py,
            PyLineFormat::new(
                self.presentation.clone_ref(py),
                self.path.clone(),
                FillTarget::Line,
            ),
        )
    }

    /// The shadow of a shape, picture, connector, or group.
    ///
    /// A graphic frame raises `NotImplementedError`, as in python-pptx.
    #[getter]
    fn shadow(&self, py: Python<'_>) -> PyResult<Py<PyShadowFormat>> {
        match self.read(py, |shape| shape.kind())? {
            rpptx::ShapeKind::Shape
            | rpptx::ShapeKind::Picture
            | rpptx::ShapeKind::Connector
            | rpptx::ShapeKind::Group => {}
            rpptx::ShapeKind::GraphicFrame => {
                return Err(PyNotImplementedError::new_err(
                    "shadow property on GraphicFrame not yet supported",
                ));
            }
            rpptx::ShapeKind::AlternateContent => {
                return Err(PyValueError::new_err("shape has no shadow"));
            }
        }
        Py::new(
            py,
            PyShadowFormat::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    /// The theme effect style index the shape's `p:style` references, or
    /// `None` without one. Writing 0 removes the theme's effect, such as the
    /// shadow of a connector from `add_connector`. The setter takes an `int`
    /// only: `None` is refused, because dropping `p:style` would leave a
    /// connector without a direct line invisible.
    #[getter]
    fn theme_effect_index(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        self.read(py, |shape| shape.theme_effect_index())
    }

    #[setter]
    fn set_theme_effect_index(&self, py: Python<'_>, value: u32) -> PyResult<()> {
        self.edit(py, |shape| shape.set_theme_effect_index(value))
    }

    /// The shape element serialized on its own, as bytes.
    #[getter]
    fn xml<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let xml = self
            .read(py, |shape| shape.xml())?
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        Ok(PyBytes::new(py, &xml))
    }

    #[getter]
    fn image(&self, py: Python<'_>) -> PyResult<PyImage> {
        let (slide, shape_id) = self.picture_id(py)?;
        let presentation = self.presentation.borrow(py);
        let image = presentation
            .inner
            .picture_image(slide, shape_id)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        Ok(PyImage {
            ext: image_extension(&image.content_type, &image.part_name),
            content_type: image.content_type,
            blob: image.bytes.to_vec(),
        })
    }

    /// The share of the image cropped from the picture's left edge.
    #[getter]
    fn crop_left(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.0)
    }

    #[setter]
    fn set_crop_left(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.0, value)
    }

    /// The share of the image cropped from the picture's top edge.
    #[getter]
    fn crop_top(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.1)
    }

    #[setter]
    fn set_crop_top(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.1, value)
    }

    /// The share of the image cropped from the picture's right edge.
    #[getter]
    fn crop_right(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.2)
    }

    #[setter]
    fn set_crop_right(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.2, value)
    }

    /// The share of the image cropped from the picture's bottom edge.
    #[getter]
    fn crop_bottom(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.3)
    }

    #[setter]
    fn set_crop_bottom(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.3, value)
    }

    /// Replaces the picture's image, keeping its position, size, and crop.
    fn replace_image(&self, py: Python<'_>, image_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let (slide, shape_id) = self.picture_id(py)?;
        let (bytes, _) = image_bytes(image_file)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .replace_picture_image(slide, shape_id, &bytes)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<Py<PyShapeCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.text())
            .ok_or_else(|| PyValueError::new_err("shape has no text"))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        shape_mut_at(&mut presentation.inner, &self.path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .set_text(text)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    #[getter]
    fn has_text_frame(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        Ok(shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.text_frame())
            .is_some())
    }

    #[getter]
    fn text_frame(&self, py: Python<'_>) -> PyResult<Py<PyTextFrame>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        if shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.text_frame())
            .is_none()
        {
            return Err(PyValueError::new_err("shape has no text frame"));
        }
        drop(presentation);
        Py::new(
            py,
            PyTextFrame::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    #[getter]
    fn has_table(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        Ok(shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.table())
            .is_some())
    }

    #[getter]
    fn table(&self, py: Python<'_>) -> PyResult<Py<PyTable>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        if shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.table())
            .is_none()
        {
            return Err(PyValueError::new_err("shape has no table"));
        }
        drop(presentation);
        Py::new(
            py,
            PyTable::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    // Placeholders (#311).

    /// True when the shape is a placeholder, a `p:ph` in its properties.
    #[getter]
    fn is_placeholder(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |shape| shape.placeholder_idx().is_some())
    }

    /// The placeholder's type and index, like python-pptx `placeholder_format`.
    ///
    /// Raises `ValueError` when the shape is not a placeholder.
    #[getter]
    fn placeholder_format(&self, py: Python<'_>) -> PyResult<PyPlaceholderFormat> {
        let (idx, token) = self.read(py, |shape| {
            (
                shape.placeholder_idx(),
                shape.placeholder_type().map(str::to_owned),
            )
        })?;
        let idx = idx.ok_or_else(|| PyValueError::new_err("shape is not a placeholder"))?;
        // An omitted type is the schema default `obj`, as python-pptx reads it.
        let token = token.unwrap_or_else(|| "obj".to_owned());
        let placeholder_type = py
            .import("rpptx.enum.shapes")?
            .getattr("PP_PLACEHOLDER")?
            .call_method1("from_xml", (token,))?
            .unbind();
        Ok(PyPlaceholderFormat {
            idx,
            placeholder_type,
        })
    }
}

/// A placeholder's index and type, read when `Shape.placeholder_format` is.
#[pyclass(name = "PlaceholderFormat", frozen)]
pub struct PyPlaceholderFormat {
    idx: u32,
    placeholder_type: Py<PyAny>,
}

#[pymethods]
impl PyPlaceholderFormat {
    #[getter]
    fn idx(&self) -> u32 {
        self.idx
    }

    #[getter]
    fn r#type(&self, py: Python<'_>) -> Py<PyAny> {
        self.placeholder_type.clone_ref(py)
    }
}

#[pyclass(name = "ShapeClickAction")]
pub struct PyShapeClickAction {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyShapeClickAction {
    #[getter]
    fn hyperlink(&self, py: Python<'_>) -> PyResult<Py<PyShapeHyperlink>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "shape click action",
            ".click_action",
        )?;
        Py::new(
            py,
            PyShapeHyperlink {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
            },
        )
    }

    /// The slide the click action jumps to, or `None` for any other action.
    #[getter]
    fn target_slide(&self, py: Python<'_>) -> PyResult<Option<Py<PySlide>>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "shape click action",
            ".click_action",
        )?;
        let shape_id = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.non_visual_id())
            .ok_or_else(|| PyValueError::new_err("shape has no id"))?;
        let Some(target) = presentation
            .inner
            .shape_target_slide(slide_index(&self.path)?, shape_id)
            .map_err(|error| rpptx_to_pyerr(py, error))?
        else {
            return Ok(None);
        };
        let path = presentation
            .revisions
            .capture(smallvec![PathSeg::Slide(target)]);
        Py::new(py, PySlide::new(self.presentation.clone_ref(py), path)).map(Some)
    }

    /// Makes the click action jump to `slide`, or removes it for `None`, as
    /// python-pptx does. The relationship the old click action used goes when
    /// nothing else on the slide uses it.
    #[setter]
    fn set_target_slide(&self, py: Python<'_>, slide: Option<PyRef<'_, PySlide>>) -> PyResult<()> {
        let target = match slide {
            Some(slide) if !slide.presentation.is(&self.presentation) => {
                return Err(PyValueError::new_err("slide is not in this presentation"));
            }
            Some(slide) => Some(slide.validate(py)?),
            None => None,
        };
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "shape click action",
            ".click_action",
        )?;
        let shape_id = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.non_visual_id())
            .ok_or_else(|| PyValueError::new_err("shape has no id"))?;
        presentation
            .inner
            .set_shape_target_slide(slide_index(&self.path)?, shape_id, target)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pyclass(name = "ShapeHyperlink")]
pub struct PyShapeHyperlink {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyShapeHyperlink {
    #[getter]
    fn address(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "shape hyperlink",
            ".click_action.hyperlink",
        )?;
        let shape_id = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.non_visual_id())
            .ok_or_else(|| PyValueError::new_err("shape has no id"))?;
        presentation
            .inner
            .part_shape_hyperlink_address(part_of(&self.path)?, shape_id)
            .map(|address| address.map(str::to_owned))
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[setter]
    fn set_address(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let mut presentation = self.presentation.borrow_mut(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "shape hyperlink",
            ".click_action.hyperlink",
        )?;
        let shape_id = shape_ref_at(&presentation.inner, &self.path)
            .and_then(|shape| shape.non_visual_id())
            .ok_or_else(|| PyValueError::new_err("shape has no id"))?;
        presentation
            .inner
            .set_part_shape_hyperlink(
                part_of(&self.path)?,
                shape_id,
                value.filter(|value| !value.is_empty()),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pyclass(name = "ShapeCollection")]
pub struct PyShapeCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyShapeCollection {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<rpptx::PartRef> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "shape collection", ".shapes")?;
        part_of(&self.path)
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let part = self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        if self
            .path
            .segs
            .iter()
            .any(|segment| matches!(segment, PathSeg::Shape(_)))
        {
            return Ok(shape_ref_at(&presentation.inner, &self.path)
                .map_or(0, |shape| shape.child_count()));
        }
        Ok(presentation
            .inner
            .part_shapes(part)
            .map_or(0, |shapes| shapes.len()))
    }

    fn require_slide_root(&self) -> PyResult<()> {
        if self
            .path
            .segs
            .iter()
            .any(|segment| matches!(segment, PathSeg::Shape(_)))
        {
            return Err(PyValueError::new_err(
                "nested shape collections cannot remove or move shapes",
            ));
        }
        Ok(())
    }

    /// Returns the slide, layout or master and the group path that
    /// additions to this collection target, or the error an addition raises.
    fn target(&self, py: Python<'_>) -> PyResult<(rpptx::PartRef, Vec<usize>)> {
        let part = self.validate(py)?;
        let group = shape_indices(&self.path).collect::<Vec<_>>();
        let mut presentation = self.presentation.borrow_mut(py);
        if presentation.inner.part_shapes_mut(part, &group).is_some() {
            return Ok((part, group));
        }
        let fallback_group = shape_ref_at(&presentation.inner, &self.path)
            .is_some_and(|shape| shape.kind() == rpptx::ShapeKind::Group);
        Err(PyValueError::new_err(if fallback_group {
            "a group inside an mc:AlternateContent fallback cannot take new shapes"
        } else {
            "shape is not a group"
        }))
    }

    /// Adds one shape to the slide or group this collection holds, advances
    /// the revision once, and returns the new shape captured at it.
    fn add(
        &self,
        py: Python<'_>,
        add: impl FnOnce(&mut rpptx::ShapesMut<'_>) -> rpptx::Result<()>,
    ) -> PyResult<Py<PyShape>> {
        let (part, group) = self.target(py)?;
        let index = self.len(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut shapes = presentation
            .inner
            .part_shapes_mut(part, &group)
            .expect("the target is a slide, layout, master or group");
        add(&mut shapes).map_err(|error| rpptx_to_pyerr(py, error))?;
        drop(presentation);
        self.capture_added(py, index)
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyShape>> {
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Shape(index));
        let path = self.presentation.borrow(py).revisions.capture(segments);
        Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))
    }

    fn capture_added(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyShape>> {
        let mut presentation = self.presentation.borrow_mut(py);
        presentation.revisions.bump();
        let mut segments = self.path.segs.clone();
        segments.push(PathSeg::Shape(index));
        let path = presentation.revisions.capture(segments);
        drop(presentation);
        Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))
    }
}

#[pymethods]
impl PyShapeCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        self.len(py)
    }

    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let len = self.len(py)?;
        if let Ok(index) = key.extract::<isize>() {
            return Ok(self
                .item(py, normalize_index(index, len, "shape")?)?
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
            "shape indices must be integers or slices",
        ))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyShapeIterator>> {
        self.validate(py)?;
        Py::new(
            py,
            PyShapeIterator {
                presentation: self.presentation.clone_ref(py),
                path: self.path.clone(),
                index: 0,
            },
        )
    }

    #[getter]
    fn title(&self, py: Python<'_>) -> PyResult<Option<Py<PyShape>>> {
        let part = self.validate(py)?;
        let index = {
            let presentation = self.presentation.borrow(py);
            let Ok(mut shapes) = presentation.inner.part_shapes(part) else {
                return Ok(None);
            };
            shapes.position(|shape| matches!(shape.placeholder_type(), Some("title" | "ctrTitle")))
        };
        index.map(|index| self.item(py, index)).transpose()
    }

    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    fn add_textbox(
        &mut self,
        py: Python<'_>,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
    ) -> PyResult<Py<PyShape>> {
        self.add(py, |shapes| {
            shapes
                .add_textbox(
                    rpptx::Emu(left),
                    rpptx::Emu(top),
                    rpptx::Emu(width),
                    rpptx::Emu(height),
                )
                .map(drop)
        })
    }

    fn add_shape(
        &mut self,
        py: Python<'_>,
        shape_type: &Bound<'_, PyAny>,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
    ) -> PyResult<Py<PyShape>> {
        let preset = preset_name(py, shape_type)?;
        self.add(py, |shapes| {
            shapes
                .add_shape(
                    &preset,
                    rpptx::Emu(left),
                    rpptx::Emu(top),
                    rpptx::Emu(width),
                    rpptx::Emu(height),
                )
                .map(drop)
        })
    }

    fn add_connector(
        &mut self,
        py: Python<'_>,
        connector_type: i64,
        begin_x: i64,
        begin_y: i64,
        end_x: i64,
        end_y: i64,
    ) -> PyResult<Py<PyShape>> {
        let connector = match connector_type {
            1 => rpptx::ConnectorType::Straight,
            2 => rpptx::ConnectorType::Elbow,
            3 => rpptx::ConnectorType::Curve,
            _ => return Err(PyValueError::new_err("unsupported MSO_CONNECTOR value")),
        };
        self.add(py, |shapes| {
            shapes
                .add_connector(
                    connector,
                    rpptx::Emu(begin_x),
                    rpptx::Emu(begin_y),
                    rpptx::Emu(end_x),
                    rpptx::Emu(end_y),
                )
                .map(drop)
        })
    }

    /// Appends an empty group, which its own `shapes` collection populates.
    fn add_group_shape(&mut self, py: Python<'_>) -> PyResult<Py<PyShape>> {
        self.add(py, |shapes| shapes.add_group_shape().map(drop))
    }

    /// Removes one shape of this slide, layout or master and the package
    /// parts only it used.
    fn remove(&mut self, py: Python<'_>, shape: &Bound<'_, PyAny>) -> PyResult<()> {
        self.require_slide_root()?;
        let part = self.validate(py)?;
        let shape = shape.extract::<PyRef<'_, PyShape>>()?;
        if !shape.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("shape is not in this collection"));
        }
        shape.validate(py)?;
        let shape_index = match shape.path.segs.as_slice() {
            [_, PathSeg::Shape(index)] if part_of(&shape.path)? == part => *index,
            _ => return Err(PyValueError::new_err("shape is not in this collection")),
        };
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .remove_part_shape(part, shape_index)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    /// Moves the shape at `from_` so that it ends up at z-order index `to`,
    /// where later shapes draw on top.
    #[pyo3(name = "move")]
    fn move_shape(&mut self, py: Python<'_>, from_: isize, to: isize) -> PyResult<()> {
        self.require_slide_root()?;
        let part = self.validate(py)?;
        let len = self.len(py)?;
        let from_ = normalize_index(from_, len, "shape")?;
        let to = normalize_index(to, len, "shape")?;
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .move_part_shape(part, from_, to)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        presentation.revisions.bump();
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn add_table(
        &mut self,
        py: Python<'_>,
        rows: usize,
        cols: usize,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
    ) -> PyResult<Py<PyShape>> {
        self.add(py, |shapes| {
            shapes
                .add_table(
                    rows,
                    cols,
                    rpptx::Emu(left),
                    rpptx::Emu(top),
                    rpptx::Emu(width),
                    rpptx::Emu(height),
                )
                .map(drop)
        })
    }

    #[pyo3(signature = (image_file, left, top, width = None, height = None))]
    fn add_picture(
        &mut self,
        py: Python<'_>,
        image_file: &Bound<'_, PyAny>,
        left: i64,
        top: i64,
        width: Option<i64>,
        height: Option<i64>,
    ) -> PyResult<Py<PyShape>> {
        self.target(py)?;
        let (bytes, filename) = image_bytes(image_file)?;
        self.add(py, |shapes| {
            shapes
                .add_picture(
                    &bytes,
                    &filename,
                    rpptx::Emu(left),
                    rpptx::Emu(top),
                    width.map(rpptx::Emu),
                    height.map(rpptx::Emu),
                )
                .map(drop)
        })
    }
}

#[pyclass]
struct PyShapeIterator {
    presentation: Py<PyPresentation>,
    path: ContentPath,
    index: usize,
}

#[pymethods]
impl PyShapeIterator {
    fn __iter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyShape>>> {
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        if self.index >= collection.len(py)? {
            return Ok(None);
        }
        let index = self.index;
        self.index += 1;
        collection.item(py, index).map(Some)
    }
}

#[pyclass(name = "PlaceholderCollection")]
pub struct PyPlaceholderCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyPlaceholderCollection {
    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }
}

impl PyPlaceholderCollection {
    /// The z-order indices of the placeholders, in z-order.
    fn positions(&self, py: Python<'_>) -> PyResult<Vec<usize>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "placeholder collection",
            ".placeholders",
        )?;
        let shapes = presentation
            .inner
            .part_shapes(part_of(&self.path)?)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        Ok(shapes
            .enumerate()
            .filter(|(_, shape)| shape.placeholder_idx().is_some())
            .map(|(index, _)| index)
            .collect())
    }
}

#[pymethods]
impl PyPlaceholderCollection {
    /// The placeholder whose `idx` is `placeholder_idx`, as python-pptx
    /// `placeholders[idx]` finds it. A title placeholder has `idx` 0.
    fn __getitem__(&self, py: Python<'_>, placeholder_idx: u32) -> PyResult<Py<PyShape>> {
        let positions = self.positions(py)?;
        let presentation = self.presentation.borrow(py);
        let part = part_of(&self.path)?;
        let shape_index = positions
            .into_iter()
            .find(|index| {
                presentation
                    .inner
                    .part_shape(part, *index)
                    .and_then(|shape| shape.placeholder_idx())
                    == Some(placeholder_idx)
            })
            .ok_or_else(|| PyIndexError::new_err("placeholder index out of range"))?;
        drop(presentation);
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        collection.item(py, shape_index)
    }

    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.positions(py)?.len())
    }

    /// Iterates the placeholders in z-order.
    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        let items = self
            .positions(py)?
            .into_iter()
            .map(|index| collection.item(py, index))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, items)?
            .into_any()
            .try_iter()
            .map(Bound::into_any)
    }
}

/// The preset adjustments of one shape, like python-pptx `AdjustmentCollection`.
///
/// Values are normalized so that 1.0 is the raw guide value 100000.
#[pyclass(name = "AdjustmentCollection")]
pub struct PyAdjustmentCollection {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyAdjustmentCollection {
    fn values(&self, py: Python<'_>) -> PyResult<Vec<(String, f64)>> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "adjustments", ".adjustments")?;
        shape_ref_at(&presentation.inner, &self.path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .adjustments()
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyAdjustmentCollection {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.values(py)?.len())
    }

    fn __getitem__(&self, py: Python<'_>, index: isize) -> PyResult<f64> {
        let values = self.values(py)?;
        let index = normalize_index(index, values.len(), "adjustment")?;
        Ok(values[index].1 / 100_000.0)
    }

    fn __setitem__(&self, py: Python<'_>, index: isize, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let values = self.values(py)?;
        let index = normalize_index(index, values.len(), "adjustment")?;
        let value = value.extract::<f64>().map_err(|_| {
            PyValueError::new_err(format!(
                "adjustment value must be numeric, got {}",
                value
                    .repr()
                    .map_or_else(|_| "?".to_owned(), |repr| repr.to_string())
            ))
        })?;
        let raw = (value * 100_000.0).trunc();
        let mut presentation = self.presentation.borrow_mut(py);
        shape_mut_at(&mut presentation.inner, &self.path)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .set_adjust_value(&values[index].0, raw)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let values = self.values(py)?;
        PyList::new(py, values.iter().map(|(_, value)| value / 100_000.0))?.try_iter()
    }
}

/// A snapshot of the image one picture shows, like python-pptx `Image`.
#[pyclass(name = "Image", frozen)]
pub struct PyImage {
    blob: Vec<u8>,
    content_type: String,
    ext: String,
}

#[pymethods]
impl PyImage {
    #[getter]
    fn blob<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.blob)
    }

    #[getter]
    fn content_type(&self) -> &str {
        &self.content_type
    }

    #[getter]
    fn ext(&self) -> &str {
        &self.ext
    }
}

/// Returns the python-pptx style extension, such as `jpg` for JPEG.
fn image_extension(content_type: &str, part_name: &str) -> String {
    let known = match content_type.to_ascii_lowercase().as_str() {
        "image/png" => Some("png"),
        "image/jpeg" | "image/jpg" => Some("jpg"),
        "image/gif" => Some("gif"),
        "image/bmp" | "image/x-bmp" => Some("bmp"),
        "image/tiff" => Some("tiff"),
        "image/x-wmf" | "image/wmf" => Some("wmf"),
        "image/x-emf" | "image/emf" => Some("emf"),
        "image/svg+xml" => Some("svg"),
        "image/webp" => Some("webp"),
        _ => None,
    };
    known.map_or_else(
        || {
            part_name
                .rsplit_once('.')
                .map_or_else(String::new, |(_, extension)| extension.to_ascii_lowercase())
        },
        str::to_owned,
    )
}
