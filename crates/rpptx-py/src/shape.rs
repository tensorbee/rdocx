use std::path::PathBuf;

use oxml_py_support::{ContentPath, PathSeg};
use pyo3::exceptions::{PyIndexError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyByteArray, PyBytes, PyIterator, PyList, PySlice, PyString};

use crate::dml::{FillTarget, PyFillFormat, PyLineFormat};
use crate::normalize_index;
use crate::presentation::PyPresentation;
use crate::rpptx_to_pyerr;
use crate::table::PyTable;
use crate::text::PyTextFrame;
use crate::validate_path;

const MIN_COORDINATE: i64 = -27_273_042_329_600;
const MAX_COORDINATE: i64 = 27_273_042_316_900;
const ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
const ANGLE_UNITS_PER_TURN: i64 = 21_600_000;

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyShape>()?;
    module.add_class::<PyShapeCollection>()?;
    module.add_class::<PyPlaceholderCollection>()?;
    module.add_class::<PyImage>()?;
    module.add_class::<PyAdjustmentCollection>()?;
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
fn image_bytes(image_file: &Bound<'_, PyAny>) -> PyResult<(Vec<u8>, String)> {
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

fn check_coordinate(name: &str, value: i64, minimum: i64) -> PyResult<()> {
    if (minimum..=MAX_COORDINATE).contains(&value) {
        Ok(())
    } else {
        Err(PyValueError::new_err(format!(
            "{name} must be between {minimum} and {MAX_COORDINATE} EMU, got {value}"
        )))
    }
}

pub(crate) fn slide_index(path: &ContentPath) -> PyResult<usize> {
    path.segs
        .iter()
        .find_map(|segment| match segment {
            PathSeg::Slide(index) => Some(*index),
            _ => None,
        })
        .ok_or_else(|| PyIndexError::new_err("slide index is missing"))
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
    let slide = presentation.slide(slide_index(path).ok()?)?;
    let mut indices = shape_indices(path);
    let mut shape = slide.shape(indices.next()?)?;
    for index in indices {
        shape = shape.child(index)?;
    }
    Some(shape)
}

pub(crate) fn shape_mut_at<'a>(
    presentation: &'a mut rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeMut<'a>> {
    let slide = presentation.slide_mut(slide_index(path).ok()?)?;
    let mut indices = shape_indices(path);
    let mut shape = slide.into_shape_mut(indices.next()?)?;
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
        let top = self.read(py, |shape| {
            shape.position().map_or(rpptx::Emu(0), |(_, top)| top)
        })?;
        self.edit(py, |shape| shape.set_position(rpptx::Emu(value), top))
    }

    #[setter]
    fn set_top(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("top", value, MIN_COORDINATE)?;
        let left = self.read(py, |shape| {
            shape.position().map_or(rpptx::Emu(0), |(left, _)| left)
        })?;
        self.edit(py, |shape| shape.set_position(left, rpptx::Emu(value)))
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("width", value, 0)?;
        let height = self.read(py, |shape| {
            shape.size().map_or(rpptx::Emu(0), |(_, height)| height)
        })?;
        self.edit(py, |shape| shape.set_size(rpptx::Emu(value), height))
    }

    #[setter]
    fn set_height(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("height", value, 0)?;
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
            PyLineFormat::new(self.presentation.clone_ref(py), self.path.clone()),
        )
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

    fn validate(&self, py: Python<'_>) -> PyResult<usize> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "shape collection", ".shapes")?;
        slide_index(&self.path)
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let slide = self.validate(py)?;
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
            .slide(slide)
            .map_or(0, |slide| slide.shapes().len()))
    }

    fn require_slide_root(&self) -> PyResult<()> {
        if self
            .path
            .segs
            .iter()
            .any(|segment| matches!(segment, PathSeg::Shape(_)))
        {
            return Err(PyValueError::new_err(
                "nested shape collections are read-only",
            ));
        }
        Ok(())
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
        let slide_index = self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let Some(slide) = presentation.inner.slide(slide_index) else {
            return Ok(None);
        };
        let Some(title) = slide.title() else {
            return Ok(None);
        };
        let index = slide
            .shapes()
            .position(|shape| shape.placeholder_idx() == title.placeholder_idx())
            .expect("title is an immediate slide shape");
        drop(presentation);
        self.item(py, index).map(Some)
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
        self.require_slide_root()?;
        let slide_index = self.validate(py)?;
        let index = self.len(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .slide_mut(slide_index)
            .expect("validated slide")
            .add_textbox(
                rpptx::Emu(left),
                rpptx::Emu(top),
                rpptx::Emu(width),
                rpptx::Emu(height),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.capture_added(py, index)
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
        self.require_slide_root()?;
        let preset = if shape_type.is_instance_of::<PyString>() {
            shape_type.extract::<String>()?
        } else {
            let value = shape_type.extract::<i64>()?;
            py.import("rpptx.enum.shapes")?
                .getattr("MSO_SHAPE")?
                .call1((value,))
                .map_err(|_| PyValueError::new_err("unsupported MSO_SHAPE value"))?
                .getattr("xml_value")?
                .extract::<String>()?
        };
        let slide_index = self.validate(py)?;
        let index = self.len(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .slide_mut(slide_index)
            .expect("validated slide")
            .add_shape(
                &preset,
                rpptx::Emu(left),
                rpptx::Emu(top),
                rpptx::Emu(width),
                rpptx::Emu(height),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.capture_added(py, index)
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
        self.require_slide_root()?;
        let connector = match connector_type {
            1 => rpptx::ConnectorType::Straight,
            2 => rpptx::ConnectorType::Elbow,
            3 => rpptx::ConnectorType::Curve,
            _ => return Err(PyValueError::new_err("unsupported MSO_CONNECTOR value")),
        };
        let slide_index = self.validate(py)?;
        let index = self.len(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .slide_mut(slide_index)
            .expect("validated slide")
            .add_connector(
                connector,
                rpptx::Emu(begin_x),
                rpptx::Emu(begin_y),
                rpptx::Emu(end_x),
                rpptx::Emu(end_y),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.capture_added(py, index)
    }

    /// Appends an empty group. Populating groups is not supported yet.
    fn add_group_shape(&mut self, py: Python<'_>) -> PyResult<Py<PyShape>> {
        self.require_slide_root()?;
        let slide_index = self.validate(py)?;
        let index = self.len(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .slide_mut(slide_index)
            .expect("validated slide")
            .add_group_shape()
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.capture_added(py, index)
    }

    /// Removes one shape of this slide and the package parts only it used.
    fn remove(&mut self, py: Python<'_>, shape: &Bound<'_, PyAny>) -> PyResult<()> {
        self.require_slide_root()?;
        let slide_index = self.validate(py)?;
        let shape = shape.extract::<PyRef<'_, PyShape>>()?;
        if !shape.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err("shape is not in this collection"));
        }
        shape.validate(py)?;
        let shape_index = match shape.path.segs.as_slice() {
            [PathSeg::Slide(slide), PathSeg::Shape(index)] if *slide == slide_index => *index,
            _ => return Err(PyValueError::new_err("shape is not in this collection")),
        };
        let mut presentation = self.presentation.borrow_mut(py);
        presentation
            .inner
            .remove_shape(slide_index, shape_index)
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
        self.require_slide_root()?;
        let slide_index = self.validate(py)?;
        let index = self.len(py)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .slide_mut(slide_index)
            .expect("validated slide")
            .add_table(
                rows,
                cols,
                rpptx::Emu(left),
                rpptx::Emu(top),
                rpptx::Emu(width),
                rpptx::Emu(height),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.capture_added(py, index)
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
        self.require_slide_root()?;
        let slide_index = self.validate(py)?;
        let index = self.len(py)?;
        let (bytes, filename) = image_bytes(image_file)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .add_picture(
                slide_index,
                &bytes,
                &filename,
                rpptx::Emu(left),
                rpptx::Emu(top),
                width.map(rpptx::Emu),
                height.map(rpptx::Emu),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        self.capture_added(py, index)
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

#[pymethods]
impl PyPlaceholderCollection {
    fn __getitem__(&self, py: Python<'_>, placeholder_idx: u32) -> PyResult<Py<PyShape>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "placeholder collection",
            ".placeholders",
        )?;
        let slide_index = slide_index(&self.path)?;
        let slide = presentation
            .inner
            .slide(slide_index)
            .ok_or_else(|| PyIndexError::new_err("slide index out of range"))?;
        let placeholder = slide
            .placeholder(placeholder_idx)
            .ok_or_else(|| PyIndexError::new_err("placeholder index out of range"))?;
        let shape_index = slide
            .shapes()
            .position(|shape| shape.placeholder_idx() == placeholder.placeholder_idx())
            .expect("placeholder is an immediate slide shape");
        drop(presentation);
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        collection.item(py, shape_index)
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
