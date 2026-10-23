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

// These are the concrete scalar setters with a pre-native local change proof.
// All other callbacks keep the candidate transaction.
enum ShapeEdit {
    Staged,
    AltText(Option<String>),
    Position(rpptx::Emu, rpptx::Emu),
    Size(rpptx::Emu, rpptx::Emu),
    Rotation(rpptx::Angle),
}

fn part_xml(inner: &rpptx::Presentation, owner: rpptx::ShapeOwner) -> rpptx::Result<Vec<u8>> {
    match owner {
        rpptx::ShapeOwner::Slide(index) => inner.slide_xml(index),
        rpptx::ShapeOwner::Layout(index) => inner.layout_xml(index),
        rpptx::ShapeOwner::Master(index) => inner.master_xml(index),
    }
}

/// Left, top, right, and bottom picture crop insets, as the facade reports them.
type Crop = (
    rpptx::Percent1000,
    rpptx::Percent1000,
    rpptx::Percent1000,
    rpptx::Percent1000,
);

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyPlaceholderFormat>()?;
    module.add_class::<PyShape>()?;
    module.add_class::<PyChart>()?;
    module.add_class::<PyChartSeries>()?;
    module.add_class::<PyShapeClickAction>()?;
    module.add_class::<PyShapeHyperlink>()?;
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

pub(crate) fn slide_index(path: &ContentPath) -> PyResult<usize> {
    path.segs
        .iter()
        .find_map(|segment| match segment {
            PathSeg::Slide(index) => Some(*index),
            _ => None,
        })
        .ok_or_else(|| PyIndexError::new_err("slide index is missing"))
}

pub(crate) fn shape_owner(path: &ContentPath) -> PyResult<rpptx::ShapeOwner> {
    path.segs
        .iter()
        .find_map(|segment| match segment {
            PathSeg::Slide(index) => Some(rpptx::ShapeOwner::Slide(*index)),
            PathSeg::Layout(index) => Some(rpptx::ShapeOwner::Layout(*index)),
            PathSeg::Master(index) => Some(rpptx::ShapeOwner::Master(*index)),
            _ => None,
        })
        .ok_or_else(|| PyIndexError::new_err("shape part owner is missing"))
}

fn shape_indices(path: &ContentPath) -> impl Iterator<Item = usize> + '_ {
    path.segs.iter().filter_map(|segment| match segment {
        PathSeg::Shape(index) => Some(*index),
        _ => None,
    })
}

pub(crate) fn shape_ref_at<'a>(
    presentation: &'a PyPresentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeRef<'a>> {
    let current = presentation.current_path(path)?;
    native_shape_ref_at(&presentation.inner, &current)
}

pub(crate) fn native_shape_ref_at<'a>(
    presentation: &'a rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeRef<'a>> {
    let owner = shape_owner(path).ok()?;
    let indices = shape_indices(path).collect::<Vec<_>>();
    presentation.part_shape_at(owner, &indices).ok()
}

pub(crate) fn shape_mut_at<'a>(
    presentation: &'a mut PyPresentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeMut<'a>> {
    let current = presentation.current_path(path)?;
    native_shape_mut_at(&mut presentation.inner, &current)
}

pub(crate) fn native_shape_mut_at<'a>(
    presentation: &'a mut rpptx::Presentation,
    path: &ContentPath,
) -> Option<rpptx::ShapeMut<'a>> {
    let owner = shape_owner(path).ok()?;
    let indices = shape_indices(path).collect::<Vec<_>>();
    presentation.shape_mut_at(owner, &indices).ok()
}

const CHART_TYPES: [(i64, rpptx::ChartKind); 7] = [
    (1, rpptx::ChartKind::Area),
    (51, rpptx::ChartKind::Bar),
    (-4120, rpptx::ChartKind::Doughnut),
    (65, rpptx::ChartKind::Line),
    (5, rpptx::ChartKind::Pie),
    (81, rpptx::ChartKind::Radar),
    (74, rpptx::ChartKind::Scatter),
];

fn checked_chart_inputs(
    chart_type: &Bound<'_, PyAny>,
    data: &Bound<'_, PyAny>,
) -> PyResult<(rpptx::ChartKind, rpptx::ChartData)> {
    if chart_type.is_instance_of::<pyo3::types::PyBool>() {
        return Err(PyTypeError::new_err(
            "chart type must be an XL_CHART_TYPE integer",
        ));
    }
    let value = chart_type.extract::<i64>()?;
    let kind = CHART_TYPES
        .iter()
        .find_map(|(v, k)| (*v == value).then_some(*k))
        .ok_or_else(|| PyValueError::new_err(format!("unsupported XL_CHART_TYPE value {value}")))?;
    let (categories, series, number_format) =
        data.call_method0("_data")?
            .extract::<(Vec<String>, Vec<(String, Vec<f64>)>, Option<String>)>()?;
    Ok((
        kind,
        rpptx::ChartData {
            categories,
            series,
            number_format,
            ..Default::default()
        },
    ))
}

#[pyclass(name = "Shape")]
pub struct PyShape {
    pub(crate) presentation: Py<PyPresentation>,
    pub(crate) path: ContentPath,
}

impl PyShape {
    fn set_flip_axis(&self, py: Python<'_>, value: bool, horizontal: bool) -> PyResult<()> {
        self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner absent"))?;
        let shape = native_shape_ref_at(&presentation.inner, &current)
            .ok_or_else(|| PyIndexError::new_err("shape owner absent"))?;
        let current_flags = shape
            .flips()
            .ok_or_else(|| PyValueError::new_err("opaque shape has no flip flags"))?;
        if value
            == if horizontal {
                current_flags.0
            } else {
                current_flags.1
            }
        {
            return Ok(());
        }
        if matches!(
            shape.shape_type(),
            Some(rpptx::ShapeType::Table | rpptx::ShapeType::Chart)
        ) {
            return Err(PyValueError::new_err(
                "PowerPoint does not support flip flags on table or chart frames",
            ));
        }
        let mut candidate = presentation.inner.clone();
        candidate
            .materialize_owner_geometry(
                shape_owner(&current)?,
                &shape_indices(&current).collect::<Vec<_>>(),
            )
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        let flags = native_shape_ref_at(&candidate, &current)
            .and_then(|shape| shape.flips())
            .ok_or_else(|| PyValueError::new_err("opaque shape has no flip flags"))?;
        let flags = if horizontal {
            (value, flags.1)
        } else {
            (flags.0, value)
        };
        native_shape_mut_at(&mut candidate, &current)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .set_flips(flags.0, flags.1)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        if candidate.to_bytes().map_err(|e| rpptx_to_pyerr(py, e))?
            == presentation
                .inner
                .to_bytes()
                .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let survivors = presentation.unchanged_owners()?;
        let operation = if horizontal {
            "Shape.flip_h"
        } else {
            "Shape.flip_v"
        };
        let prepared =
            presentation.prepare_owners(&candidate, operation, &survivors, false, false, false)?;
        presentation.publish_owners(candidate, prepared)
    }
    fn connect_end(
        &self,
        py: Python<'_>,
        target: &Bound<'_, PyAny>,
        site: u32,
        end: rpptx::ConnectorEnd,
        operation: &'static str,
    ) -> PyResult<()> {
        let target = target.extract::<PyRef<'_, PyShape>>()?;
        if !target.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "shape belongs to another presentation",
            ));
        }
        self.validate(py)?;
        target.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let connector_path = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("connector owner absent"))?;
        let target_path = presentation
            .current_path(&target.path)
            .ok_or_else(|| PyIndexError::new_err("target owner absent"))?;
        let slide = shape_owner(&connector_path)?;
        if shape_owner(&target_path)? != slide {
            return Err(PyValueError::new_err(
                "connection target belongs to another slide",
            ));
        }
        let id = presentation.checked_shape_id(&target.path)?;
        let mut candidate = presentation.inner.clone();
        candidate
            .connect_connector_at(
                slide,
                &shape_indices(&connector_path).collect::<Vec<_>>(),
                end,
                id,
                site,
            )
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        if candidate.to_bytes().map_err(|e| rpptx_to_pyerr(py, e))?
            == presentation
                .inner
                .to_bytes()
                .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let survivors = presentation.unchanged_owners()?;
        let prepared =
            presentation.prepare_owners(&candidate, operation, &survivors, false, false, false)?;
        presentation.publish_owners(candidate, prepared)
    }

    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<()> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "shape", "")
    }

    fn local_edit_change(
        py: Python<'_>,
        presentation: &PyPresentation,
        current: &ContentPath,
        edit: &ShapeEdit,
    ) -> PyResult<Option<bool>> {
        let shape = native_shape_ref_at(&presentation.inner, current)
            .ok_or_else(|| PyIndexError::new_err("shape owner absent"))?;
        if let ShapeEdit::AltText(value) = edit {
            if shape.kind() != rpptx::ShapeKind::Shape {
                return Ok(None);
            }
            let xml = presentation
                .inner
                .shape_xml_at(
                    shape_owner(current)?,
                    &shape_indices(current).collect::<Vec<_>>(),
                )
                .map_err(|error| rpptx_to_pyerr(py, error))?;
            let Ok(mut preview) = rpptx_oxml::shape_tree::CT_Shape::from_xml(&xml) else {
                return Ok(None);
            };
            if preview.alt_text() != shape.alt_text() {
                return Ok(None);
            }
            let original = preview.clone();
            if preview.set_alt_text(value.as_deref()).is_err() {
                return Ok(None);
            }
            // The exact native model setter validates and edits only cNvPr.
            // Equality includes its lexical attribute bytes, not just the getter.
            return Ok(Some(preview != original));
        }
        if matches!(edit, ShapeEdit::Staged) {
            return Ok(None);
        }
        let rpptx::ShapeOwner::Slide(slide) = shape_owner(current)? else {
            return Ok(None);
        };
        // Connector/group context can have side effects outside this leaf.
        // Uncertainty stages. This walks only the affected slide.
        for child in presentation
            .inner
            .slide(slide)
            .ok_or_else(|| PyIndexError::new_err("slide owner absent"))?
            .shapes()
        {
            crate::presentation::note_fx260_ppt(1);
            if matches!(
                child.kind(),
                rpptx::ShapeKind::Connector
                    | rpptx::ShapeKind::Group
                    | rpptx::ShapeKind::AlternateContent
            ) {
                return Ok(None);
            }
        }
        let (Some(position), Some(size)) = (shape.position(), shape.size()) else {
            return Ok(None);
        };
        if shape.non_visual_id().is_none() || size.0.0 < 0 || size.1.0 < 0 {
            return Ok(None);
        }
        Ok(match edit {
            ShapeEdit::Position(x, y) => Some(position != (*x, *y)),
            ShapeEdit::Size(width, height) => Some(size != (*width, *height)),
            ShapeEdit::Rotation(angle) => shape.rotation().map(|old| old != *angle),
            ShapeEdit::AltText(_) | ShapeEdit::Staged => None,
        })
    }

    fn edit(
        &self,
        py: Python<'_>,
        edit: ShapeEdit,
        operation: &'static str,
        change: impl FnOnce(&mut rpptx::ShapeMut<'_>) -> rpptx::Result<()>,
    ) -> PyResult<()> {
        self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner absent"))?;
        match Self::local_edit_change(py, &presentation, &current, &edit)? {
            Some(false) => return Ok(()),
            Some(true) => {
                let revision = presentation.prepare_local_property(operation)?;
                let mut shape = native_shape_mut_at(&mut presentation.inner, &current)
                    .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?;
                change(&mut shape).map_err(|error| rpptx_to_pyerr(py, error))?;
                return presentation.commit_local_property(revision);
            }
            None => {}
        }
        let owner = shape_owner(&current)?;
        let before_xml =
            part_xml(&presentation.inner, owner).map_err(|error| rpptx_to_pyerr(py, error))?;
        let before_identity = {
            let shape = native_shape_ref_at(&presentation.inner, &current)
                .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?;
            (shape.non_visual_id(), shape.kind(), shape.table().is_some())
        };
        let mut candidate = presentation.inner.clone();
        let mut shape = native_shape_mut_at(&mut candidate, &current)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?;
        change(&mut shape).map_err(|error| rpptx_to_pyerr(py, error))?;
        let after_identity = {
            let shape = native_shape_ref_at(&candidate, &current)
                .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?;
            (shape.non_visual_id(), shape.kind(), shape.table().is_some())
        };
        if before_identity != after_identity {
            return Err(PyValueError::new_err(
                "survival proof changes physical owner identity",
            ));
        }
        let after_xml = part_xml(&candidate, owner).map_err(|error| rpptx_to_pyerr(py, error))?;
        if before_xml == after_xml {
            return Ok(());
        }
        let revision = presentation.prepare_local_property(operation)?;
        presentation.commit_local_property(revision)?;
        presentation.inner = candidate;
        Ok(())
    }

    fn read<T>(&self, py: Python<'_>, read: impl FnOnce(rpptx::ShapeRef<'_>) -> T) -> PyResult<T> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        shape_ref_at(&presentation, &self.path)
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
        operation: &'static str,
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
        self.edit(py, ShapeEdit::Staged, operation, |shape| {
            shape.set_crop(left, top, right, bottom)
        })
    }

    fn prepared_geometry(
        &self,
        py: Python<'_>,
    ) -> PyResult<Option<(rpptx::Emu, rpptx::Emu, rpptx::Emu, rpptx::Emu)>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner absent"))?;
        let path = shape_indices(&current).collect::<Vec<_>>();
        let owner = shape_owner(&current)?;
        if shape_ref_at(&presentation, &self.path)
            .is_some_and(|shape| shape.position().is_some() && shape.size().is_some())
        {
            return presentation
                .inner
                .effective_geometry_at(owner, &path)
                .map_err(|error| rpptx_to_pyerr(py, error));
        }
        let mut candidate = presentation.inner.clone();
        candidate
            .materialize_owner_geometry(owner, &path)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        candidate
            .effective_geometry_at(owner, &path)
            .map_err(|e| rpptx_to_pyerr(py, e))
    }

    fn picture_id(&self, py: Python<'_>) -> PyResult<(rpptx::ShapeOwner, u32)> {
        let (kind, id) = self.read(py, |shape| (shape.kind(), shape.non_visual_id()))?;
        let resolved = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        match (kind, id) {
            (rpptx::ShapeKind::Picture, Some(_)) => Ok((
                shape_owner(&resolved)?,
                self.presentation.borrow(py).checked_shape_id(&self.path)?,
            )),
            _ => Err(PyValueError::new_err("shape is not a picture")),
        }
    }
}

#[pymethods]
impl PyShape {
    #[getter]
    fn _element(&self, py: Python<'_>) -> PyResult<()> {
        self.read(py, |_| ())?;
        Err(PyNotImplementedError::new_err(
            "Shape._element does not expose lxml; use shape.element.xml to inspect XML or shape.element.replace_xml(...) for checked replacement",
        ))
    }

    #[getter]
    pub(crate) fn element(
        &self,
        py: Python<'_>,
    ) -> PyResult<Py<crate::presentation::PyPptElement>> {
        self.read(py, |_| ())?;
        let presentation = self.presentation.borrow(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyValueError::new_err("Element owner is absent"))?;
        let path = presentation.capture_path(current.segs, None)?;
        drop(presentation);
        crate::presentation::PyPptElement::new_checked(
            py,
            self.presentation.clone_ref(py),
            path,
            crate::presentation::PptElementKind::Shape,
        )
    }

    #[getter]
    fn is_placeholder(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |shape| shape.placeholder_idx().is_some())
    }

    /// Removes `p:style` from a shape or connector. Absence is not an error.
    ///
    /// A picture, table, group, or alternate content raises and does not publish.
    fn remove_style(&self, py: Python<'_>) -> PyResult<()> {
        self.edit(py, ShapeEdit::Staged, "Shape.remove_style", |shape| {
            shape.clear_style().map(drop)
        })
    }

    #[getter]
    fn placeholder_format(&self, py: Python<'_>) -> PyResult<PyPlaceholderFormat> {
        if !self.is_placeholder(py)? {
            return Err(PyValueError::new_err("shape is not a placeholder"));
        }
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Ok(PyPlaceholderFormat {
            presentation: self.presentation.clone_ref(py),
            path,
        })
    }

    fn ungroup<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, pyo3::types::PyList>> {
        self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("group owner absent"))?;
        let Some((PathSeg::Shape(index), parent)) = current.segs.split_last() else {
            return Err(PyValueError::new_err("missing group path"));
        };
        let slide = shape_owner(&current)?;
        let group = parent
            .iter()
            .filter_map(|s| {
                if let PathSeg::Shape(i) = s {
                    Some(*i)
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        let mut candidate = presentation.inner.clone();
        let range = candidate
            .shapes_mut_at(slide, &group)
            .map_err(|error| rpptx_to_pyerr(py, error))?
            .ungroup(*index)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        let survivors = presentation.ungrouped_shape_owners(parent, *index, range.len())?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Shape.ungroup",
            &survivors,
            false,
            false,
            false,
        )?;
        let mut handles = Vec::new();
        handles.try_reserve(range.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("ungroup result allocation failed")
        })?;
        for index in range {
            let mut destination = smallvec::SmallVec::from_slice(parent);
            destination.push(PathSeg::Shape(index));
            let path = presentation.capture_prepared_path(&prepared, destination, None)?;
            handles.push(Py::new(
                py,
                PyShape::new(self.presentation.clone_ref(py), path),
            )?);
        }
        let result = pyo3::types::PyList::new(py, handles)?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    #[getter]
    fn flip_h(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |shape| shape.flips())?
            .map(|flags| flags.0)
            .ok_or_else(|| PyValueError::new_err("opaque shape has no flip flags"))
    }
    #[setter]
    fn set_flip_h(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_flip_axis(py, value, true)
    }
    #[getter]
    fn flip_v(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |shape| shape.flips())?
            .map(|flags| flags.1)
            .ok_or_else(|| PyValueError::new_err("opaque shape has no flip flags"))
    }
    #[setter]
    fn set_flip_v(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.set_flip_axis(py, value, false)
    }
    fn begin_connect(&self, py: Python<'_>, shape: &Bound<'_, PyAny>, site: u32) -> PyResult<()> {
        self.connect_end(
            py,
            shape,
            site,
            rpptx::ConnectorEnd::Begin,
            "Shape.begin_connect",
        )
    }
    fn end_connect(&self, py: Python<'_>, shape: &Bound<'_, PyAny>, site: u32) -> PyResult<()> {
        self.connect_end(
            py,
            shape,
            site,
            rpptx::ConnectorEnd::End,
            "Shape.end_connect",
        )
    }
    #[getter]
    fn begin_x(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::Begin)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        length(py, Some(point.0))
    }
    #[setter]
    fn set_begin_x(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("begin_x", value, MIN_COORDINATE)?;
        let mut point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::Begin)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        point.0 = rpptx::Emu(value);
        self.edit(py, ShapeEdit::Staged, "Shape.begin_x", |shape| {
            shape.set_connector_endpoint(rpptx::ConnectorEnd::Begin, point.0, point.1)
        })
    }
    #[getter]
    fn begin_y(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::Begin)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        length(py, Some(point.1))
    }
    #[setter]
    fn set_begin_y(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("begin_y", value, MIN_COORDINATE)?;
        let mut point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::Begin)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        point.1 = rpptx::Emu(value);
        self.edit(py, ShapeEdit::Staged, "Shape.begin_y", |shape| {
            shape.set_connector_endpoint(rpptx::ConnectorEnd::Begin, point.0, point.1)
        })
    }
    #[getter]
    fn end_x(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::End)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        length(py, Some(point.0))
    }
    #[setter]
    fn set_end_x(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("end_x", value, MIN_COORDINATE)?;
        let mut point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::End)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        point.0 = rpptx::Emu(value);
        self.edit(py, ShapeEdit::Staged, "Shape.end_x", |shape| {
            shape.set_connector_endpoint(rpptx::ConnectorEnd::End, point.0, point.1)
        })
    }
    #[getter]
    fn end_y(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::End)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        length(py, Some(point.1))
    }
    #[setter]
    fn set_end_y(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("end_y", value, MIN_COORDINATE)?;
        let mut point = self
            .read(py, |shape| {
                shape.connector_endpoint(rpptx::ConnectorEnd::End)
            })?
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        point.1 = rpptx::Emu(value);
        self.edit(py, ShapeEdit::Staged, "Shape.end_y", |shape| {
            shape.set_connector_endpoint(rpptx::ConnectorEnd::End, point.0, point.1)
        })
    }

    #[getter]
    fn click_action(&self, py: Python<'_>) -> PyResult<Py<PyShapeClickAction>> {
        self.validate(py)?;
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Py::new(
            py,
            PyShapeClickAction {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
    }

    #[getter]
    fn left(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.position())
            .map(|(left, _)| left);
        drop(presentation);
        length(py, value)
    }

    #[getter]
    fn alt_text(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.read(py, |shape| shape.alt_text())
    }

    #[setter]
    pub(crate) fn set_alt_text(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.edit(
            py,
            ShapeEdit::AltText(value.map(str::to_owned)),
            "Shape.alt_text",
            |shape| shape.set_alt_text(value),
        )
    }

    #[getter]
    fn alt_title(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.read(py, |shape| shape.title())
    }

    #[setter]
    fn set_alt_title(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.edit(py, ShapeEdit::Staged, "Shape.alt_title", |shape| {
            shape.set_title(value)
        })
    }

    #[getter]
    fn decorative(&self, py: Python<'_>) -> PyResult<bool> {
        self.read(py, |shape| shape.decorative())
    }

    #[setter]
    fn set_decorative(&self, py: Python<'_>, value: bool) -> PyResult<()> {
        self.edit(py, ShapeEdit::Staged, "Shape.decorative", |shape| {
            shape.set_decorative(value)
        })
    }

    #[getter]
    fn top(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.position())
            .map(|(_, top)| top);
        drop(presentation);
        length(py, value)
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.size())
            .map(|(width, _)| width);
        drop(presentation);
        length(py, value)
    }

    #[getter]
    fn height(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let value = shape_ref_at(&presentation, &self.path)
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
        let resolved = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let shape_path = shape_indices(&resolved).collect::<Vec<_>>();
        let geometry = self
            .presentation
            .borrow(py)
            .inner
            .effective_geometry_at(shape_owner(&resolved)?, &shape_path)
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
        Ok(shape_ref_at(&self.presentation.borrow(py), &self.path)
            .and_then(|shape| shape.non_visual_id()))
    }

    #[setter]
    pub(crate) fn set_left(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("left", value, MIN_COORDINATE)?;
        let top = self
            .prepared_geometry(py)?
            .map_or(rpptx::Emu(0), |geometry| geometry.1);
        self.edit(
            py,
            ShapeEdit::Position(rpptx::Emu(value), top),
            "Shape.left",
            |shape| shape.set_position(rpptx::Emu(value), top),
        )
    }

    #[setter]
    fn set_top(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("top", value, MIN_COORDINATE)?;
        let left = self
            .prepared_geometry(py)?
            .map_or(rpptx::Emu(0), |geometry| geometry.0);
        self.edit(
            py,
            ShapeEdit::Position(left, rpptx::Emu(value)),
            "Shape.top",
            |shape| shape.set_position(left, rpptx::Emu(value)),
        )
    }

    #[setter]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("width", value, 0)?;
        let height = self
            .prepared_geometry(py)?
            .map_or(rpptx::Emu(0), |geometry| geometry.3);
        self.edit(
            py,
            ShapeEdit::Size(rpptx::Emu(value), height),
            "Shape.width",
            |shape| shape.set_size(rpptx::Emu(value), height),
        )
    }

    #[setter]
    fn set_height(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        check_coordinate("height", value, 0)?;
        let width = self
            .prepared_geometry(py)?
            .map_or(rpptx::Emu(0), |geometry| geometry.2);
        self.edit(
            py,
            ShapeEdit::Size(width, rpptx::Emu(value)),
            "Shape.height",
            |shape| shape.set_size(width, rpptx::Emu(value)),
        )
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        self.validate(py)?;
        Ok(shape_ref_at(&self.presentation.borrow(py), &self.path)
            .and_then(|shape| shape.non_visual_name()))
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.edit(py, ShapeEdit::Staged, "Shape.name", |shape| {
            shape.set_name(value)
        })
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
        self.edit(py, ShapeEdit::Rotation(angle), "Shape.rotation", |shape| {
            shape.set_rotation(angle)
        })
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
        self.edit(py, ShapeEdit::Staged, "Shape.auto_shape_type", |shape| {
            shape.set_auto_shape_type(&preset)
        })
    }

    #[getter]
    fn adjustments(&self, py: Python<'_>) -> PyResult<Py<PyAdjustmentCollection>> {
        if self.read(py, |shape| shape.kind())? != rpptx::ShapeKind::Shape {
            return Err(PyValueError::new_err("shape has no adjustments"));
        }
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Py::new(
            py,
            PyAdjustmentCollection {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
    }

    #[getter]
    pub(crate) fn fill(&self, py: Python<'_>) -> PyResult<Py<PyFillFormat>> {
        self.require_shape_properties(py, "fill")?;
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Py::new(
            py,
            PyFillFormat::new(self.presentation.clone_ref(py), path, FillTarget::Shape),
        )
    }

    #[getter]
    fn line(&self, py: Python<'_>) -> PyResult<Py<PyLineFormat>> {
        self.require_shape_properties(py, "line")?;
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Py::new(
            py,
            PyLineFormat::new(self.presentation.clone_ref(py), path, FillTarget::Line),
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
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Py::new(
            py,
            PyShadowFormat::new(self.presentation.clone_ref(py), path),
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
        self.edit(py, ShapeEdit::Staged, "Shape.theme_effect_index", |shape| {
            shape.set_theme_effect_index(value)
        })
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
            .picture_image_at(slide, shape_id)
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
        self.set_crop_edge(py, |crop| &mut crop.0, value, "Shape.crop_left")
    }

    /// The share of the image cropped from the picture's top edge.
    #[getter]
    fn crop_top(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.1)
    }

    #[setter]
    fn set_crop_top(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.1, value, "Shape.crop_top")
    }

    /// The share of the image cropped from the picture's right edge.
    #[getter]
    fn crop_right(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.2)
    }

    #[setter]
    fn set_crop_right(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.2, value, "Shape.crop_right")
    }

    /// The share of the image cropped from the picture's bottom edge.
    #[getter]
    fn crop_bottom(&self, py: Python<'_>) -> PyResult<f64> {
        self.crop_edge(py, |crop| &mut crop.3)
    }

    #[setter]
    fn set_crop_bottom(&self, py: Python<'_>, value: f64) -> PyResult<()> {
        self.set_crop_edge(py, |crop| &mut crop.3, value, "Shape.crop_bottom")
    }

    /// Replaces this picture placeholder and returns the new picture owner.
    fn insert_picture(
        &self,
        py: Python<'_>,
        image_file: &Bound<'_, PyAny>,
    ) -> PyResult<Py<PyShape>> {
        self.validate(py)?;
        let (bytes, filename) = image_bytes(image_file)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("placeholder owner is absent"))?;
        if current
            .segs
            .iter()
            .any(|seg| !matches!(seg, PathSeg::Slide(_) | PathSeg::Shape(_)))
        {
            return Err(PyValueError::new_err(
                "picture insertion requires a physical slide placeholder",
            ));
        }
        let mut candidate = presentation.inner.clone();
        candidate
            .fill_picture_placeholder(
                slide_index(&current)?,
                &shape_indices(&current).collect::<Vec<_>>(),
                &bytes,
                &filename,
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = presentation.replaced_shape_owners(&current.segs)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Shape.insert_picture",
            &survivors,
            false,
            false,
            false,
        )?;
        let path = presentation.capture_prepared_path(&prepared, current.segs.clone(), None)?;
        let result = Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    /// Replaces the picture's image, keeping its position, size, and crop.
    fn replace_image(&self, py: Python<'_>, image_file: &Bound<'_, PyAny>) -> PyResult<()> {
        let (slide, shape_id) = self.picture_id(py)?;
        let (bytes, _) = image_bytes(image_file)?;
        self.presentation
            .borrow_mut(py)
            .inner
            .replace_picture_image_at(slide, shape_id, &bytes)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<Py<PyShapeCollection>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("group owner is absent"))?;
        let path = presentation.capture_xml_children_path(current.segs)?;
        drop(presentation);
        Py::new(
            py,
            PyShapeCollection::new(self.presentation.clone_ref(py), path),
        )
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.text())
            .ok_or_else(|| PyValueError::new_err("shape has no text"))
    }

    #[setter]
    fn set_text(&self, py: Python<'_>, text: &str) -> PyResult<()> {
        self.validate(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let mut candidate = presentation.inner.clone();
        native_shape_mut_at(&mut candidate, &current)
            .ok_or_else(|| PyIndexError::new_err("shape index out of range"))?
            .set_text(text)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = presentation.text_changed_owners(&self.path)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Shape.text",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }

    #[getter]
    fn has_text_frame(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        Ok(shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.text_frame())
            .is_some())
    }

    #[getter]
    fn text_frame(&self, py: Python<'_>) -> PyResult<Py<PyTextFrame>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        if shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.text_frame())
            .is_none()
        {
            return Err(PyValueError::new_err("shape has no text frame"));
        }
        drop(presentation);
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_text_body_path(current.segs, None)?;
        Py::new(py, PyTextFrame::new(self.presentation.clone_ref(py), path))
    }

    #[getter]
    fn has_chart(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.read(py, |shape| shape.shape_type())? == Some(rpptx::ShapeType::Chart))
    }

    #[getter]
    fn chart(&self, py: Python<'_>) -> PyResult<Py<PyChart>> {
        self.validate(py)?;
        if !self.has_chart(py)? {
            return Err(PyValueError::new_err("shape has no chart"));
        }
        let presentation = self.presentation.borrow(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("chart owner is absent"))?;
        let path = presentation.capture_xml_children_path(current.segs)?;
        drop(presentation);
        Py::new(
            py,
            PyChart {
                presentation: self.presentation.clone_ref(py),
                path,
            },
        )
    }

    #[getter]
    fn has_table(&self, py: Python<'_>) -> PyResult<bool> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        Ok(shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.table())
            .is_some())
    }

    #[getter]
    fn table(&self, py: Python<'_>) -> PyResult<Py<PyTable>> {
        self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        if shape_ref_at(&presentation, &self.path)
            .and_then(|shape| shape.table())
            .is_none()
        {
            return Err(PyValueError::new_err("shape has no table"));
        }
        drop(presentation);
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let path = self
            .presentation
            .borrow(py)
            .capture_xml_children_path(current.segs)?;
        Py::new(py, PyTable::new(self.presentation.clone_ref(py), path))
    }
}

/// Read-only placeholder metadata bound to its current causal shape owner.
#[pyclass(name = "PlaceholderFormat")]
pub struct PyPlaceholderFormat {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

#[pymethods]
impl PyPlaceholderFormat {
    #[getter]
    fn idx(&self, py: Python<'_>) -> PyResult<u32> {
        PyShape {
            presentation: self.presentation.clone_ref(py),
            path: self.path.clone(),
        }
        .read(py, |shape| shape.placeholder_idx())?
        .ok_or_else(|| PyValueError::new_err("shape is not a placeholder"))
    }

    #[getter]
    fn r#type(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let shape = PyShape {
            presentation: self.presentation.clone_ref(py),
            path: self.path.clone(),
        };
        let (idx, token) = shape.read(py, |shape| {
            (
                shape.placeholder_idx(),
                shape.placeholder_type().map(str::to_owned),
            )
        })?;
        idx.ok_or_else(|| PyValueError::new_err("shape is not a placeholder"))?;
        Ok(py
            .import("rpptx.enum.shapes")?
            .getattr("PP_PLACEHOLDER")?
            .call_method1("from_xml", (token.unwrap_or_else(|| "obj".into()),))?
            .unbind())
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
        let resolved = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let shape_id = presentation.checked_shape_id(&self.path)?;
        let Some(target) = presentation
            .inner
            .shape_target_slide(slide_index(&resolved)?, shape_id)
            .map_err(|error| rpptx_to_pyerr(py, error))?
        else {
            return Ok(None);
        };
        let path = presentation.capture_path(smallvec![PathSeg::Slide(target)], None)?;
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
        let resolved = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let shape_id = presentation.checked_shape_id(&self.path)?;
        presentation
            .inner
            .set_shape_target_slide(slide_index(&resolved)?, shape_id, target)
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
        let resolved = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let shape_id = presentation.checked_shape_id(&self.path)?;
        presentation
            .inner
            .shape_hyperlink_address_at(shape_owner(&resolved)?, shape_id)
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
        let resolved = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let shape_id = presentation.checked_shape_id(&self.path)?;
        presentation
            .inner
            .set_shape_hyperlink_at(
                shape_owner(&resolved)?,
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

struct CurrentShapeSelection {
    slide: rpptx::ShapeOwner,
    group: Vec<usize>,
    parent: smallvec::SmallVec<[PathSeg; 5]>,
    indices: Vec<usize>,
}

impl PyShapeCollection {
    fn selected_current_indices(
        &self,
        py: Python<'_>,
        shapes: &Bound<'_, PyAny>,
    ) -> PyResult<CurrentShapeSelection> {
        let mut handles = Vec::new();
        for value in shapes.try_iter()? {
            let shape = value?.extract::<Py<PyShape>>()?;
            if !shape.borrow(py).presentation.is(&self.presentation) {
                return Err(PyValueError::new_err(
                    "shape owner is not in this presentation collection",
                ));
            }
            handles.try_reserve(1).map_err(|_| {
                pyo3::exceptions::PyMemoryError::new_err("shape selection allocation failed")
            })?;
            handles.push(shape);
        }
        let (slide, group) = self.target(py)?;
        let presentation = self.presentation.borrow(py);
        let parent = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("collection owner absent"))?
            .segs;
        let mut indices = Vec::new();
        indices.try_reserve(handles.len()).map_err(|_| {
            pyo3::exceptions::PyMemoryError::new_err("shape selection allocation failed")
        })?;
        for shape in handles {
            let shape = shape.borrow(py);
            shape.validate(py)?;
            let current = presentation
                .current_path(&shape.path)
                .ok_or_else(|| PyIndexError::new_err("shape owner absent"))?;
            let Some((PathSeg::Shape(index), prefix)) = current.segs.split_last() else {
                return Err(PyValueError::new_err("shape has no physical member path"));
            };
            if prefix != parent.as_slice() {
                return Err(PyValueError::new_err(
                    "shape owner is not in this presentation collection",
                ));
            }
            indices.push(*index);
        }
        indices.sort_unstable();
        if indices.windows(2).any(|p| p[0] == p[1]) {
            return Err(PyValueError::new_err("shape selection contains duplicates"));
        }
        Ok(CurrentShapeSelection {
            slide,
            group,
            parent,
            indices,
        })
    }

    pub(crate) fn new(presentation: Py<PyPresentation>, path: ContentPath) -> Self {
        Self { presentation, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<rpptx::ShapeOwner> {
        let presentation = self.presentation.borrow(py);
        validate_path(py, &presentation, &self.path, "shape collection", ".shapes")?;
        shape_owner(
            &presentation
                .current_path(&self.path)
                .ok_or_else(|| PyIndexError::new_err("owner is absent"))?,
        )
    }

    fn len(&self, py: Python<'_>) -> PyResult<usize> {
        let owner = self.validate(py)?;
        let presentation = self.presentation.borrow(py);
        if self
            .path
            .segs
            .iter()
            .any(|segment| matches!(segment, PathSeg::Shape(_)))
        {
            return Ok(
                shape_ref_at(&presentation, &self.path).map_or(0, |shape| shape.child_count())
            );
        }
        Ok(match owner {
            rpptx::ShapeOwner::Slide(index) => presentation
                .inner
                .slide(index)
                .map_or(0, |slide| slide.shapes().len()),
            rpptx::ShapeOwner::Layout(index) => {
                presentation.inner.layout(index).map_or(0, |layout| {
                    layout.common_slide_data.shape_tree.children.len()
                })
            }
            rpptx::ShapeOwner::Master(index) => {
                presentation.inner.master(index).map_or(0, |master| {
                    master.common_slide_data.shape_tree.children.len()
                })
            }
        })
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

    /// Returns the slide index and group path that additions to this
    /// collection target, or the error an addition raises.
    fn target(&self, py: Python<'_>) -> PyResult<(rpptx::ShapeOwner, Vec<usize>)> {
        let slide_index = self.validate(py)?;
        let current = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?;
        let group = shape_indices(&current).collect::<Vec<_>>();
        let mut presentation = self.presentation.borrow_mut(py);
        if presentation
            .inner
            .shapes_mut_at(slide_index, &group)
            .is_ok()
        {
            return Ok((slide_index, group));
        }
        let fallback_group = shape_ref_at(&presentation, &self.path)
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
        operation: &'static str,
        add: impl FnOnce(&mut rpptx::ShapesMut<'_>) -> rpptx::Result<()>,
    ) -> PyResult<Py<PyShape>> {
        let (slide_index, group) = self.target(py)?;
        let index = self.len(py)?;
        if operation == "ShapeCollection.add_textbox"
            && group.is_empty()
            && let rpptx::ShapeOwner::Slide(slide) = slide_index
        {
            let mut presentation = self.presentation.borrow_mut(py);
            // The append shortcut must preserve the structural proof's refusal
            // of imported ambiguous identities, without visiting other slides.
            let mut pending = presentation
                .inner
                .slide(slide)
                .ok_or_else(|| PyIndexError::new_err("slide owner is absent"))?
                .shapes()
                .collect::<Vec<_>>();
            let mut ids = std::collections::HashSet::new();
            while let Some(shape) = pending.pop() {
                if shape.kind() != rpptx::ShapeKind::AlternateContent {
                    let id = shape.non_visual_id().ok_or_else(|| {
                        PyValueError::new_err("structural owner proof requires a shape id")
                    })?;
                    if !ids.insert(id) {
                        return Err(PyValueError::new_err(
                            "structural owner proof has an ambiguous shape id",
                        ));
                    }
                }
                pending.extend(shape.children());
            }
            let mut prepared = presentation.prepare_appended_textbox(operation, slide, index)?;
            let mut segments = presentation
                .current_path(&self.path)
                .ok_or_else(|| PyIndexError::new_err("owner is absent"))?
                .segs;
            segments.push(PathSeg::Shape(index));
            let path = presentation.capture_prepared_path(&prepared, segments, None)?;
            let result = Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))?;
            {
                let mut shapes = presentation
                    .inner
                    .shapes_mut_at(slide_index, &group)
                    .map_err(|error| rpptx_to_pyerr(py, error))?;
                add(&mut shapes).map_err(|error| rpptx_to_pyerr(py, error))?;
            }
            presentation.bind_appended_textbox(&mut prepared, slide, index);
            presentation.commit_prepared(prepared)?;
            return Ok(result);
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        {
            let mut shapes = candidate
                .shapes_mut_at(slide_index, &group)
                .map_err(|error| rpptx_to_pyerr(py, error))?;
            add(&mut shapes).map_err(|error| rpptx_to_pyerr(py, error))?;
        }
        let survivors = presentation.unchanged_owners()?;
        let prepared =
            presentation.prepare_owners(&candidate, operation, &survivors, false, false, false)?;
        let mut segments = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?
            .segs;
        segments.push(PathSeg::Shape(index));
        let path = presentation.capture_prepared_path(&prepared, segments, None)?;
        let result = Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    fn item(&self, py: Python<'_>, index: usize) -> PyResult<Py<PyShape>> {
        let mut segments = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?
            .segs;
        segments.push(PathSeg::Shape(index));
        let path = self.presentation.borrow(py).capture_path(segments, None)?;
        Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))
    }
}

#[pymethods]
impl PyShapeCollection {
    /// Adds a video, or an audio clip for an `audio/` type, played on click,
    /// with `poster_frame_image` shown until it plays, as python-pptx
    /// `add_movie` does. Without a poster a black frame shows. `mime_type`
    /// defaults to the one the file extension names.
    #[pyo3(signature = (movie_file, left, top, width, height, poster_frame_image = None, mime_type = None))]
    #[allow(clippy::too_many_arguments)]
    fn add_movie(
        &mut self,
        py: Python<'_>,
        movie_file: &Bound<'_, PyAny>,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
        poster_frame_image: Option<&Bound<'_, PyAny>>,
        mime_type: Option<&str>,
    ) -> PyResult<Py<PyShape>> {
        let (owner, group) = self.target(py)?;
        let rpptx::ShapeOwner::Slide(slide_index) = owner else {
            return Err(PyValueError::new_err("movies require a slide owner"));
        };
        if !group.is_empty() {
            return Err(PyValueError::new_err(
                "a movie goes on the slide's own shapes, not in a group",
            ));
        }
        let media = MediaFile::read(movie_file, mime_type)?;
        let (poster, poster_name) = match poster_frame_image {
            Some(image) => image_bytes(image)?,
            None => (BLACK_POSTER.to_vec(), "poster.png".to_owned()),
        };
        let index = self.len(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .add_media(
                slide_index,
                media.kind(),
                media.input(),
                rpptx::MediaPoster {
                    bytes: &poster,
                    filename: &poster_name,
                },
                rpptx::Emu(left),
                rpptx::Emu(top),
                rpptx::Emu(width),
                rpptx::Emu(height),
                rpptx::MediaPlaybackSettings::default(),
            )
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.add_movie",
            &survivors,
            false,
            false,
            false,
        )?;
        let mut segments = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("owner is absent"))?
            .segs;
        segments.push(PathSeg::Shape(index));
        let path = presentation.capture_prepared_path(&prepared, segments, None)?;
        let result = Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    fn group(&self, py: Python<'_>, shapes: &Bound<'_, PyAny>) -> PyResult<Py<PyShape>> {
        let CurrentShapeSelection {
            slide,
            group,
            parent,
            indices,
        } = self.selected_current_indices(py, shapes)?;
        let start = *indices
            .first()
            .ok_or_else(|| PyValueError::new_err("empty group selection"))?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .shapes_mut_at(slide, &group)
            .map_err(|_| PyValueError::new_err("missing group"))?
            .group(&indices)
            .map(drop)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        let survivors = presentation.grouped_shape_owners(&parent, start, indices.len())?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.group",
            &survivors,
            false,
            false,
            false,
        )?;
        let mut destination = parent;
        destination.push(PathSeg::Shape(start));
        let path = presentation.capture_prepared_path(&prepared, destination, None)?;
        let result = Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }
    #[pyo3(signature=(shapes,alignment,reference=None,*,relative_to=None))]
    fn align(
        &self,
        py: Python<'_>,
        shapes: &Bound<'_, PyAny>,
        alignment: &str,
        reference: Option<&str>,
        relative_to: Option<&str>,
    ) -> PyResult<()> {
        let alignment = match alignment {
            "left" => rpptx::ShapeAlignment::Left,
            "center" => rpptx::ShapeAlignment::Center,
            "right" => rpptx::ShapeAlignment::Right,
            "top" => rpptx::ShapeAlignment::Top,
            "middle" => rpptx::ShapeAlignment::Middle,
            "bottom" => rpptx::ShapeAlignment::Bottom,
            _ => return Err(PyValueError::new_err("unknown alignment")),
        };
        if reference.is_some() && relative_to.is_some() && reference != relative_to {
            return Err(PyValueError::new_err("reference and relative_to disagree"));
        }
        let reference = match relative_to.or(reference).unwrap_or("selection") {
            "selection" => rpptx::ArrangeReference::Selection,
            "slide" => rpptx::ArrangeReference::Slide,
            _ => return Err(PyValueError::new_err("unknown reference")),
        };
        let CurrentShapeSelection {
            slide,
            group,
            indices,
            ..
        } = self.selected_current_indices(py, shapes)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .shapes_mut_at(slide, &group)
            .map_err(|_| PyValueError::new_err("missing group"))?
            .align(&indices, alignment, reference)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        if candidate.to_bytes().map_err(|e| rpptx_to_pyerr(py, e))?
            == presentation
                .inner
                .to_bytes()
                .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.align",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }
    #[pyo3(signature=(shapes,direction,reference=None,*,relative_to=None))]
    fn distribute(
        &self,
        py: Python<'_>,
        shapes: &Bound<'_, PyAny>,
        direction: &str,
        reference: Option<&str>,
        relative_to: Option<&str>,
    ) -> PyResult<()> {
        let direction = match direction {
            "horizontal" => rpptx::DistributeDirection::Horizontal,
            "vertical" => rpptx::DistributeDirection::Vertical,
            _ => return Err(PyValueError::new_err("unknown direction")),
        };
        if reference.is_some() && relative_to.is_some() && reference != relative_to {
            return Err(PyValueError::new_err("reference and relative_to disagree"));
        }
        let reference = match relative_to.or(reference).unwrap_or("selection") {
            "selection" => rpptx::ArrangeReference::Selection,
            "slide" => rpptx::ArrangeReference::Slide,
            _ => return Err(PyValueError::new_err("unknown reference")),
        };
        let CurrentShapeSelection {
            slide,
            group,
            indices,
            ..
        } = self.selected_current_indices(py, shapes)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .shapes_mut_at(slide, &group)
            .map_err(|_| PyValueError::new_err("missing group"))?
            .distribute(&indices, direction, reference)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        if candidate.to_bytes().map_err(|e| rpptx_to_pyerr(py, e))?
            == presentation
                .inner
                .to_bytes()
                .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.distribute",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }

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
        self.validate(py)?;
        for index in 0..self.len(py)? {
            let shape = self.item(py, index)?;
            if shape.borrow(py).read(py, |shape| {
                matches!(shape.placeholder_type(), Some("title" | "ctrTitle"))
            })? {
                return Ok(Some(shape));
            }
        }
        Ok(None)
    }

    #[getter]
    fn placeholders(&self, py: Python<'_>) -> PyResult<Py<PyPlaceholderCollection>> {
        self.validate(py)?;
        Py::new(
            py,
            PyPlaceholderCollection::new(self.presentation.clone_ref(py), self.path.clone()),
        )
    }

    #[allow(clippy::too_many_arguments)]
    fn add_chart(
        &self,
        py: Python<'_>,
        chart_type: &Bound<'_, PyAny>,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
        chart_data: &Bound<'_, PyAny>,
    ) -> PyResult<Py<PyShape>> {
        self.require_slide_root()?;
        let (kind, data) = checked_chart_inputs(chart_type, chart_data)?;
        self.validate(py)?;
        let index = self.len(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("chart owner is absent"))?;
        let mut candidate = presentation.inner.clone();
        candidate
            .add_chart(
                slide_index(&current)?,
                kind,
                rpptx::Emu(left),
                rpptx::Emu(top),
                rpptx::Emu(width),
                rpptx::Emu(height),
                &data,
            )
            .map(drop)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.add_chart",
            &survivors,
            false,
            false,
            false,
        )?;
        let mut segments = current.segs;
        segments.push(PathSeg::Shape(index));
        let path = presentation.capture_prepared_path(&prepared, segments, None)?;
        let result = Py::new(py, PyShape::new(self.presentation.clone_ref(py), path))?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(result)
    }

    pub(crate) fn add_textbox(
        &mut self,
        py: Python<'_>,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
    ) -> PyResult<Py<PyShape>> {
        self.add(py, "ShapeCollection.add_textbox", |shapes| {
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
        self.add(py, "ShapeCollection.add_shape", |shapes| {
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
        self.add(py, "ShapeCollection.add_connector", |shapes| {
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
        self.add(py, "ShapeCollection.add_group_shape", |shapes| {
            shapes.add_group_shape().map(drop)
        })
    }

    /// Removes one shape of this slide and the package parts only it used.
    fn remove(&mut self, py: Python<'_>, shape: &Bound<'_, PyAny>) -> PyResult<()> {
        self.require_slide_root()?;
        let owner = self.validate(py)?;
        let parent = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape collection owner is absent"))?;
        let shape = shape.extract::<PyRef<'_, PyShape>>()?;
        if !shape.presentation.is(&self.presentation) {
            return Err(PyValueError::new_err(
                "shape owner is not in this presentation collection",
            ));
        }
        shape.validate(py)?;
        let current = self
            .presentation
            .borrow(py)
            .current_path(&shape.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let shape_index = match current.segs.as_slice() {
            [_, PathSeg::Shape(index)] if shape_owner(&current)? == owner => *index,
            _ => {
                return Err(PyValueError::new_err(
                    "shape owner is not in this presentation collection",
                ));
            }
        };
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .remove_shape_at(owner, shape_index)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = presentation.removed_shape_owners(&parent.segs, shape_index)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.remove",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
        Ok(())
    }

    /// Moves the shape at `from_` so that it ends up at z-order index `to`,
    /// where later shapes draw on top.
    #[pyo3(name = "move")]
    fn move_shape(&mut self, py: Python<'_>, from_: isize, to: isize) -> PyResult<()> {
        self.require_slide_root()?;
        let owner = self.validate(py)?;
        let parent = self
            .presentation
            .borrow(py)
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape collection owner is absent"))?;
        let len = self.len(py)?;
        let from_ = normalize_index(from_, len, "shape")?;
        let to = normalize_index(to, len, "shape")?;
        if from_ == to {
            return Ok(());
        }
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .move_shape_at(owner, from_, to)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        let survivors = presentation.moved_shape_owners(&parent.segs, from_, to)?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "ShapeCollection.move",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)?;
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
        self.add(py, "ShapeCollection.add_table", |shapes| {
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
        self.add(py, "ShapeCollection.add_picture", |shapes| {
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

/// A one-pixel black PNG, the poster of a movie added without one.
const BLACK_POSTER: [u8; 67] = [
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x00, 0x00, 0x00, 0x00, 0x3A, 0x7E, 0x9B,
    0x55, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0x60, 0x00, 0x00, 0x00,
    0x02, 0x00, 0x01, 0xE5, 0x27, 0xDE, 0xFC, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

/// Media file extensions and the content types they name.
const MEDIA_TYPES: [(&str, &str); 14] = [
    ("mp4", "video/mp4"),
    ("m4v", "video/mp4"),
    ("mov", "video/quicktime"),
    ("avi", "video/x-msvideo"),
    ("wmv", "video/x-ms-wmv"),
    ("webm", "video/webm"),
    ("mpg", "video/mpeg"),
    ("mpeg", "video/mpeg"),
    ("mp3", "audio/mpeg"),
    ("m4a", "audio/mp4"),
    ("wav", "audio/wav"),
    ("wma", "audio/x-ms-wma"),
    ("aac", "audio/aac"),
    ("ogg", "audio/ogg"),
];

/// Media bytes from a path, bytes, or a binary file object, with a file
/// name whose extension matches the content type.
pub(crate) struct MediaFile {
    bytes: Vec<u8>,
    filename: String,
    content_type: String,
}

impl MediaFile {
    pub(crate) fn read(file: &Bound<'_, PyAny>, mime_type: Option<&str>) -> PyResult<Self> {
        let (bytes, filename) = image_bytes(file)?;
        let extension = filename
            .rsplit_once('.')
            .map(|(_, extension)| extension.to_ascii_lowercase());
        let content_type = match mime_type {
            Some(mime_type) => mime_type.to_owned(),
            None => extension
                .as_deref()
                .and_then(|extension| MEDIA_TYPES.iter().find(|(known, _)| *known == extension))
                .map(|(_, content_type)| (*content_type).to_owned())
                .ok_or_else(|| {
                    PyValueError::new_err(
                        "pass mime_type, such as \"video/mp4\", for media without a known extension",
                    )
                })?,
        };
        let filename = if extension.is_some() {
            filename
        } else {
            let extension = MEDIA_TYPES
                .iter()
                .find(|(_, known)| *known == content_type)
                .map_or("bin", |(extension, _)| *extension);
            format!("media.{extension}")
        };
        Ok(Self {
            bytes,
            filename,
            content_type,
        })
    }

    pub(crate) fn kind(&self) -> rpptx::MediaKind {
        if self.content_type.starts_with("audio/") {
            rpptx::MediaKind::Audio
        } else {
            rpptx::MediaKind::Video
        }
    }

    pub(crate) fn input(&self) -> rpptx::MediaSourceInput<'_> {
        rpptx::MediaSourceInput::Embedded(rpptx::EmbeddedMediaInput {
            bytes: &self.bytes,
            filename: &self.filename,
            content_type: &self.content_type,
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

    /// Shape indexes in iteration order.
    ///
    /// A slide owner uses ascending placeholder idx, with equal indexes keeping
    /// document order. A layout or master owner keeps document order. Lookup
    /// by idx stays on `__getitem__`.
    fn ordered_placeholder_indices(&self, py: Python<'_>) -> PyResult<Vec<usize>> {
        let presentation = self.presentation.borrow(py);
        validate_path(
            py,
            &presentation,
            &self.path,
            "placeholder collection",
            ".placeholders",
        )?;
        let resolved = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let owner = shape_owner(&resolved)?;
        let by_idx = resolved
            .segs
            .iter()
            .any(|segment| matches!(segment, PathSeg::Slide(_)))
            && !resolved
                .segs
                .iter()
                .any(|segment| matches!(segment, PathSeg::Layout(_) | PathSeg::Master(_)));
        let parent = shape_indices(&resolved).collect::<Vec<_>>();
        drop(presentation);
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        let mut found = Vec::new();
        for index in 0..collection.len(py)? {
            let presentation = self.presentation.borrow(py);
            let mut path = parent.clone();
            path.push(index);
            if let Some(idx) = presentation
                .inner
                .part_shape_at(owner, &path)
                .ok()
                .and_then(|shape| shape.placeholder_idx())
            {
                found.push((index, idx));
            }
        }
        if by_idx {
            found.sort_by(|left, right| left.1.cmp(&right.1).then(left.0.cmp(&right.0)));
        }
        Ok(found.into_iter().map(|(index, _)| index).collect())
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
        let resolved = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("shape owner is absent"))?;
        let owner = shape_owner(&resolved)?;
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        let shape_index = (0..collection.len(py)?)
            .find(|index| {
                presentation
                    .inner
                    .part_shape_at(owner, &[*index])
                    .is_ok_and(|shape| shape.placeholder_idx() == Some(placeholder_idx))
            })
            .ok_or_else(|| PyIndexError::new_err("placeholder index out of range"))?;
        drop(presentation);
        collection.item(py, shape_index)
    }

    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.ordered_placeholder_indices(py)?.len())
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let indices = self.ordered_placeholder_indices(py)?;
        let collection = PyShapeCollection::new(self.presentation.clone_ref(py), self.path.clone());
        let items = PyList::empty(py);
        for index in indices {
            items.append(collection.item(py, index)?)?;
        }
        items.try_iter()
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
        shape_ref_at(&presentation, &self.path)
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
        shape_mut_at(&mut presentation, &self.path)
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

#[pyclass(name = "Chart")]
pub struct PyChart {
    presentation: Py<PyPresentation>,
    path: ContentPath,
}

impl PyChart {
    fn owner(&self, py: Python<'_>) -> PyResult<(usize, u32)> {
        validate_path(
            py,
            &self.presentation.borrow(py),
            &self.path,
            "chart",
            "Re-fetch it with shape.chart.",
        )?;
        let presentation = self.presentation.borrow(py);
        let current = presentation
            .current_path(&self.path)
            .ok_or_else(|| PyIndexError::new_err("chart owner is absent"))?;
        Ok((
            slide_index(&current)?,
            presentation.checked_shape_id(&self.path)?,
        ))
    }
    fn legend(&self, py: Python<'_>) -> PyResult<rpptx::LegendSettings> {
        let (slide, id) = self.owner(py)?;
        self.presentation
            .borrow(py)
            .inner
            .chart_legend(slide, id)
            .map_err(|e| rpptx_to_pyerr(py, e))
    }
    fn set_legend(
        &self,
        py: Python<'_>,
        visible: Option<bool>,
        position: Option<rpptx::LegendPosition>,
        overlay: Option<bool>,
    ) -> PyResult<()> {
        let (slide, id) = self.owner(py)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .set_chart_legend(slide, id, visible, position, overlay)
            .map_err(|e| rpptx_to_pyerr(py, e))?;
        if candidate.to_bytes().map_err(|e| rpptx_to_pyerr(py, e))?
            == presentation
                .inner
                .to_bytes()
                .map_err(|e| rpptx_to_pyerr(py, e))?
        {
            return Ok(());
        }
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Chart legend",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }
    fn data(&self, py: Python<'_>) -> PyResult<(rpptx::ChartKind, rpptx::ChartData)> {
        let (slide, id) = self.owner(py)?;
        self.presentation
            .borrow(py)
            .inner
            .chart_data(slide, id)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
}

#[pymethods]
impl PyChart {
    #[getter]
    fn has_legend(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.legend(py)?.visible)
    }
    #[setter]
    fn set_has_legend(&self, py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<()> {
        if !value.is_instance_of::<pyo3::types::PyBool>() {
            return Err(PyTypeError::new_err("has_legend requires bool"));
        }
        self.set_legend(py, Some(value.extract()?), None, None)
    }
    #[getter]
    fn legend_position(&self, py: Python<'_>) -> PyResult<&'static str> {
        Ok(self.legend(py)?.position.as_str())
    }
    #[setter]
    fn set_legend_position(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let value = rpptx::LegendPosition::parse(value)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
        self.set_legend(py, None, Some(value), None)
    }
    #[getter]
    fn legend_overlay(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.legend(py)?.overlay)
    }
    #[setter]
    fn set_legend_overlay(&self, py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<()> {
        if !value.is_instance_of::<pyo3::types::PyBool>() {
            return Err(PyTypeError::new_err("legend_overlay requires bool"));
        }
        self.set_legend(py, None, None, Some(value.extract()?))
    }
    #[getter]
    fn chart_type(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        let (slide, id) = self.owner(py)?;
        self.presentation
            .borrow(py)
            .inner
            .chart_type(slide, id)
            .map_err(|error| rpptx_to_pyerr(py, error))
    }
    #[getter]
    fn categories<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        PyTuple::new(py, self.data(py)?.1.categories)
    }
    #[getter]
    fn series<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let values = self
            .data(py)?
            .1
            .series
            .into_iter()
            .map(|(name, values)| Py::new(py, PyChartSeries { name, values }))
            .collect::<PyResult<Vec<_>>>()?;
        PyTuple::new(py, values)
    }
    fn replace_data(&self, py: Python<'_>, chart_data: &Bound<'_, PyAny>) -> PyResult<()> {
        let (slide, id) = self.owner(py)?;
        let kind = self.data(py)?.0;
        let chart_type = CHART_TYPES
            .iter()
            .find_map(|(v, k)| (*k == kind).then_some(*v))
            .expect("supported kind");
        let chart_type = chart_type.into_pyobject(py)?;
        let (_, data) = checked_chart_inputs(chart_type.as_any(), chart_data)?;
        let mut presentation = self.presentation.borrow_mut(py);
        let mut candidate = presentation.inner.clone();
        candidate
            .replace_chart_data(slide, id, &data)
            .map_err(|error| rpptx_to_pyerr(py, error))?;
        if candidate
            .to_bytes()
            .map_err(|error| rpptx_to_pyerr(py, error))?
            == presentation
                .inner
                .to_bytes()
                .map_err(|error| rpptx_to_pyerr(py, error))?
        {
            return Ok(());
        }
        let survivors = presentation.unchanged_owners()?;
        let prepared = presentation.prepare_owners(
            &candidate,
            "Chart.replace_data",
            &survivors,
            false,
            false,
            false,
        )?;
        presentation.publish_owners(candidate, prepared)
    }
}

#[pyclass(name = "ChartSeries", frozen, get_all)]
pub struct PyChartSeries {
    name: String,
    values: Vec<f64>,
}
