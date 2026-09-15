use oxml_py_support::ContentPath;
use pyo3::exceptions::{PyIndexError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyAny;

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, paragraph_location};
use crate::{enum_object, length_object, stale_to_pyerr};

pub(crate) fn alignment_from_int(value: i32) -> PyResult<rdocx::Alignment> {
    match value {
        0 => Ok(rdocx::Alignment::Left),
        1 => Ok(rdocx::Alignment::Center),
        2 => Ok(rdocx::Alignment::Right),
        3 => Ok(rdocx::Alignment::Justify),
        _ => Err(PyValueError::new_err("unsupported paragraph alignment")),
    }
}

pub(crate) fn alignment_to_int(value: rdocx::Alignment) -> i32 {
    match value {
        rdocx::Alignment::Left => 0,
        rdocx::Alignment::Center => 1,
        rdocx::Alignment::Right => 2,
        rdocx::Alignment::Justify => 3,
    }
}

fn checked_underline_code(value: i32) -> PyResult<i32> {
    match value {
        0 | 1 | 2 | 3 | 4 | 6 | 7 | 9 | 10 | 11 => Ok(value),
        _ => Err(PyValueError::new_err("unsupported underline style")),
    }
}

pub(crate) struct FontSnapshot {
    name: Option<String>,
    size: Option<f64>,
    color: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<i32>,
    strike: Option<bool>,
    highlight: Option<String>,
    pub(crate) style_id: Option<String>,
}

impl FontSnapshot {
    fn from_run(run: rdocx::RunRef<'_>) -> Self {
        Self {
            name: run.font_name().map(str::to_owned),
            size: run.size(),
            color: run.color().map(str::to_owned),
            bold: run.bold_value(),
            italic: run.italic_value(),
            underline: run.underline_code_value(),
            strike: run.strike_value(),
            highlight: run.highlight(),
            style_id: run.style_id().map(str::to_owned),
        }
    }
}

pub(crate) enum FontUpdate<'a> {
    Name(Option<&'a str>),
    Size(Option<f64>),
    Color(Option<String>),
    Bold(Option<bool>),
    Italic(Option<bool>),
    Underline(Option<i32>),
    Strike(Option<bool>),
    Highlight(Option<&'a str>),
    Style(Option<&'a str>),
}

impl FontUpdate<'_> {
    fn apply(self, run: &mut rdocx::Run<'_>) {
        match self {
            Self::Name(value) => run.set_font_value(value),
            Self::Size(value) => run.set_size_value(value),
            Self::Color(value) => run.set_color_value(value.as_deref()),
            Self::Bold(value) => run.set_bold_value(value),
            Self::Italic(value) => run.set_italic_value(value),
            Self::Underline(value) => {
                let applied = run.set_underline_code_value(value);
                debug_assert!(applied);
            }
            Self::Strike(value) => run.set_strike_value(value),
            Self::Highlight(value) => run.set_highlight_value(value),
            Self::Style(value) => run.set_style_value(value),
        }
    }
}

#[pyclass(name = "Font")]
pub struct PyFont {
    document: Py<PyDocument>,
    path: ContentPath,
}

impl PyFont {
    pub(crate) fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<(ParagraphLocation, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "font",
                "Re-fetch it with paragraph.runs[i].font.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        let run = self
            .path
            .segs
            .iter()
            .find_map(|segment| match segment {
                oxml_py_support::PathSeg::Run(index) => Some(*index),
                _ => None,
            })
            .ok_or_else(|| PyIndexError::new_err("run index is missing"))?;
        Ok((paragraph_location(&self.path)?, run))
    }

    fn snapshot(&self, py: Python<'_>) -> PyResult<FontSnapshot> {
        let (location, run_index) = self.validate(py)?;
        run_snapshot(py, &self.document, location, run_index)
    }

    fn apply(&self, py: Python<'_>, update: FontUpdate<'_>) -> PyResult<()> {
        let (location, run_index) = self.validate(py)?;
        apply_run_update(py, &self.document, location, run_index, update)
    }
}

/// Read the run properties at a run location the caller has validated.
pub(crate) fn run_snapshot(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    run_index: usize,
) -> PyResult<FontSnapshot> {
    let document = document.borrow(py);
    let snapshot = match location {
        ParagraphLocation::Body(index) => document
            .inner
            .paragraph(index)
            .and_then(|paragraph| paragraph.run(run_index).map(FontSnapshot::from_run)),
        ParagraphLocation::Cell {
            table,
            row,
            cell,
            paragraph,
        } => document.inner.table(table).and_then(|table| {
            let cell = table.cell(row, cell)?;
            let paragraph = cell.paragraph(paragraph)?;
            paragraph.run(run_index).map(FontSnapshot::from_run)
        }),
    };
    snapshot.ok_or_else(|| PyIndexError::new_err("run index out of range"))
}

/// Apply one run property update at a run location the caller has validated.
pub(crate) fn apply_run_update(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    run_index: usize,
    update: FontUpdate<'_>,
) -> PyResult<()> {
    let mut document = document.borrow_mut(py);
    match location {
        ParagraphLocation::Body(index) => {
            let mut paragraph = document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            let mut run = paragraph
                .run_mut(run_index)
                .ok_or_else(|| PyIndexError::new_err("run index out of range"))?;
            update.apply(&mut run);
        }
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
            let mut paragraph = cell
                .paragraph_mut(paragraph)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            let mut run = paragraph
                .run_mut(run_index)
                .ok_or_else(|| PyIndexError::new_err("run index out of range"))?;
            update.apply(&mut run);
        }
    }
    Ok(())
}

#[pymethods]
impl PyFont {
    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.name)
    }

    #[setter]
    fn set_name(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(py, FontUpdate::Name(value))
    }

    #[getter]
    fn size(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .size
            .map(|points| length_object(py, rdocx::Length::pt(points)))
            .transpose()
    }

    #[setter]
    fn set_size(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            FontUpdate::Size(value.map(|emu| rdocx::Length::emu(emu).to_pt())),
        )
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let Some(value) = self.snapshot(py)?.color else {
            return Ok(None);
        };
        if value.eq_ignore_ascii_case("auto") {
            return Ok(None);
        }
        py.import("rdocx")?
            .getattr("RGBColor")?
            .call_method1("from_string", (value,))
            .map(Bound::unbind)
            .map(Some)
    }

    #[setter]
    fn set_color(&self, py: Python<'_>, value: Option<(u8, u8, u8)>) -> PyResult<()> {
        self.apply(
            py,
            FontUpdate::Color(value.map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}"))),
        )
    }

    #[getter]
    fn bold(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.bold)
    }

    #[setter]
    fn set_bold(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Bold(value))
    }

    #[getter]
    fn italic(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.italic)
    }

    #[setter]
    fn set_italic(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Italic(value))
    }

    #[getter]
    fn underline(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        match self.snapshot(py)?.underline {
            None => Ok(None),
            Some(0) => Ok(Some(
                false.into_pyobject(py)?.to_owned().unbind().into_any(),
            )),
            Some(1) => Ok(Some(true.into_pyobject(py)?.to_owned().unbind().into_any())),
            Some(code) => enum_object(py, "WD_UNDERLINE", code).map(Some),
        }
    }

    #[setter]
    fn set_underline(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let style = match value {
            None => None,
            Some(value) if value.is_none() => None,
            Some(value) if value.is_instance_of::<pyo3::types::PyBool>() => {
                Some(if value.extract::<bool>()? { 1 } else { 0 })
            }
            Some(value) => Some(checked_underline_code(value.extract::<i32>()?)?),
        };
        self.apply(py, FontUpdate::Underline(style))
    }

    #[getter]
    fn strike(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.strike)
    }

    #[setter]
    fn set_strike(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Strike(value))
    }

    #[getter]
    fn highlight(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.highlight)
    }

    #[setter]
    fn set_highlight(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        if let Some(value) = value
            && !(value.eq_ignore_ascii_case("auto")
                || (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit())))
        {
            return Err(PyValueError::new_err(
                "highlight must be six hexadecimal digits or auto",
            ));
        }
        self.apply(py, FontUpdate::Highlight(value))
    }
}

pub(crate) struct ParagraphSnapshot {
    pub(crate) style_id: Option<String>,
    pub(crate) numbering: Option<(u32, u32)>,
    alignment: Option<rdocx::Alignment>,
    space_before: Option<rdocx::Length>,
    space_after: Option<rdocx::Length>,
    left_indent: Option<rdocx::Length>,
    right_indent: Option<rdocx::Length>,
    first_line_indent: Option<rdocx::Length>,
    line_spacing: Option<rdocx::Length>,
    line_spacing_multiple: Option<f64>,
    keep_with_next: Option<bool>,
    keep_together: Option<bool>,
    page_break_before: Option<bool>,
    widow_control: Option<bool>,
}

impl ParagraphSnapshot {
    fn from_paragraph(paragraph: rdocx::ParagraphRef<'_>) -> Self {
        Self {
            style_id: paragraph.style_id().map(str::to_owned),
            numbering: paragraph.numbering(),
            alignment: paragraph.alignment(),
            space_before: paragraph.space_before(),
            space_after: paragraph.space_after(),
            left_indent: paragraph.indent_left(),
            right_indent: paragraph.indent_right(),
            first_line_indent: paragraph.first_line_indent(),
            line_spacing: paragraph.line_spacing(),
            line_spacing_multiple: paragraph.line_spacing_multiple(),
            keep_with_next: paragraph.keep_with_next_value(),
            keep_together: paragraph.keep_together_value(),
            page_break_before: paragraph.page_break_before_value(),
            widow_control: paragraph.widow_control_value(),
        }
    }
}

pub(crate) enum ParagraphUpdate {
    Style(Option<String>),
    Numbering(Option<(u32, u32)>),
    Alignment(Option<rdocx::Alignment>),
    SpaceBefore(Option<rdocx::Length>),
    SpaceAfter(Option<rdocx::Length>),
    LeftIndent(Option<rdocx::Length>),
    RightIndent(Option<rdocx::Length>),
    FirstLineIndent(Option<rdocx::Length>),
    LineSpacing(Option<rdocx::Length>),
    LineSpacingMultiple(f64),
    KeepWithNext(Option<bool>),
    KeepTogether(Option<bool>),
    PageBreakBefore(Option<bool>),
    WidowControl(Option<bool>),
}

impl ParagraphUpdate {
    fn apply(self, paragraph: &mut rdocx::Paragraph<'_>) {
        match self {
            Self::Style(value) => paragraph.set_style_value(value.as_deref()),
            Self::Numbering(value) => {
                let applied = paragraph.set_numbering_value(value);
                debug_assert!(applied);
            }
            Self::Alignment(value) => paragraph.set_alignment_value(value),
            Self::SpaceBefore(value) => paragraph.set_space_before_value(value),
            Self::SpaceAfter(value) => paragraph.set_space_after_value(value),
            Self::LeftIndent(value) => paragraph.set_indent_left_value(value),
            Self::RightIndent(value) => paragraph.set_indent_right_value(value),
            Self::FirstLineIndent(value) => paragraph.set_signed_first_line_indent_value(value),
            Self::LineSpacing(Some(value)) => paragraph.set_line_spacing(value.to_pt()),
            Self::LineSpacing(None) => paragraph.clear_line_spacing(),
            Self::LineSpacingMultiple(value) => paragraph.set_line_spacing_multiple(value),
            Self::KeepWithNext(value) => paragraph.set_keep_with_next_value(value),
            Self::KeepTogether(value) => paragraph.set_keep_together_value(value),
            Self::PageBreakBefore(value) => paragraph.set_page_break_before_value(value),
            Self::WidowControl(value) => paragraph.set_widow_control_value(value),
        }
    }
}

#[pyclass(name = "ParagraphFormat")]
pub struct PyParagraphFormat {
    document: Py<PyDocument>,
    path: ContentPath,
}

impl PyParagraphFormat {
    pub(crate) fn new(document: Py<PyDocument>, path: ContentPath) -> Self {
        Self { document, path }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<ParagraphLocation> {
        let document = self.document.borrow(py);
        self.path
            .validate_revision(
                document.revisions.current(),
                "paragraph format",
                "Re-fetch it with paragraph.paragraph_format.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        paragraph_location(&self.path)
    }

    fn snapshot(&self, py: Python<'_>) -> PyResult<ParagraphSnapshot> {
        let location = self.validate(py)?;
        paragraph_snapshot(py, &self.document, location)
    }

    fn apply(&self, py: Python<'_>, update: ParagraphUpdate) -> PyResult<()> {
        let location = self.validate(py)?;
        apply_paragraph_update(py, &self.document, location, update)
    }
}

/// Read the paragraph properties at a location the caller has validated.
pub(crate) fn paragraph_snapshot(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
) -> PyResult<ParagraphSnapshot> {
    let document = document.borrow(py);
    let snapshot = match location {
        ParagraphLocation::Body(index) => document
            .inner
            .paragraph(index)
            .map(ParagraphSnapshot::from_paragraph),
        ParagraphLocation::Cell {
            table,
            row,
            cell,
            paragraph,
        } => document.inner.table(table).and_then(|table| {
            let cell = table.cell(row, cell)?;
            cell.paragraph(paragraph)
                .map(ParagraphSnapshot::from_paragraph)
        }),
    };
    snapshot.ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))
}

/// Apply one paragraph property update at a location the caller has validated.
pub(crate) fn apply_paragraph_update(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    update: ParagraphUpdate,
) -> PyResult<()> {
    let mut document = document.borrow_mut(py);
    match location {
        ParagraphLocation::Body(index) => {
            let mut paragraph = document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            update.apply(&mut paragraph);
        }
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
            let mut paragraph = cell
                .paragraph_mut(paragraph)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            update.apply(&mut paragraph);
        }
    }
    Ok(())
}

#[pymethods]
impl PyParagraphFormat {
    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .alignment
            .map(|value| enum_object(py, "WD_ALIGN_PARAGRAPH", alignment_to_int(value)))
            .transpose()
    }

    #[setter]
    fn set_alignment(&self, py: Python<'_>, value: Option<i32>) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::Alignment(value.map(alignment_from_int).transpose()?),
        )
    }

    #[getter]
    fn space_before(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .space_before
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_space_before(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::SpaceBefore(value.map(rdocx::Length::emu)),
        )
    }

    #[getter]
    fn space_after(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .space_after
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_space_after(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::SpaceAfter(value.map(rdocx::Length::emu)),
        )
    }

    #[getter]
    fn left_indent(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .left_indent
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_left_indent(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::LeftIndent(value.map(rdocx::Length::emu)),
        )
    }

    #[getter]
    fn right_indent(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .right_indent
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_right_indent(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::RightIndent(value.map(rdocx::Length::emu)),
        )
    }

    #[getter]
    fn first_line_indent(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .first_line_indent
            .map(|value| length_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_first_line_indent(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::FirstLineIndent(value.map(rdocx::Length::emu)),
        )
    }

    #[getter]
    fn line_spacing(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let snapshot = self.snapshot(py)?;
        if let Some(value) = snapshot.line_spacing {
            return length_object(py, value).map(Some);
        }
        snapshot
            .line_spacing_multiple
            .map(|value| {
                value
                    .into_pyobject(py)
                    .map(|value| value.to_owned().unbind().into_any())
            })
            .transpose()
            .map_err(Into::into)
    }

    #[setter]
    fn set_line_spacing(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        match value {
            None => self.apply(py, ParagraphUpdate::LineSpacing(None)),
            Some(value) if value.is_none() => self.apply(py, ParagraphUpdate::LineSpacing(None)),
            Some(value) if value.is_instance_of::<pyo3::types::PyFloat>() => {
                self.apply(py, ParagraphUpdate::LineSpacingMultiple(value.extract()?))
            }
            Some(value) => self.apply(
                py,
                ParagraphUpdate::LineSpacing(Some(rdocx::Length::emu(value.extract()?))),
            ),
        }
    }

    #[getter]
    fn keep_with_next(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.keep_with_next)
    }
    #[setter]
    fn set_keep_with_next(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, ParagraphUpdate::KeepWithNext(value))
    }
    #[getter]
    fn keep_together(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.keep_together)
    }
    #[setter]
    fn set_keep_together(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, ParagraphUpdate::KeepTogether(value))
    }
    #[getter]
    fn page_break_before(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.page_break_before)
    }
    #[setter]
    fn set_page_break_before(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, ParagraphUpdate::PageBreakBefore(value))
    }
    #[getter]
    fn widow_control(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.widow_control)
    }
    #[setter]
    fn set_widow_control(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, ParagraphUpdate::WidowControl(value))
    }
}
