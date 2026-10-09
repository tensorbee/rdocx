use oxml_py_support::ContentPath;
use pyo3::exceptions::{PyIndexError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyIterator, PyList};

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, paragraph_location};
use crate::{enum_object, length_object, normalize_index, stale_to_pyerr};

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

fn checked_highlight(value: &str) -> PyResult<&str> {
    match value {
        "black" | "blue" | "cyan" | "darkBlue" | "darkCyan" | "darkGray" | "darkGreen"
        | "darkMagenta" | "darkRed" | "darkYellow" | "green" | "lightGray" | "magenta" | "none"
        | "red" | "white" | "yellow" => Ok(value),
        _ => Err(PyValueError::new_err(
            "highlight must be a Word highlight color name",
        )),
    }
}

fn checked_shading(value: &str) -> PyResult<&str> {
    checked_hex_or_auto("shading", value)
}

/// A colour as the hex-string setters take it: six hex digits or `auto`.
pub(crate) fn checked_hex_or_auto<'a>(param: &str, value: &'a str) -> PyResult<&'a str> {
    if value.eq_ignore_ascii_case("auto") {
        Ok("auto")
    } else if value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Ok(value)
    } else {
        Err(PyValueError::new_err(format!(
            "{param} must be six hexadecimal digits or auto, such as \"FF0000\""
        )))
    }
}

/// A tab stop alignment as `WD_TAB_ALIGNMENT` numbers it.
pub(crate) fn tab_alignment_from_int(value: i32) -> PyResult<rdocx::TabAlignment> {
    match value {
        0 => Ok(rdocx::TabAlignment::Left),
        1 => Ok(rdocx::TabAlignment::Center),
        2 => Ok(rdocx::TabAlignment::Right),
        3 => Ok(rdocx::TabAlignment::Decimal),
        _ => Err(PyValueError::new_err(
            "tab alignment must be WD_TAB_ALIGNMENT.LEFT, CENTER, RIGHT or DECIMAL",
        )),
    }
}

fn tab_alignment_to_int(value: rdocx::TabAlignment) -> i32 {
    match value {
        rdocx::TabAlignment::Left => 0,
        rdocx::TabAlignment::Center => 1,
        rdocx::TabAlignment::Right => 2,
        rdocx::TabAlignment::Decimal => 3,
    }
}

/// A tab leader as `WD_TAB_LEADER` numbers it, `SPACES` writing none.
pub(crate) fn tab_leader_from_int(value: i32) -> PyResult<Option<rdocx::TabLeader>> {
    match value {
        0 => Ok(None),
        1 => Ok(Some(rdocx::TabLeader::Dot)),
        2 => Ok(Some(rdocx::TabLeader::Hyphen)),
        3 => Ok(Some(rdocx::TabLeader::Underscore)),
        _ => Err(PyValueError::new_err(
            "tab leader must be WD_TAB_LEADER.SPACES, DOTS, DASHES or LINES",
        )),
    }
}

fn tab_leader_to_int(value: rdocx::TabLeader) -> i32 {
    match value {
        rdocx::TabLeader::None => 0,
        rdocx::TabLeader::Dot => 1,
        rdocx::TabLeader::Hyphen => 2,
        rdocx::TabLeader::Underscore => 3,
    }
}

/// A paragraph border edge by its `w:pBdr` child name.
pub(crate) fn paragraph_border_edge(value: &str) -> PyResult<rdocx::ParagraphBorderEdge> {
    match value {
        "top" => Ok(rdocx::ParagraphBorderEdge::Top),
        "bottom" => Ok(rdocx::ParagraphBorderEdge::Bottom),
        "left" => Ok(rdocx::ParagraphBorderEdge::Left),
        "right" => Ok(rdocx::ParagraphBorderEdge::Right),
        "between" => Ok(rdocx::ParagraphBorderEdge::Between),
        "bar" => Ok(rdocx::ParagraphBorderEdge::Bar),
        _ => Err(PyValueError::new_err(
            "border edge must be top, bottom, left, right, between or bar",
        )),
    }
}

/// Line spacing as `ParagraphFormat.line_spacing` takes it: a float is a
/// multiple of single spacing and a `Length` an exact height.
pub(crate) enum LineSpacing {
    Multiple(f64),
    Exact(rdocx::Length),
}

pub(crate) fn checked_line_spacing(value: &Bound<'_, PyAny>) -> PyResult<LineSpacing> {
    let spacing = if value.is_instance_of::<pyo3::types::PyFloat>() {
        let multiple = value.extract::<f64>()?;
        (multiple.is_finite() && multiple > 0.0).then_some(LineSpacing::Multiple(multiple))
    } else {
        let emu = value.extract::<i64>()?;
        (emu > 0).then_some(LineSpacing::Exact(rdocx::Length::emu(emu)))
    };
    spacing.ok_or_else(|| {
        PyValueError::new_err(
            "line_spacing must be a positive Length such as Pt(18) or a positive float multiple such as 1.5",
        )
    })
}

/// A paragraph border style and size, `size` in eighths of a point.
pub(crate) fn checked_border(style: &str, size: u32) -> PyResult<rdocx::BorderStyle> {
    let style = crate::table::border_style_from_name(style)?;
    if size > 96 || (style != rdocx::BorderStyle::None && size == 0) {
        return Err(PyValueError::new_err(
            "border size must be 1 to 96 eighths of a point for a visible edge",
        ));
    }
    Ok(style)
}

/// A language tag such as `en-US`: letters, digits and hyphens in segments
/// of one to eight characters, starting with two to eight letters.
fn checked_language<'a>(param: &str, value: Option<&'a str>) -> PyResult<Option<&'a str>> {
    let Some(tag) = value else {
        return Ok(None);
    };
    let mut segments = tag.split('-');
    let primary = segments.next().unwrap_or_default();
    let valid = (2..=8).contains(&primary.len())
        && primary.bytes().all(|byte| byte.is_ascii_alphabetic())
        && segments.all(|segment| {
            (1..=8).contains(&segment.len())
                && segment.bytes().all(|byte| byte.is_ascii_alphanumeric())
        });
    if valid {
        Ok(Some(tag))
    } else {
        Err(PyValueError::new_err(format!(
            "{param} must be a language tag such as \"en-US\" or \"ja-JP\", got {tag:?}"
        )))
    }
}

/// Read a `bool | None` toggle the way python-docx's tri-state properties
/// take it.
fn optional_bool(value: Option<&Bound<'_, PyAny>>) -> PyResult<Option<bool>> {
    match value {
        None => Ok(None),
        Some(value) if value.is_none() => Ok(None),
        Some(value) => value.extract::<bool>().map(Some),
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
    shading: Option<String>,
    pub(crate) style_id: Option<String>,
    vertical_alignment: Option<String>,
    all_caps: Option<bool>,
    small_caps: Option<bool>,
    double_strike: Option<bool>,
    hidden: Option<bool>,
    character_spacing: Option<rdocx::Twips>,
    language: Option<String>,
    east_asian_language: Option<String>,
    complex_script_language: Option<String>,
    east_asian_name: Option<String>,
    complex_script_name: Option<String>,
    rtl: Option<bool>,
}

impl FontSnapshot {
    fn from_run(run: rdocx::RunRef<'_>) -> Self {
        Self {
            vertical_alignment: run.vert_align().map(str::to_owned),
            all_caps: run.all_caps_value(),
            small_caps: run.small_caps_value(),
            double_strike: run.double_strike_value(),
            hidden: run.hidden_value(),
            character_spacing: run.character_spacing(),
            language: run.language().map(str::to_owned),
            east_asian_language: run.language_east_asia().map(str::to_owned),
            complex_script_language: run.language_bidi().map(str::to_owned),
            east_asian_name: run
                .slot_font(rdocx::RunFontSlot::EastAsia)
                .map(str::to_owned),
            complex_script_name: run
                .slot_font(rdocx::RunFontSlot::ComplexScript)
                .map(str::to_owned),
            rtl: run.rtl_value(),
            name: run.font_name().map(str::to_owned),
            size: run.size(),
            color: run.color().map(str::to_owned),
            bold: run.bold_value(),
            italic: run.italic_value(),
            underline: run.underline_code_value(),
            strike: run.strike_value(),
            highlight: run.highlight_color().map(str::to_owned),
            shading: run.shading_fill().map(str::to_owned),
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
    Shading(Option<&'a str>),
    Style(Option<&'a str>),
    VerticalAlignment(Option<rdocx::RunVerticalAlignment>),
    AllCaps(Option<bool>),
    SmallCaps(Option<bool>),
    DoubleStrike(Option<bool>),
    Hidden(Option<bool>),
    CharacterSpacing(Option<rdocx::Length>),
    Language(Option<&'a str>),
    EastAsianLanguage(Option<&'a str>),
    ComplexScriptLanguage(Option<&'a str>),
    SlotFont(rdocx::RunFontSlot, Option<&'a str>),
    Rtl(Option<bool>),
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
            Self::Highlight(value) => {
                let applied = run.set_highlight_value(value);
                debug_assert!(applied);
            }
            Self::Shading(value) => run.set_shading_value(value),
            Self::Style(value) => run.set_style_value(value),
            Self::VerticalAlignment(value) => run.set_vertical_alignment_value(value),
            Self::AllCaps(value) => run.set_all_caps_value(value),
            Self::SmallCaps(value) => run.set_small_caps_value(value),
            Self::DoubleStrike(value) => run.set_double_strike_value(value),
            Self::Hidden(value) => run.set_hidden_value(value),
            Self::CharacterSpacing(value) => run.set_character_spacing_value(value),
            Self::Language(value) => run.set_language_value(value),
            Self::EastAsianLanguage(value) => run.set_language_east_asia_value(value),
            Self::ComplexScriptLanguage(value) => run.set_language_bidi_value(value),
            Self::SlotFont(slot, value) => run.set_slot_font(slot, value),
            Self::Rtl(value) => run.set_rtl_value(value),
        }
    }
}

#[pyclass(name = "Font")]
pub struct PyFont {
    document: Py<PyDocument>,
    path: ContentPath,
    run_path: rdocx::AcceptedRunPath,
}

impl PyFont {
    pub(crate) fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        run_path: rdocx::AcceptedRunPath,
    ) -> Self {
        Self {
            document,
            path,
            run_path,
        }
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
        let (location, _) = self.validate(py)?;
        apply_run_update(py, &self.document, location, &self.run_path, update)
    }

    /// Read `superscript` or `subscript` as python-docx does: `None` without
    /// `w:vertAlign`, otherwise whether it holds `wanted`.
    fn vertical_alignment_is(&self, py: Python<'_>, wanted: &str) -> PyResult<Option<bool>> {
        Ok(self
            .snapshot(py)?
            .vertical_alignment
            .map(|value| value == wanted))
    }

    /// Write `superscript` or `subscript` as python-docx does: `True` sets
    /// `wanted`, `False` removes `w:vertAlign` only when it holds `wanted`,
    /// and `None` removes it.
    fn set_vertical_alignment(
        &self,
        py: Python<'_>,
        wanted: rdocx::RunVerticalAlignment,
        name: &str,
        value: Option<bool>,
    ) -> PyResult<()> {
        let update = match value {
            Some(true) => Some(wanted),
            Some(false) if self.vertical_alignment_is(py, name)? != Some(true) => return Ok(()),
            Some(false) | None => None,
        };
        self.apply(py, FontUpdate::VerticalAlignment(update))
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
    run_path: &rdocx::AcceptedRunPath,
    update: FontUpdate<'_>,
) -> PyResult<()> {
    let mut document = document.borrow_mut(py);
    match location {
        ParagraphLocation::Body(index) => {
            let mut paragraph = document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            paragraph
                .edit_run(run_path, |run| update.apply(run))
                .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
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
            paragraph
                .edit_run(run_path, |run| update.apply(run))
                .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
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
        let value = value.map(checked_highlight).transpose()?;
        self.apply(py, FontUpdate::Highlight(value))
    }

    #[getter]
    fn shading(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.shading)
    }

    #[setter]
    fn set_shading(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let value = value.map(checked_shading).transpose()?;
        self.apply(py, FontUpdate::Shading(value))
    }

    #[getter]
    fn superscript(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.vertical_alignment_is(py, "superscript")
    }

    #[setter]
    fn set_superscript(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.set_vertical_alignment(
            py,
            rdocx::RunVerticalAlignment::Superscript,
            "superscript",
            optional_bool(value)?,
        )
    }

    #[getter]
    fn subscript(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        self.vertical_alignment_is(py, "subscript")
    }

    #[setter]
    fn set_subscript(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        self.set_vertical_alignment(
            py,
            rdocx::RunVerticalAlignment::Subscript,
            "subscript",
            optional_bool(value)?,
        )
    }

    #[getter]
    fn all_caps(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.all_caps)
    }

    #[setter]
    fn set_all_caps(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::AllCaps(value))
    }

    #[getter]
    fn small_caps(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.small_caps)
    }

    #[setter]
    fn set_small_caps(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::SmallCaps(value))
    }

    #[getter]
    fn double_strike(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.double_strike)
    }

    #[setter]
    fn set_double_strike(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::DoubleStrike(value))
    }

    #[getter]
    fn hidden(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.hidden)
    }

    #[setter]
    fn set_hidden(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Hidden(value))
    }

    #[getter]
    fn character_spacing(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .character_spacing
            .map(|twips| length_object(py, rdocx::Length::twips(twips.0)))
            .transpose()
    }

    #[setter]
    fn set_character_spacing(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        self.apply(
            py,
            FontUpdate::CharacterSpacing(value.map(rdocx::Length::emu)),
        )
    }

    #[getter]
    fn language(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.language)
    }

    #[setter]
    fn set_language(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let value = checked_language("language", value)?;
        self.apply(py, FontUpdate::Language(value))
    }

    #[getter]
    fn east_asian_language(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.east_asian_language)
    }

    #[setter]
    fn set_east_asian_language(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let value = checked_language("east_asian_language", value)?;
        self.apply(py, FontUpdate::EastAsianLanguage(value))
    }

    #[getter]
    fn complex_script_language(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.complex_script_language)
    }

    #[setter]
    fn set_complex_script_language(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let value = checked_language("complex_script_language", value)?;
        self.apply(py, FontUpdate::ComplexScriptLanguage(value))
    }

    #[getter]
    fn east_asian_name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.east_asian_name)
    }

    #[setter]
    fn set_east_asian_name(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(
            py,
            FontUpdate::SlotFont(rdocx::RunFontSlot::EastAsia, value),
        )
    }

    #[getter]
    fn complex_script_name(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.complex_script_name)
    }

    #[setter]
    fn set_complex_script_name(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(
            py,
            FontUpdate::SlotFont(rdocx::RunFontSlot::ComplexScript, value),
        )
    }

    #[getter]
    fn rtl(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.rtl)
    }

    #[setter]
    fn set_rtl(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Rtl(value))
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
    read_paragraph(py, document, location, ParagraphSnapshot::from_paragraph)
}

/// Read from the paragraph at a location the caller has validated.
pub(crate) fn read_paragraph<T>(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    read: impl FnOnce(rdocx::ParagraphRef<'_>) -> T,
) -> PyResult<T> {
    let document = document.borrow(py);
    let value = match location {
        ParagraphLocation::Body(index) => document.inner.paragraph(index).map(read),
        ParagraphLocation::Cell {
            table,
            row,
            cell,
            paragraph,
        } => document
            .inner
            .table(table)
            .and_then(|table| table.cell(row, cell)?.paragraph(paragraph).map(read)),
    };
    value.ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))
}

/// Apply one paragraph property update at a location the caller has validated.
pub(crate) fn apply_paragraph_update(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    update: ParagraphUpdate,
) -> PyResult<()> {
    edit_paragraph(py, document, location, |paragraph| update.apply(paragraph))
}

/// Edit the paragraph at a location the caller has validated.
pub(crate) fn edit_paragraph<T>(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    edit: impl FnOnce(&mut rdocx::Paragraph<'_>) -> T,
) -> PyResult<T> {
    let mut document = document.borrow_mut(py);
    match location {
        ParagraphLocation::Body(index) => {
            let mut paragraph = document
                .inner
                .paragraph_mut(index)
                .ok_or_else(|| PyIndexError::new_err("paragraph index out of range"))?;
            Ok(edit(&mut paragraph))
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
            Ok(edit(&mut paragraph))
        }
    }
}

/// One stored tab stop: its position in twips, its alignment (`None` for a
/// `bar`, `clear` or `num` tab) and its leader (`None` for none or one rdocx
/// does not model).
#[derive(Clone, Copy)]
struct TabData {
    twips: i32,
    alignment: Option<rdocx::TabAlignment>,
    leader: Option<rdocx::TabLeader>,
}

impl TabData {
    fn from_ref(tab: rdocx::TabStopRef<'_>) -> Self {
        Self {
            twips: tab.position().as_twips().0,
            alignment: tab.alignment(),
            leader: tab
                .leader()
                .filter(|leader| *leader != rdocx::TabLeader::None),
        }
    }
}

fn tab_alignment_name(value: Option<rdocx::TabAlignment>) -> &'static str {
    match value {
        Some(rdocx::TabAlignment::Left) => "WD_TAB_ALIGNMENT.LEFT",
        Some(rdocx::TabAlignment::Center) => "WD_TAB_ALIGNMENT.CENTER",
        Some(rdocx::TabAlignment::Right) => "WD_TAB_ALIGNMENT.RIGHT",
        Some(rdocx::TabAlignment::Decimal) => "WD_TAB_ALIGNMENT.DECIMAL",
        None => "None",
    }
}

fn tab_leader_name(value: Option<rdocx::TabLeader>) -> &'static str {
    match value {
        None | Some(rdocx::TabLeader::None) => "WD_TAB_LEADER.SPACES",
        Some(rdocx::TabLeader::Dot) => "WD_TAB_LEADER.DOTS",
        Some(rdocx::TabLeader::Hyphen) => "WD_TAB_LEADER.DASHES",
        Some(rdocx::TabLeader::Underscore) => "WD_TAB_LEADER.LINES",
    }
}

/// One tab stop of a paragraph, a live handle as in python-docx.
///
/// It is found by its position, which no other tab stop of the paragraph
/// shares, so adding or deleting another tab stop keeps it valid.
#[pyclass(name = "TabStop")]
pub struct PyTabStop {
    format: PyParagraphFormat,
    twips: i32,
}

impl PyTabStop {
    fn find(&self, py: Python<'_>) -> PyResult<(usize, TabData)> {
        tab_list(py, &self.format)?
            .into_iter()
            .enumerate()
            .find(|(_, tab)| tab.twips == self.twips)
            .ok_or_else(|| {
                PyValueError::new_err(
                    "this tab stop was deleted or moved; re-read paragraph_format.tab_stops",
                )
            })
    }

    fn replace(&self, py: Python<'_>, index: usize, tab: TabData) -> PyResult<()> {
        let Some(alignment) = tab.alignment else {
            return Err(PyValueError::new_err(
                "rdocx does not model this bar, clear or num tab; delete it with del tab_stops[i] and add a new one",
            ));
        };
        let location = self.format.validate(py)?;
        edit_paragraph(py, &self.format.document, location, |paragraph| {
            paragraph.set_tab_stop(
                index,
                alignment,
                rdocx::Length::twips(tab.twips),
                tab.leader,
            )
        })?;
        Ok(())
    }
}

#[pymethods]
impl PyTabStop {
    #[getter]
    fn position(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        length_object(py, rdocx::Length::twips(self.find(py)?.1.twips))
    }

    /// Move the tab stop, keeping the tab stops in position order.
    #[setter]
    fn set_position(&mut self, py: Python<'_>, value: i64) -> PyResult<()> {
        let (index, tab) = self.find(py)?;
        let twips = rdocx::Length::emu(value).as_twips().0;
        if twips == tab.twips {
            return Ok(());
        }
        let tabs = tab_list(py, &self.format)?;
        if tabs.iter().any(|other| other.twips == twips) {
            return Err(duplicate_tab_error());
        }
        let Some(alignment) = tab.alignment else {
            return Err(PyValueError::new_err(
                "rdocx does not model this bar, clear or num tab; delete it with del tab_stops[i] and add a new one",
            ));
        };
        let target = tabs
            .iter()
            .enumerate()
            .filter(|(other, _)| *other != index)
            .filter(|(_, other)| other.twips < twips)
            .count();
        let location = self.format.validate(py)?;
        edit_paragraph(py, &self.format.document, location, |paragraph| {
            paragraph.remove_tab_stop(index);
            paragraph.insert_tab_stop(target, alignment, rdocx::Length::twips(twips), tab.leader)
        })?;
        self.twips = twips;
        Ok(())
    }

    /// `None` for a `bar`, `clear` or `num` tab, which rdocx does not model.
    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.find(py)?
            .1
            .alignment
            .map(|value| enum_object(py, "WD_TAB_ALIGNMENT", tab_alignment_to_int(value)))
            .transpose()
    }

    #[setter]
    fn set_alignment(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let alignment = tab_alignment_from_int(value)?;
        let (index, tab) = self.find(py)?;
        self.replace(
            py,
            index,
            TabData {
                alignment: Some(alignment),
                ..tab
            },
        )
    }

    #[getter]
    fn leader(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let leader = self.find(py)?.1.leader.map_or(0, tab_leader_to_int);
        enum_object(py, "WD_TAB_LEADER", leader)
    }

    #[setter]
    fn set_leader(&self, py: Python<'_>, value: i32) -> PyResult<()> {
        let leader = tab_leader_from_int(value)?;
        let (index, tab) = self.find(py)?;
        self.replace(py, index, TabData { leader, ..tab })
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let tab = self.find(py)?.1;
        Ok(format!(
            "TabStop(position=Twips({}), alignment={}, leader={})",
            tab.twips,
            tab_alignment_name(tab.alignment),
            tab_leader_name(tab.leader)
        ))
    }
}

fn duplicate_tab_error() -> PyErr {
    PyValueError::new_err(
        "a tab stop already exists at this position; change it through tab_stops[i] or delete it with del tab_stops[i]",
    )
}

fn tab_list(py: Python<'_>, format: &PyParagraphFormat) -> PyResult<Vec<TabData>> {
    let location = format.validate(py)?;
    read_paragraph(py, &format.document, location, |paragraph| {
        (0..paragraph.tab_stop_count())
            .filter_map(|index| paragraph.tab_stop(index).map(TabData::from_ref))
            .collect()
    })
}

/// The direct tab stops of one paragraph, as python-docx's
/// `paragraph_format.tab_stops` exposes them.
#[pyclass(name = "TabStops")]
pub struct PyTabStops {
    format: PyParagraphFormat,
}

impl PyTabStops {
    fn handle(&self, py: Python<'_>, tab: TabData) -> PyTabStop {
        PyTabStop {
            format: PyParagraphFormat::new(
                self.format.document.clone_ref(py),
                self.format.path.clone(),
            ),
            twips: tab.twips,
        }
    }
}

#[pymethods]
impl PyTabStops {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(tab_list(py, &self.format)?.len())
    }

    fn __getitem__(&self, py: Python<'_>, index: isize) -> PyResult<PyTabStop> {
        let tabs = tab_list(py, &self.format)?;
        let index = normalize_index(index, tabs.len(), "tab stop")?;
        Ok(self.handle(py, tabs[index]))
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let handles = tab_list(py, &self.format)?
            .into_iter()
            .map(|tab| self.handle(py, tab))
            .collect::<Vec<_>>();
        PyList::new(py, handles)?.try_iter()
    }

    fn __delitem__(&self, py: Python<'_>, index: isize) -> PyResult<()> {
        let count = tab_list(py, &self.format)?.len();
        let index = normalize_index(index, count, "tab stop")?;
        let location = self.format.validate(py)?;
        edit_paragraph(py, &self.format.document, location, |paragraph| {
            paragraph.remove_tab_stop(index)
        })?;
        Ok(())
    }

    /// Add a tab stop in position order, as python-docx does, and return it.
    /// Unlike python-docx, a second tab stop at one position raises.
    #[pyo3(signature = (position, alignment = 0, leader = 0))]
    fn add_tab_stop(
        &self,
        py: Python<'_>,
        position: i64,
        alignment: i32,
        leader: i32,
    ) -> PyResult<PyTabStop> {
        let alignment = tab_alignment_from_int(alignment)?;
        let leader = tab_leader_from_int(leader)?;
        let position = rdocx::Length::emu(position);
        let twips = position.as_twips().0;
        let tabs = tab_list(py, &self.format)?;
        if tabs.iter().any(|tab| tab.twips == twips) {
            return Err(duplicate_tab_error());
        }
        let index = tabs.iter().filter(|tab| tab.twips < twips).count();
        let location = self.format.validate(py)?;
        edit_paragraph(py, &self.format.document, location, |paragraph| {
            paragraph.insert_tab_stop(index, alignment, position, leader)
        })?;
        Ok(self.handle(
            py,
            TabData {
                twips,
                alignment: Some(alignment),
                leader,
            },
        ))
    }

    fn clear_all(&self, py: Python<'_>) -> PyResult<()> {
        let location = self.format.validate(py)?;
        edit_paragraph(py, &self.format.document, location, |paragraph| {
            paragraph.clear_tab_stops()
        })
    }
}

/// A border edge as `(style, size in eighths of a point, color)`, as
/// `Table.border` reads one.
type BorderSnapshot = (String, Option<u32>, Option<String>);

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
        let update = match value.filter(|value| !value.is_none()) {
            None => ParagraphUpdate::LineSpacing(None),
            Some(value) => match checked_line_spacing(value)? {
                LineSpacing::Multiple(multiple) => ParagraphUpdate::LineSpacingMultiple(multiple),
                LineSpacing::Exact(length) => ParagraphUpdate::LineSpacing(Some(length)),
            },
        };
        self.apply(py, update)
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

    #[getter]
    fn tab_stops(&self, py: Python<'_>) -> PyResult<PyTabStops> {
        self.validate(py)?;
        Ok(PyTabStops {
            format: Self::new(self.document.clone_ref(py), self.path.clone()),
        })
    }

    #[getter]
    fn outline_level(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        let location = self.validate(py)?;
        read_paragraph(py, &self.document, location, |paragraph| {
            paragraph.outline_level()
        })
    }

    #[setter]
    fn set_outline_level(&self, py: Python<'_>, value: Option<u32>) -> PyResult<()> {
        if value.is_some_and(|level| level > 9) {
            return Err(PyValueError::new_err("outline_level must be from 0 to 9"));
        }
        let location = self.validate(py)?;
        let applied = edit_paragraph(py, &self.document, location, |paragraph| {
            paragraph.set_outline_level_value(value)
        })?;
        debug_assert!(applied);
        Ok(())
    }

    #[getter]
    fn right_to_left(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        let location = self.validate(py)?;
        read_paragraph(py, &self.document, location, |paragraph| {
            paragraph.right_to_left_value()
        })
    }

    #[setter]
    fn set_right_to_left(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        let location = self.validate(py)?;
        edit_paragraph(py, &self.document, location, |paragraph| {
            paragraph.set_right_to_left_value(value)
        })
    }

    /// The direct shading fill, six hex digits or `auto`.
    #[getter]
    fn shading(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let location = self.validate(py)?;
        read_paragraph(py, &self.document, location, |paragraph| {
            paragraph.shading_fill().map(str::to_owned)
        })
    }

    /// Write a clear `w:shd` with this fill, or remove the shading.
    #[setter]
    fn set_shading(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        let fill = value.map(checked_shading).transpose()?;
        let location = self.validate(py)?;
        edit_paragraph(py, &self.document, location, |paragraph| {
            paragraph.set_shading_value(fill.map(|fill| ("clear", fill, "auto")))
        })
    }

    fn border(&self, py: Python<'_>, edge: &str) -> PyResult<Option<BorderSnapshot>> {
        let edge = paragraph_border_edge(edge)?;
        let location = self.validate(py)?;
        read_paragraph(py, &self.document, location, |paragraph| {
            paragraph.border(edge).map(|border| {
                (
                    border.style().to_owned(),
                    border.size_eighths_pt(),
                    border.color().map(str::to_owned),
                )
            })
        })
    }

    /// Set one border edge, a single line by default as Google Docs reads it.
    #[pyo3(signature = (edge, style = "single", *, size = 4, color = "auto"))]
    fn set_border(
        &self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        size: u32,
        color: &str,
    ) -> PyResult<()> {
        let edge = paragraph_border_edge(edge)?;
        let style = checked_border(style, size)?;
        let color = checked_hex_or_auto("color", color)?;
        let location = self.validate(py)?;
        edit_paragraph(py, &self.document, location, |paragraph| {
            paragraph.set_border_value(edge, Some((style, size, color)))
        })
    }

    fn remove_border(&self, py: Python<'_>, edge: &str) -> PyResult<()> {
        let edge = paragraph_border_edge(edge)?;
        let location = self.validate(py)?;
        edit_paragraph(py, &self.document, location, |paragraph| {
            paragraph.set_border_value(edge, None)
        })
    }

    fn clear_borders(&self, py: Python<'_>) -> PyResult<()> {
        let location = self.validate(py)?;
        edit_paragraph(py, &self.document, location, |paragraph| {
            paragraph.clear_borders()
        })
    }
}
