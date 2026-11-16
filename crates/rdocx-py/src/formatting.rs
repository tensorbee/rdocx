use oxml_py_support::ContentPath;
use pyo3::exceptions::{PyIndexError, PyOverflowError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyAny;

use crate::document::PyDocument;
use crate::paragraph::{ParagraphLocation, paragraph_location};
use crate::{color_hex, color_object, enum_object, length_object, stale_to_pyerr};

type BorderSnapshot = (String, Option<u32>, Option<Py<PyAny>>);

/// Python line multiple checked before [`rdocx::auto_line_spacing`].
///
/// An exact `n / 240` grid point inside signed 32-bit twips is accepted even
/// when the raw product is one ulp outside `n`. Every other multiple must have
/// a finite raw product inside the closed 32-bit range. Non-finite values and
/// a huge int whose float conversion overflows are `ValueError`.
///
/// The Python number is extracted once. Callers publish this checked `f64`.
pub(crate) fn require_auto_line_spacing(value: &Bound<'_, PyAny>) -> PyResult<f64> {
    let multiple = match value.extract::<f64>() {
        Ok(multiple) => multiple,
        Err(error) if error.is_instance_of::<PyOverflowError>(value.py()) => {
            return Err(PyValueError::new_err(
                "line_spacing multiple must be finite and fit signed twips",
            ));
        }
        Err(error) => return Err(error),
    };
    if !multiple.is_finite() {
        return Err(PyValueError::new_err(
            "line_spacing multiple must be finite and fit signed twips",
        ));
    }
    let scaled = multiple * 240.0;
    let nearest = scaled.round();
    let rescued =
        nearest >= i32::MIN as f64 && nearest <= i32::MAX as f64 && nearest / 240.0 == multiple;
    if !rescued && !(scaled.is_finite() && scaled >= i32::MIN as f64 && scaled <= i32::MAX as f64) {
        return Err(PyValueError::new_err(
            "line_spacing multiple must be finite and fit signed twips",
        ));
    }
    Ok(multiple)
}

// These two Word spacing properties share EMU storage but have different signs.
fn checked_spacing_length(
    py: Python<'_>,
    value: Option<&Bound<'_, PyAny>>,
    name: &str,
    signed: bool,
) -> PyResult<Option<rdocx::Length>> {
    value.map(|value| {
        if value.is_instance_of::<pyo3::types::PyBool>()
            || !value.is_instance_of::<pyo3::types::PyInt>()
        {
            return Err(PyTypeError::new_err(format!(
                "{name} requires integer EMU or Length, not bool or other types; use Pt, Twips or Emu"
            )));
        }
        let emu = value.extract::<i64>()?;
        let twips = emu / 635;
        let explicit = value.is_instance(&py.import("rdocx.shared")?.getattr("Length")?)?;
        if emu != 0 && twips == 0 && !explicit {
            return Err(PyValueError::new_err(format!(
                "{name} bare integer EMU would store zero twips; use Pt, Twips or explicit Emu"
            )));
        }
        if !signed && emu < 0 {
            return Err(PyValueError::new_err(format!("{name} must be nonnegative; use Pt or Twips")));
        }
        i32::try_from(twips).map_err(|_| PyValueError::new_err(format!(
            "{name} exceeds signed twip range; use a representable Pt, Twips or Emu Length"
        )))?;
        Ok(rdocx::Length::emu(emu))
    }).transpose()
}

pub(crate) fn tab_alignment_from_int(value: i32) -> PyResult<rdocx::TabAlignment> {
    match value {
        0 => Ok(rdocx::TabAlignment::Left),
        1 => Ok(rdocx::TabAlignment::Center),
        2 => Ok(rdocx::TabAlignment::Right),
        3 => Ok(rdocx::TabAlignment::Decimal),
        4 => Ok(rdocx::TabAlignment::Bar),
        6 => Ok(rdocx::TabAlignment::List),
        7 => Ok(rdocx::TabAlignment::Clear),
        102 => Ok(rdocx::TabAlignment::End),
        103 => Ok(rdocx::TabAlignment::Num),
        104 => Ok(rdocx::TabAlignment::Start),
        _ => Err(PyValueError::new_err("unsupported tab alignment")),
    }
}

fn tab_alignment_to_int(value: rdocx::TabAlignment) -> i32 {
    match value {
        rdocx::TabAlignment::Left => 0,
        rdocx::TabAlignment::Center => 1,
        rdocx::TabAlignment::Right => 2,
        rdocx::TabAlignment::Decimal => 3,
        rdocx::TabAlignment::Bar => 4,
        rdocx::TabAlignment::List => 6,
        rdocx::TabAlignment::Clear => 7,
        rdocx::TabAlignment::End => 102,
        rdocx::TabAlignment::Num => 103,
        rdocx::TabAlignment::Start => 104,
    }
}

pub(crate) fn tab_leader_from_int(value: Option<i32>) -> PyResult<Option<rdocx::TabLeader>> {
    match value {
        None | Some(0) => Ok(None),
        Some(1) => Ok(Some(rdocx::TabLeader::Dot)),
        Some(2) => Ok(Some(rdocx::TabLeader::Hyphen)),
        Some(3) => Ok(Some(rdocx::TabLeader::Underscore)),
        Some(4) => Ok(Some(rdocx::TabLeader::Heavy)),
        Some(5) => Ok(Some(rdocx::TabLeader::MiddleDot)),
        _ => Err(PyValueError::new_err("unsupported tab leader")),
    }
}

fn tab_leader_to_int(value: Option<rdocx::TabLeader>) -> i32 {
    match value {
        None | Some(rdocx::TabLeader::None) => 0,
        Some(rdocx::TabLeader::Dot) => 1,
        Some(rdocx::TabLeader::Hyphen) => 2,
        Some(rdocx::TabLeader::Underscore) => 3,
        Some(rdocx::TabLeader::Heavy) => 4,
        Some(rdocx::TabLeader::MiddleDot) => 5,
    }
}

pub(crate) fn paragraph_border_edge(value: &str) -> PyResult<rdocx::ParagraphBorderEdge> {
    match value {
        "top" => Ok(rdocx::ParagraphBorderEdge::Top),
        "bottom" => Ok(rdocx::ParagraphBorderEdge::Bottom),
        "left" => Ok(rdocx::ParagraphBorderEdge::Left),
        "right" => Ok(rdocx::ParagraphBorderEdge::Right),
        "between" => Ok(rdocx::ParagraphBorderEdge::Between),
        "bar" => Ok(rdocx::ParagraphBorderEdge::Bar),
        _ => Err(PyValueError::new_err(
            "paragraph border edge must be top, bottom, left, right, between or bar",
        )),
    }
}

pub(crate) fn paragraph_color_hex(value: &Bound<'_, PyAny>, parameter: &str) -> PyResult<String> {
    if value
        .extract::<String>()
        .is_ok_and(|value| value.eq_ignore_ascii_case("auto"))
    {
        Ok("auto".to_owned())
    } else {
        color_hex(value, parameter)
    }
}

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

const HIGHLIGHT_COLORS: [&str; 17] = [
    "none",
    "black",
    "blue",
    "cyan",
    "green",
    "magenta",
    "red",
    "yellow",
    "white",
    "darkBlue",
    "darkCyan",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "darkGray",
    "lightGray",
];

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

const THEME_COLORS: [&str; 16] = [
    "dark1",
    "light1",
    "dark2",
    "light2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hyperlink",
    "followedHyperlink",
    "text1",
    "background1",
    "text2",
    "background2",
];

pub(crate) struct FontSnapshot {
    name: Option<String>,
    size: Option<f64>,
    color: Option<String>,
    color_theme: Option<String>,
    color_theme_tint: Option<u8>,
    color_theme_shade: Option<u8>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<i32>,
    strike: Option<bool>,
    highlight: Option<String>,
    shading: Option<String>,
    superscript: Option<bool>,
    subscript: Option<bool>,
    all_caps: Option<bool>,
    small_caps: Option<bool>,
    double_strike: Option<bool>,
    hidden: Option<bool>,
    rtl: Option<bool>,
    language: Option<String>,
    language_east_asia: Option<String>,
    language_bidi: Option<String>,
    name_east_asia: Option<String>,
    name_complex_script: Option<String>,
    character_spacing: Option<rdocx::Length>,
    pub(crate) style_id: Option<String>,
}

impl FontSnapshot {
    fn from_run(run: rdocx::RunRef<'_>) -> Self {
        Self {
            name: run.font_name().map(str::to_owned),
            size: run.size(),
            color: run.color().map(str::to_owned),
            color_theme: run.color_theme().map(str::to_owned),
            color_theme_tint: run.color_theme_tint(),
            color_theme_shade: run.color_theme_shade(),
            bold: run.bold_value(),
            italic: run.italic_value(),
            underline: run.underline_code_value(),
            strike: run.strike_value(),
            highlight: run.highlight_color().map(str::to_owned),
            shading: run.shading_fill().map(str::to_owned),
            superscript: run.vert_align().map(|value| value == "superscript"),
            subscript: run.vert_align().map(|value| value == "subscript"),
            all_caps: run.all_caps_value(),
            small_caps: run.small_caps_value(),
            double_strike: run.double_strike_value(),
            hidden: run.hidden_value(),
            rtl: run.rtl_value(),
            language: run.language().map(str::to_owned),
            language_east_asia: run.language_east_asia().map(str::to_owned),
            language_bidi: run.language_bidi().map(str::to_owned),
            name_east_asia: run
                .slot_font(rdocx::RunFontSlot::EastAsia)
                .map(str::to_owned),
            name_complex_script: run
                .slot_font(rdocx::RunFontSlot::ComplexScript)
                .map(str::to_owned),
            character_spacing: run
                .character_spacing()
                .map(|value| rdocx::Length::twips(value.0)),
            style_id: run.style_id().map(str::to_owned),
        }
    }
}

pub(crate) enum FontUpdate<'a> {
    Name(Option<&'a str>),
    Size(Option<f64>),
    Color(Option<String>),
    ColorTheme(Option<String>, Option<u8>, Option<u8>),
    Bold(Option<bool>),
    Italic(Option<bool>),
    Underline(Option<i32>),
    Strike(Option<bool>),
    Highlight(Option<&'a str>),
    Shading(Option<&'a str>),
    Superscript(Option<bool>),
    Subscript(Option<bool>),
    AllCaps(Option<bool>),
    SmallCaps(Option<bool>),
    DoubleStrike(Option<bool>),
    Hidden(Option<bool>),
    Rtl(Option<bool>),
    Language(Option<&'a str>),
    LanguageEastAsia(Option<&'a str>),
    LanguageBidi(Option<&'a str>),
    NameEastAsia(Option<&'a str>),
    NameComplexScript(Option<&'a str>),
    CharacterSpacing(Option<rdocx::Length>),
    Style(Option<&'a str>),
}

impl FontUpdate<'_> {
    fn apply(self, run: &mut rdocx::Run<'_>) -> rdocx::Result<()> {
        match self {
            Self::Name(value) => run.set_font_value(value),
            Self::Size(value) => run.set_size_value(value),
            Self::Color(value) => run.set_color_value(value.as_deref()),
            Self::ColorTheme(value, tint, shade) => {
                run.set_color_theme(value.as_deref(), tint, shade)
            }
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
            Self::Superscript(value) => run.set_superscript_value(value),
            Self::Subscript(value) => run.set_subscript_value(value),
            Self::AllCaps(value) => run.set_all_caps_value(value),
            Self::SmallCaps(value) => run.set_small_caps_value(value),
            Self::DoubleStrike(value) => run.set_double_strike_value(value),
            Self::Hidden(value) => run.set_hidden_value(value),
            Self::Rtl(value) => run.set_rtl_value(value),
            Self::Language(value) => run.set_language_value(value),
            Self::LanguageEastAsia(value) => run.set_language_east_asia_value(value),
            Self::LanguageBidi(value) => run.set_language_bidi_value(value),
            Self::NameEastAsia(value) => run.set_slot_font(rdocx::RunFontSlot::EastAsia, value),
            Self::NameComplexScript(value) => {
                run.set_slot_font(rdocx::RunFontSlot::ComplexScript, value)
            }
            Self::CharacterSpacing(value) => return run.try_set_character_spacing_value(value),
            Self::Style(value) => run.set_style_value(value),
        }
        Ok(())
    }
}

#[pyclass(name = "Font")]
pub struct PyFont {
    document: Py<PyDocument>,
    path: ContentPath,
    run_path: rdocx::AcceptedRunPath,
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
}

impl PyFont {
    pub(crate) fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        run_path: rdocx::AcceptedRunPath,
        related: Option<rdocx::ContentLocation>,
        owner: Option<oxml_py_support::RevisionState>,
    ) -> Self {
        Self {
            document,
            path,
            run_path,
            related,
            owner,
        }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<(ParagraphLocation, usize)> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
                "font",
                "Re-fetch it with paragraph.runs[i].font.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        let location = match &self.related {
            Some(location) => match &self.owner {
                Some(owner) => ParagraphLocation::Related(document.current_run_owner(owner)?.0),
                None => ParagraphLocation::Related(location.clone()),
            },
            None => paragraph_location(&self.path)?,
        };
        let index = location
            .read(&document.inner, |paragraph| {
                (0..paragraph.run_count())
                    .find(|index| paragraph.run_path(*index).as_ref() == Some(&self.run_path))
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?
            .ok_or_else(|| PyIndexError::new_err("font run physical owner is no longer present"))?;
        Ok((location, index))
    }

    fn snapshot(&self, py: Python<'_>) -> PyResult<FontSnapshot> {
        let (location, run_index) = self.validate(py)?;
        run_snapshot(py, &self.document, location, run_index)
    }

    fn apply(&self, py: Python<'_>, update: FontUpdate<'_>) -> PyResult<()> {
        let (location, _) = self.validate(py)?;
        apply_run_update(py, &self.document, location, &self.run_path, update)
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
    location
        .read(&document.inner, |paragraph| {
            paragraph.run(run_index).map(FontSnapshot::from_run)
        })
        .map_err(|error| crate::rdocx_to_pyerr(py, error))?
        .ok_or_else(|| PyIndexError::new_err("run index out of range"))
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
    location
        .edit(&mut document.inner, |paragraph| {
            let mut result = Ok(());
            paragraph.edit_run(run_path, |run| result = update.apply(run))?;
            result
        })
        .map_err(|error| crate::rdocx_to_pyerr(py, error))
}

#[pymethods]
impl PyFont {
    #[getter]
    fn superscript(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.superscript)
    }
    #[setter]
    fn set_superscript(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Superscript(value))
    }

    #[getter]
    fn subscript(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.subscript)
    }
    #[setter]
    fn set_subscript(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Subscript(value))
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
    fn rtl(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.rtl)
    }
    #[setter]
    fn set_rtl(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Rtl(value))
    }

    #[getter]
    fn language(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.language)
    }
    #[setter]
    fn set_language(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        if let Some(value) = value
            && (!value
                .as_bytes()
                .first()
                .is_some_and(|byte| byte.is_ascii_alphabetic())
                || !value.split('-').all(|subtag| {
                    !subtag.is_empty() && subtag.bytes().all(|byte| byte.is_ascii_alphanumeric())
                }))
        {
            return Err(PyValueError::new_err(
                "language tag must begin with an ASCII letter and contain nonempty ASCII alphanumeric subtags separated by single hyphens",
            ));
        }
        self.apply(py, FontUpdate::Language(value))
    }

    #[getter]
    fn language_east_asia(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.language_east_asia)
    }
    #[setter]
    fn set_language_east_asia(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(py, FontUpdate::LanguageEastAsia(value))
    }

    #[getter]
    fn language_bidi(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.language_bidi)
    }
    #[setter]
    fn set_language_bidi(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(py, FontUpdate::LanguageBidi(value))
    }

    #[getter]
    fn name_east_asia(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.name_east_asia)
    }
    #[setter]
    fn set_name_east_asia(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(py, FontUpdate::NameEastAsia(value))
    }

    #[getter]
    fn name_complex_script(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.snapshot(py)?.name_complex_script)
    }
    #[setter]
    fn set_name_complex_script(&self, py: Python<'_>, value: Option<&str>) -> PyResult<()> {
        self.apply(py, FontUpdate::NameComplexScript(value))
    }

    #[getter]
    fn character_spacing(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .character_spacing
            .map(|value| length_object(py, value))
            .transpose()
    }
    #[setter]
    fn set_character_spacing(
        &self,
        py: Python<'_>,
        value: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let value = checked_spacing_length(py, value, "Font.character_spacing", true)?;
        self.apply(py, FontUpdate::CharacterSpacing(value))
    }

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
    fn set_size(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let points = value.map(|value| -> PyResult<f64> {
            if value.is_instance_of::<pyo3::types::PyBool>() {
                return Err(PyTypeError::new_err("font.size rejects bool as EMU; use rdocx.shared.Pt(12)"));
            }
            let emu = value.extract::<i64>()?;
            let points = rdocx::Length::emu(emu).to_pt();
            if emu != 0 && (points * 2.0) as u32 == 0 {
                return Err(PyValueError::new_err("font.size bare int is EMU; Word stores 0.5 point steps; use rdocx.shared.Pt(12) or a Length at least Pt(0.5)"));
            }
            Ok(points)
        }).transpose()?;
        self.apply(py, FontUpdate::Size(points))
    }

    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Py<PyColorFormat>> {
        self.snapshot(py)?;
        Py::new(
            py,
            PyColorFormat {
                document: self.document.clone_ref(py),
                path: self.path.clone(),
                run_path: self.run_path.clone(),
                related: self.related.clone(),
                owner: self.owner.clone(),
            },
        )
    }

    #[setter]
    fn set_color(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let value = value
            .filter(|value| !value.is_none())
            .map(|value| color_hex(value, "Font.color"))
            .transpose()?;
        self.apply(py, FontUpdate::Color(value))
    }

    #[getter]
    pub(crate) fn bold(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.bold)
    }

    #[setter]
    pub(crate) fn set_bold(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Bold(value))
    }

    #[getter]
    pub(crate) fn italic(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.italic)
    }

    #[setter]
    pub(crate) fn set_italic(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, FontUpdate::Italic(value))
    }

    #[getter]
    pub(crate) fn underline(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
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
    pub(crate) fn set_underline(
        &self,
        py: Python<'_>,
        value: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
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
    fn highlight_color(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .highlight
            .map(|value| {
                let index = HIGHLIGHT_COLORS
                    .iter()
                    .position(|name| *name == value)
                    .ok_or_else(|| {
                        PyValueError::new_err("use font.highlight for this retained color")
                    })?;
                py.import("rdocx.enum.text")?
                    .getattr("WD_COLOR_INDEX")?
                    .call1((index,))
                    .map(Bound::unbind)
            })
            .transpose()
    }

    #[setter]
    fn set_highlight_color(
        &self,
        py: Python<'_>,
        value: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        let color = value.filter(|v| !v.is_none()).map(|value| {
            if value.is_instance_of::<pyo3::types::PyBool>() { return Err(PyTypeError::new_err("highlight_color requires WD_COLOR_INDEX or None; use font.highlight for names")); }
            let index = value.extract::<usize>()?;
            HIGHLIGHT_COLORS.get(index).copied().ok_or_else(|| PyValueError::new_err("use WD_COLOR_INDEX or font.highlight with a supported color name"))
        }).transpose()?;
        self.apply(py, FontUpdate::Highlight(color))
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
    fn shading(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .shading
            .as_deref()
            .map(|value| color_object(py, value))
            .transpose()
    }

    #[setter]
    fn set_shading(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let value = value
            .filter(|value| !value.is_none())
            .map(|value| {
                if value
                    .extract::<String>()
                    .is_ok_and(|text| text.eq_ignore_ascii_case("auto"))
                {
                    Ok("auto".to_owned())
                } else {
                    color_hex(value, "Font.shading")
                }
            })
            .transpose()?;
        self.apply(py, FontUpdate::Shading(value.as_deref()))
    }
}

/// A live view of a run's literal and theme colour source.
#[pyclass(name = "ColorFormat")]
pub struct PyColorFormat {
    document: Py<PyDocument>,
    path: ContentPath,
    run_path: rdocx::AcceptedRunPath,
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
}

impl PyColorFormat {
    fn font(&self, py: Python<'_>) -> PyFont {
        PyFont::new(
            self.document.clone_ref(py),
            self.path.clone(),
            self.run_path.clone(),
            self.related.clone(),
            self.owner.clone(),
        )
    }
}

#[pymethods]
impl PyColorFormat {
    #[getter]
    fn rgb(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let snapshot = self.font(py).snapshot(py)?;
        match snapshot.color.as_deref() {
            None => Ok(None),
            Some(value) if value.eq_ignore_ascii_case("auto") => Ok(None),
            Some(value) => color_object(py, value).map(Some),
        }
    }

    #[setter]
    fn set_rgb(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let value = value
            .filter(|value| !value.is_none())
            .map(|value| color_hex(value, "Font.color.rgb"))
            .transpose()?;
        self.font(py).apply(py, FontUpdate::Color(value))
    }

    #[getter(r#type)]
    fn color_type(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let snapshot = self.font(py).snapshot(py)?;
        let kind = if snapshot.color_theme.is_some() {
            Some(2)
        } else {
            snapshot.color.as_deref().map(|value| {
                if value.eq_ignore_ascii_case("auto") {
                    101
                } else {
                    1
                }
            })
        };
        kind.map(|kind| {
            py.import("rdocx.enum.dml")?
                .getattr("MSO_COLOR_TYPE")?
                .call1((kind,))
                .map(Bound::unbind)
        })
        .transpose()
    }

    #[getter]
    fn theme_color(&self, py: Python<'_>) -> PyResult<Option<String>> {
        Ok(self.font(py).snapshot(py)?.color_theme)
    }

    #[setter]
    fn set_theme_color(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let value = value.filter(|value| !value.is_none()).map(|value| {
            let token = if let Ok(token) = value.extract::<String>() { token }
                else {
                    let code = if value.is_instance_of::<pyo3::types::PyBool>() {
                        None
                    } else { value.extract::<i32>().ok() };
                    let code = code.ok_or_else(|| PyTypeError::new_err(format!(
                        "Font.color.theme_color accepts a supported theme string or MSO_THEME_COLOR, got {}",
                        value.get_type().name().map(|name| name.to_string()).unwrap_or_else(|_| "unknown type".to_owned())
                    )))?;
                    THEME_COLORS.get(usize::try_from(code).ok().and_then(|code| code.checked_sub(1)).unwrap_or(usize::MAX))
                        .ok_or_else(|| PyValueError::new_err("Font.color.theme_color is not a supported MSO_THEME_COLOR"))?.to_string()
                };
            if !THEME_COLORS.contains(&token.as_str()) {
                return Err(PyValueError::new_err("Font.color.theme_color is not a supported theme token"));
            }
            Ok(token)
        }).transpose()?;
        let font = self.font(py);
        let current = font.snapshot(py)?;
        font.apply(
            py,
            FontUpdate::ColorTheme(value, current.color_theme_tint, current.color_theme_shade),
        )
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
    borders: Vec<(
        rdocx::ParagraphBorderEdge,
        String,
        Option<u32>,
        Option<String>,
    )>,
    shading: Option<String>,
    outline_level: Option<u32>,
    rtl: Option<bool>,
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
            borders: [
                rdocx::ParagraphBorderEdge::Top,
                rdocx::ParagraphBorderEdge::Bottom,
                rdocx::ParagraphBorderEdge::Left,
                rdocx::ParagraphBorderEdge::Right,
                rdocx::ParagraphBorderEdge::Between,
                rdocx::ParagraphBorderEdge::Bar,
            ]
            .into_iter()
            .filter_map(|edge| {
                paragraph.border(edge).map(|border| {
                    (
                        edge,
                        border.style().to_owned(),
                        border.size_eighths_pt(),
                        border.color().map(str::to_owned),
                    )
                })
            })
            .collect(),
            shading: paragraph.shading_fill().map(str::to_owned),
            outline_level: paragraph.outline_level(),
            rtl: paragraph.right_to_left_value(),
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
    AddTab(rdocx::TabAlignment, rdocx::Length, Option<rdocx::TabLeader>),
    RemoveTab(usize),
    ClearTabs,
    Border(
        rdocx::ParagraphBorderEdge,
        Option<(rdocx::BorderStyle, u32, String)>,
    ),
    ClearBorders,
    Shading(Option<String>),
    OutlineLevel(Option<u32>),
    Rtl(Option<bool>),
}

impl ParagraphUpdate {
    fn apply(self, paragraph: &mut rdocx::Paragraph<'_>) -> rdocx::Result<()> {
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
            Self::AddTab(alignment, position, leader) => {
                paragraph.try_add_tab_stop(alignment, position, leader)?;
            }
            Self::RemoveTab(index) => {
                paragraph.try_remove_tab_stop(index)?;
            }
            Self::ClearTabs => paragraph.try_clear_tab_stops()?,
            Self::Border(edge, value) => paragraph.set_border_value(
                edge,
                value
                    .as_ref()
                    .map(|(style, size, color)| (*style, *size, color.as_str())),
            ),
            Self::ClearBorders => paragraph.clear_borders(),
            Self::Shading(value) => paragraph.set_shading_fill_value(value.as_deref()),
            Self::OutlineLevel(value) => {
                assert!(paragraph.set_outline_level_value(value));
            }
            Self::Rtl(value) => paragraph.set_right_to_left_value(value),
        }
        Ok(())
    }
}

/// An owned read-only snapshot. It has no mutable child-position identity.
#[pyclass(name = "TabStop", frozen, from_py_object)]
#[derive(Clone)]
pub struct PyTabStop {
    position: rdocx::Length,
    alignment: rdocx::TabAlignment,
    leader: Option<rdocx::TabLeader>,
}

#[pymethods]
impl PyTabStop {
    #[getter]
    fn position(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        length_object(py, self.position)
    }
    #[getter]
    fn alignment(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        enum_object(py, "WD_TAB_ALIGNMENT", tab_alignment_to_int(self.alignment))
    }
    #[setter]
    fn set_alignment(&self, _value: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(pyo3::exceptions::PyNotImplementedError::new_err(
            "TabStop is a snapshot. First delete tabs[index], then readd with tabs.add_tab_stop(tab.position, alignment, tab.leader), preserving position and leader",
        ))
    }
    #[getter]
    fn leader(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        enum_object(py, "WD_TAB_LEADER", tab_leader_to_int(self.leader))
    }
}

#[pyclass(name = "TabStops")]
pub struct PyTabStops {
    document: Py<PyDocument>,
    path: ContentPath,
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
}

impl PyTabStops {
    fn format(&self, py: Python<'_>) -> PyParagraphFormat {
        PyParagraphFormat::new(
            self.document.clone_ref(py),
            self.path.clone(),
            self.related.clone(),
            self.owner.clone(),
        )
    }

    fn snapshots(&self, py: Python<'_>) -> PyResult<Vec<PyTabStop>> {
        let format = self.format(py);
        let location = format.validate(py)?;
        let document = self.document.borrow(py);
        let read = |paragraph: rdocx::ParagraphRef<'_>| -> PyResult<Vec<PyTabStop>> {
            (0..paragraph.tab_stop_count())
                .map(|index| {
                    let tab = paragraph
                        .tab_stop(index)
                        .expect("index is within native tab count");
                    let (alignment, leader) = tab
                        .checked_values()
                        .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
                    Ok(PyTabStop {
                        position: tab.position(),
                        alignment,
                        leader,
                    })
                })
                .collect()
        };
        location
            .read(&document.inner, read)
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?
    }
}

#[pymethods]
impl PyTabStops {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.snapshots(py)?.len())
    }
    fn __getitem__(&self, py: Python<'_>, index: isize) -> PyResult<PyTabStop> {
        let tabs = self.snapshots(py)?;
        let index = crate::normalize_index(index, tabs.len(), "tab stop")?;
        Ok(tabs[index].clone())
    }
    fn __delitem__(&self, py: Python<'_>, index: isize) -> PyResult<()> {
        let tabs = self.snapshots(py)?;
        let index = crate::normalize_index(index, tabs.len(), "tab stop")?;
        self.format(py).apply(py, ParagraphUpdate::RemoveTab(index))
    }
    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let tabs = self
            .snapshots(py)?
            .into_iter()
            .map(|tab| Py::new(py, tab))
            .collect::<PyResult<Vec<_>>>()?;
        pyo3::types::PyList::new(py, tabs)?
            .call_method0("__iter__")
            .map(Bound::unbind)
    }
    #[pyo3(signature = (position, alignment=0, leader=None))]
    fn add_tab_stop(
        &self,
        py: Python<'_>,
        position: i64,
        alignment: i32,
        leader: Option<i32>,
    ) -> PyResult<PyTabStop> {
        let position = i32::try_from(position / 635)
            .map(rdocx::Length::twips)
            .map_err(|_| {
                PyValueError::new_err("TabStops.add_tab_stop(position) exceeds signed twip range")
            })?;
        let alignment = tab_alignment_from_int(alignment)?;
        let leader = tab_leader_from_int(leader)?;
        let format = self.format(py);
        let location = format.validate(py)?;
        let duplicate = location
            .read(&self.document.borrow(py).inner, |paragraph| {
                (0..paragraph.tab_stop_count()).any(|index| {
                    paragraph
                        .tab_stop(index)
                        .is_some_and(|tab| tab.position() == position)
                })
            })
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        if duplicate {
            return Err(PyValueError::new_err(
                "a tab stop already exists at this stored position",
            ));
        }
        format.apply(py, ParagraphUpdate::AddTab(alignment, position, leader))?;
        Ok(PyTabStop {
            position,
            alignment,
            leader,
        })
    }
    fn clear_all(&self, py: Python<'_>) -> PyResult<()> {
        self.format(py).apply(py, ParagraphUpdate::ClearTabs)
    }
}

#[pyclass(name = "ParagraphFormat")]
pub struct PyParagraphFormat {
    document: Py<PyDocument>,
    path: ContentPath,
    related: Option<rdocx::ContentLocation>,
    owner: Option<oxml_py_support::RevisionState>,
}

impl PyParagraphFormat {
    pub(crate) fn new(
        document: Py<PyDocument>,
        path: ContentPath,
        related: Option<rdocx::ContentLocation>,
        owner: Option<oxml_py_support::RevisionState>,
    ) -> Self {
        Self {
            document,
            path,
            related,
            owner,
        }
    }

    fn validate(&self, py: Python<'_>) -> PyResult<ParagraphLocation> {
        let document = self.document.borrow(py);
        self.path
            .validate_scopes(
                &document.revisions,
                "paragraph format",
                "Re-fetch it with paragraph.paragraph_format.",
            )
            .map_err(|error| stale_to_pyerr(py, error))?;
        match &self.related {
            Some(location) => match &self.owner {
                Some(owner) => document
                    .current_paragraph_owner(owner)
                    .map(ParagraphLocation::Related),
                None => Ok(ParagraphLocation::Related(location.clone())),
            },
            None => paragraph_location(&self.path),
        }
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
    location
        .read(&document.inner, ParagraphSnapshot::from_paragraph)
        .map_err(|error| crate::rdocx_to_pyerr(py, error))
}

/// Apply one paragraph property update at a location the caller has validated.
pub(crate) fn apply_paragraph_update(
    py: Python<'_>,
    document: &Py<PyDocument>,
    location: ParagraphLocation,
    update: ParagraphUpdate,
) -> PyResult<()> {
    let mut document = document.borrow_mut(py);
    // A text-box paragraph is rewritten through its physical part. That
    // publication retires the handle, including the one that made the edit.
    // Other related stories keep the direct property edit.
    if let ParagraphLocation::Related(related) = &location
        && related.story().kind() == rdocx::StoryKind::TextBox
    {
        let mut candidate = document.inner.clone_for_staging();
        ParagraphLocation::Related(related.clone())
            .edit(&mut candidate, |paragraph| update.apply(paragraph))
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        let native = document
            .paragraph_route(&location, None)
            .map_err(|error| crate::rdocx_to_pyerr(py, error))?;
        return document.publish_owner_candidate(
            py,
            candidate,
            "ParagraphFormat",
            crate::document::WordPublication::ParagraphText(native),
        );
    }
    location
        .edit(&mut document.inner, |paragraph| update.apply(paragraph))
        .map_err(|error| crate::rdocx_to_pyerr(py, error))
}

#[pymethods]
impl PyParagraphFormat {
    #[pyo3(signature = (edge, style, *, size, color))]
    fn set_border(
        &self,
        py: Python<'_>,
        edge: &str,
        style: &str,
        size: u32,
        color: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let edge = paragraph_border_edge(edge)?;
        let style = crate::table::border_style_from_name(style)?;
        let color = paragraph_color_hex(color, "ParagraphFormat.set_border(color)")?;
        self.apply(
            py,
            ParagraphUpdate::Border(edge, Some((style, size, color))),
        )
    }

    fn border(&self, py: Python<'_>, edge: &str) -> PyResult<Option<BorderSnapshot>> {
        let edge = paragraph_border_edge(edge)?;
        self.snapshot(py)?
            .borders
            .into_iter()
            .find(|(candidate, _, _, _)| *candidate == edge)
            .map(|(_, style, size, color)| {
                Ok((
                    style,
                    size,
                    color
                        .as_deref()
                        .map(|value| color_object(py, value))
                        .transpose()?,
                ))
            })
            .transpose()
    }

    fn clear_border(&self, py: Python<'_>, edge: &str) -> PyResult<()> {
        self.apply(
            py,
            ParagraphUpdate::Border(paragraph_border_edge(edge)?, None),
        )
    }
    fn clear_borders(&self, py: Python<'_>) -> PyResult<()> {
        self.apply(py, ParagraphUpdate::ClearBorders)
    }
    #[getter]
    fn shading(&self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        self.snapshot(py)?
            .shading
            .as_deref()
            .map(|value| color_object(py, value))
            .transpose()
    }
    #[setter]
    fn set_shading(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let value = value
            .filter(|value| !value.is_none())
            .map(|value| paragraph_color_hex(value, "ParagraphFormat.shading"))
            .transpose()?;
        self.apply(py, ParagraphUpdate::Shading(value))
    }
    #[getter]
    fn outline_level(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        Ok(self.snapshot(py)?.outline_level)
    }
    #[setter]
    fn set_outline_level(&self, py: Python<'_>, value: Option<i64>) -> PyResult<()> {
        let value = value
            .map(|value| {
                u32::try_from(value)
                    .ok()
                    .filter(|value| *value <= 9)
                    .ok_or_else(|| {
                        PyValueError::new_err("ParagraphFormat.outline_level must be 0..9 or None")
                    })
            })
            .transpose()?;
        self.apply(py, ParagraphUpdate::OutlineLevel(value))
    }
    #[getter]
    fn rtl(&self, py: Python<'_>) -> PyResult<Option<bool>> {
        Ok(self.snapshot(py)?.rtl)
    }
    #[setter]
    fn set_rtl(&self, py: Python<'_>, value: Option<bool>) -> PyResult<()> {
        self.apply(py, ParagraphUpdate::Rtl(value))
    }

    #[getter]
    fn tab_stops(&self, py: Python<'_>) -> PyResult<PyTabStops> {
        self.validate(py)?;
        Ok(PyTabStops {
            document: self.document.clone_ref(py),
            path: self.path.clone(),
            related: self.related.clone(),
            owner: self.owner.clone(),
        })
    }

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
    fn set_space_after(&self, py: Python<'_>, value: Option<&Bound<'_, PyAny>>) -> PyResult<()> {
        let value = checked_spacing_length(py, value, "ParagraphFormat.space_after", false)?;
        self.apply(py, ParagraphUpdate::SpaceAfter(value))
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
        let Some(value) = value.filter(|value| !value.is_none()) else {
            return self.apply(py, ParagraphUpdate::LineSpacing(None));
        };
        if value.is_instance_of::<pyo3::types::PyBool>() {
            return Err(PyTypeError::new_err(
                "line_spacing rejects bool; use a numeric line multiple or Pt/Length",
            ));
        }
        if value.is_instance(&py.import("rdocx.shared")?.getattr("Length")?)? {
            let emu = value.extract::<i64>()?;
            i32::try_from(emu / 635).map_err(|_| {
                PyValueError::new_err(
                    "absolute line_spacing must fit signed twips; use a representable Pt/Length",
                )
            })?;
            return self.apply(
                py,
                ParagraphUpdate::LineSpacing(Some(rdocx::Length::emu(emu))),
            );
        }
        if !value.is_instance_of::<pyo3::types::PyInt>()
            && !value.is_instance_of::<pyo3::types::PyFloat>()
        {
            return Err(PyTypeError::new_err(
                "line_spacing requires a numeric line multiple or Pt/Length",
            ));
        }
        let multiple = require_auto_line_spacing(value)?;
        self.apply(py, ParagraphUpdate::LineSpacingMultiple(multiple))
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
