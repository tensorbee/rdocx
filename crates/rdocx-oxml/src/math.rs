//! Typed support for the Transitional OfficeMath vocabulary.

use std::io::Write;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::M_NS;
use crate::raw_xml::capture_element;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
struct Preservation {
    attributes: Vec<(String, String)>,
    raw_children: Vec<(usize, Vec<u8>)>,
    modeled_children: Vec<PreservedChild>,
    inherited_bindings: Vec<(String, String)>,
    property_raw_children: Vec<(usize, Vec<u8>)>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PreservedChild {
    name: String,
    raw: Vec<u8>,
    bindings: Vec<(String, String)>,
}

#[derive(Debug)]
struct ParsedElement {
    preservation: Preservation,
    bindings: Vec<(String, String)>,
    children: Vec<Vec<u8>>,
}

/// One inline OfficeMath equation.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CT_OMath {
    pub expressions: Vec<MathExpression>,
    preservation: Preservation,
}

impl CT_OMath {
    pub fn new(expressions: Vec<MathExpression>) -> Self {
        Self {
            expressions,
            preservation: Preservation::default(),
        }
    }

    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_raw(xml, &[])
    }

    pub(crate) fn from_raw(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        validate_xml_depth(xml)?;
        require_safe_fixed_math_prefix(xml, inherited)?;
        let mut parsed = parse_element(xml, inherited)?;
        require_root(xml, &parsed.bindings, b"oMath")?;
        let mut expressions = Vec::new();
        for child in parsed.children {
            if let Some(expression) = MathExpression::from_raw(&child, &parsed.bindings)? {
                expressions.push(expression);
            } else {
                parsed
                    .preservation
                    .raw_children
                    .push((expressions.len(), child));
            }
        }
        Ok(Self {
            expressions,
            preservation: parsed.preservation,
        })
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer, true)?;
        Ok(writer.into_inner())
    }

    /// Whether this equation retains content outside the typed OfficeMath subset.
    pub fn has_unsupported_content(&self) -> bool {
        preservation_has_unsupported_content(&self.preservation)
            || self
                .expressions
                .iter()
                .any(MathExpression::has_unsupported_content)
    }

    pub(crate) fn write_xml<W: Write>(
        &self,
        writer: &mut Writer<W>,
        declare_math: bool,
    ) -> Result<()> {
        let root = math_start("oMath", &self.preservation, declare_math);
        writer.write_event(Event::Start(root.borrow()))?;
        write_raw_slot(writer, &self.preservation, 0)?;
        for (index, expression) in self.expressions.iter().enumerate() {
            expression.write_xml(writer)?;
            write_raw_slot(writer, &self.preservation, index + 1)?;
        }
        write_raw_tail(writer, &self.preservation, self.expressions.len() + 1)?;
        writer.write_event(Event::End(BytesEnd::new("m:oMath")))?;
        Ok(())
    }
}

/// One display OfficeMath paragraph.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CT_OMathPara {
    pub properties: MathParagraphProperties,
    pub equations: Vec<CT_OMath>,
    preservation: Preservation,
}

impl CT_OMathPara {
    pub fn new(equations: Vec<CT_OMath>) -> Self {
        Self {
            equations,
            ..Self::default()
        }
    }

    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_raw(xml, &[])
    }

    pub(crate) fn from_raw(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        validate_xml_depth(xml)?;
        require_safe_fixed_math_prefix(xml, inherited)?;
        if !valid_officemath_paragraph_shape(xml, inherited)? {
            return Err(OxmlError::InvalidValue(
                "invalid OfficeMath paragraph child sequence".to_owned(),
            ));
        }
        let mut parsed = parse_element(xml, inherited)?;
        require_root(xml, &parsed.bindings, b"oMathPara")?;
        let mut properties = MathParagraphProperties::default();
        let mut equations = Vec::new();
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("oMathParaPr") if modeled == 0 => {
                    properties = MathParagraphProperties::from_raw(&child, &parsed.bindings)?;
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "oMathParaPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("oMath") => {
                    equations.push(CT_OMath::from_raw(&child, &parsed.bindings)?);
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            properties,
            equations,
            preservation: parsed.preservation,
        })
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer, true)?;
        Ok(writer.into_inner())
    }

    /// Whether this display equation retains content outside the typed subset.
    pub fn has_unsupported_content(&self) -> bool {
        preservation_has_unsupported_content(&self.preservation)
            || property_preservation_has_unsupported_content(
                &self.properties.preservation,
                "oMathParaPr",
            )
            || self.equations.iter().any(CT_OMath::has_unsupported_content)
    }

    pub(crate) fn write_xml<W: Write>(
        &self,
        writer: &mut Writer<W>,
        declare_math: bool,
    ) -> Result<()> {
        if self.equations.is_empty() {
            return Err(OxmlError::InvalidValue(
                "OfficeMath display requires at least one equation".to_owned(),
            ));
        }
        let root = math_start("oMathPara", &self.preservation, declare_math);
        writer.write_event(Event::Start(root.borrow()))?;
        let had_properties = has_preserved_modeled_child(&self.preservation, "oMathParaPr");
        let has_properties = !self.properties.is_empty() || had_properties;
        if has_properties && !had_properties {
            self.properties.write_xml(writer)?;
        }
        let mut modeled = 0usize;
        write_raw_slot(writer, &self.preservation, modeled)?;
        if has_properties && had_properties {
            self.properties.write_xml(writer)?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)?;
        }
        for equation in &self.equations {
            equation.write_xml(writer, false)?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)?;
        }
        write_raw_tail(writer, &self.preservation, modeled + 1)?;
        writer.write_event(Event::End(BytesEnd::new("m:oMathPara")))?;
        Ok(())
    }
}

/// Inline or display OfficeMath at a paragraph boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OfficeMath {
    Inline(CT_OMath),
    Display(CT_OMathPara),
}

impl OfficeMath {
    pub fn inline(expressions: Vec<MathExpression>) -> Self {
        Self::Inline(CT_OMath::new(expressions))
    }

    pub fn display(equations: Vec<CT_OMath>) -> Self {
        Self::Display(CT_OMathPara::new(equations))
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        match self {
            Self::Inline(value) => value.to_xml(),
            Self::Display(value) => value.to_xml(),
        }
    }

    /// Whether this value retains content outside the typed OfficeMath subset.
    pub fn has_unsupported_content(&self) -> bool {
        match self {
            Self::Inline(value) => value.has_unsupported_content(),
            Self::Display(value) => value.has_unsupported_content(),
        }
    }

    pub(crate) fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Option<Self>> {
        if !fixed_math_prefix_is_safe(raw, inherited)? {
            return Ok(None);
        }
        Ok(match math_local_name(raw, inherited)?.as_deref() {
            Some("oMath") => Some(Self::Inline(CT_OMath::from_raw(raw, inherited)?)),
            Some("oMathPara") if valid_officemath_paragraph_shape(raw, inherited)? => {
                Some(Self::Display(CT_OMathPara::from_raw(raw, inherited)?))
            }
            Some("oMathPara") => None,
            _ => None,
        })
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Inline(value) => value.write_xml(writer, true),
            Self::Display(value) => value.write_xml(writer, true),
        }
    }
}

/// One expression in the normalized OfficeMath tree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MathExpression {
    Run(MathRun),
    Fraction(MathFraction),
    Subscript(MathScript),
    Superscript(MathScript),
    SubSuperscript(MathSubSuperscript),
    PreSubSuperscript(MathPreSubSuperscript),
    Radical(MathRadical),
    Matrix(MathMatrix),
    LowerLimit(MathLimit),
    UpperLimit(MathLimit),
    Nary(MathNary),
    Delimiter(MathDelimiter),
    Accent(MathAccent),
}

impl MathExpression {
    /// Whether this expression or one of its arguments retains unsupported content.
    pub fn has_unsupported_content(&self) -> bool {
        match self {
            Self::Run(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_preservation_has_unsupported_content(
                        &value.properties.preservation,
                        "rPr",
                    )
                    || math_text_has_unsupported_content(&value.preservation)
            }
            Self::Fraction(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_container_has_unsupported_content(&value.preservation, "fPr")
                    || value.numerator.has_unsupported_content()
                    || value.denominator.has_unsupported_content()
            }
            Self::Subscript(value) => script_has_unsupported_content(value, "sSubPr"),
            Self::Superscript(value) => script_has_unsupported_content(value, "sSupPr"),
            Self::SubSuperscript(value) => three_argument_script_has_unsupported_content(
                &value.preservation,
                "sSubSupPr",
                &value.base,
                &value.subscript,
                &value.superscript,
            ),
            Self::PreSubSuperscript(value) => three_argument_script_has_unsupported_content(
                &value.preservation,
                "sPrePr",
                &value.base,
                &value.subscript,
                &value.superscript,
            ),
            Self::Radical(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_container_has_unsupported_content(&value.preservation, "radPr")
                    || value.degree.has_unsupported_content()
                    || value.base.has_unsupported_content()
            }
            Self::Matrix(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_container_has_unsupported_content(&value.preservation, "mPr")
                    || value.rows.iter().any(|row| {
                        preservation_has_unsupported_content(&row.preservation)
                            || row.cells.iter().any(MathArgument::has_unsupported_content)
                    })
            }
            Self::LowerLimit(value) => limit_has_unsupported_content(value, "limLowPr"),
            Self::UpperLimit(value) => limit_has_unsupported_content(value, "limUppPr"),
            Self::Nary(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_container_has_unsupported_content(&value.preservation, "naryPr")
                    || value.base.has_unsupported_content()
                    || value.subscript.has_unsupported_content()
                    || value.superscript.has_unsupported_content()
            }
            Self::Delimiter(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_container_has_unsupported_content(&value.preservation, "dPr")
                    || value
                        .arguments
                        .iter()
                        .any(MathArgument::has_unsupported_content)
            }
            Self::Accent(value) => {
                preservation_has_unsupported_content(&value.preservation)
                    || property_container_has_unsupported_content(&value.preservation, "accPr")
                    || value.base.has_unsupported_content()
            }
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Option<Self>> {
        if !valid_expression_shape(raw, inherited)? {
            return Ok(None);
        }
        Ok(match math_local_name(raw, inherited)?.as_deref() {
            Some("r") => Some(Self::Run(MathRun::from_raw(raw, inherited)?)),
            Some("f") => Some(Self::Fraction(MathFraction::from_raw(raw, inherited)?)),
            Some("sSub") => Some(Self::Subscript(MathScript::from_raw(raw, inherited)?)),
            Some("sSup") => Some(Self::Superscript(MathScript::from_raw(raw, inherited)?)),
            Some("sSubSup") => Some(Self::SubSuperscript(MathSubSuperscript::from_raw(
                raw, inherited,
            )?)),
            Some("sPre") => Some(Self::PreSubSuperscript(MathPreSubSuperscript::from_raw(
                raw, inherited,
            )?)),
            Some("rad") => Some(Self::Radical(MathRadical::from_raw(raw, inherited)?)),
            Some("m") => Some(Self::Matrix(MathMatrix::from_raw(raw, inherited)?)),
            Some("limLow") => Some(Self::LowerLimit(MathLimit::from_raw(raw, inherited)?)),
            Some("limUpp") => Some(Self::UpperLimit(MathLimit::from_raw(raw, inherited)?)),
            Some("nary") => Some(Self::Nary(MathNary::from_raw(raw, inherited)?)),
            Some("d") => Some(Self::Delimiter(MathDelimiter::from_raw(raw, inherited)?)),
            Some("acc") => Some(Self::Accent(MathAccent::from_raw(raw, inherited)?)),
            _ => None,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Run(value) => value.write_xml(writer),
            Self::Fraction(value) => value.write_xml(writer),
            Self::Subscript(value) => value.write_xml(writer, "sSub"),
            Self::Superscript(value) => value.write_xml(writer, "sSup"),
            Self::SubSuperscript(value) => value.write_xml(writer),
            Self::PreSubSuperscript(value) => value.write_xml(writer),
            Self::Radical(value) => value.write_xml(writer),
            Self::Matrix(value) => value.write_xml(writer),
            Self::LowerLimit(value) => value.write_xml(writer, "limLow"),
            Self::UpperLimit(value) => value.write_xml(writer, "limUpp"),
            Self::Nary(value) => value.write_xml(writer),
            Self::Delimiter(value) => value.write_xml(writer),
            Self::Accent(value) => value.write_xml(writer),
        }
    }
}

impl From<MathRun> for MathExpression {
    fn from(value: MathRun) -> Self {
        Self::Run(value)
    }
}

/// A recursively nested OfficeMath argument.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathArgument {
    pub expressions: Vec<MathExpression>,
    preservation: Preservation,
}

impl MathArgument {
    pub fn new(expressions: Vec<MathExpression>) -> Self {
        Self {
            expressions,
            preservation: Preservation::default(),
        }
    }

    pub fn text(text: impl Into<String>) -> Self {
        Self::new(vec![MathRun::new(text).into()])
    }

    /// Whether this argument retains content outside the typed OfficeMath subset.
    pub fn has_unsupported_content(&self) -> bool {
        preservation_has_unsupported_content(&self.preservation)
            || self
                .expressions
                .iter()
                .any(MathExpression::has_unsupported_content)
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut expressions = Vec::new();
        for child in parsed.children {
            if let Some(expression) = MathExpression::from_raw(&child, &parsed.bindings)? {
                expressions.push(expression);
            } else {
                parsed
                    .preservation
                    .raw_children
                    .push((expressions.len(), child));
            }
        }
        Ok(Self {
            expressions,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        write_container(writer, tag, &self.preservation, |writer| {
            write_raw_slot(writer, &self.preservation, 0)?;
            for (index, expression) in self.expressions.iter().enumerate() {
                expression.write_xml(writer)?;
                write_raw_slot(writer, &self.preservation, index + 1)?;
            }
            write_raw_tail(writer, &self.preservation, self.expressions.len() + 1)
        })
    }
}

/// OfficeMath run formatting needed by authoring and layout.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathRunProperties {
    pub style: Option<MathStyle>,
    pub normal: Option<bool>,
    pub literal: Option<bool>,
    pub script: Option<MathScriptStyle>,
    pub break_before: bool,
    pub break_alignment: Option<u8>,
    preservation: Preservation,
}

/// Mathematical glyph style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathStyle {
    Plain,
    Bold,
    Italic,
    BoldItalic,
}

impl MathStyle {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "p" => Some(Self::Plain),
            "b" => Some(Self::Bold),
            "i" => Some(Self::Italic),
            "bi" => Some(Self::BoldItalic),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Plain => "p",
            Self::Bold => "b",
            Self::Italic => "i",
            Self::BoldItalic => "bi",
        }
    }
}

/// Script category for a math run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathScriptStyle {
    Roman,
    Script,
    Fraktur,
    DoubleStruck,
    SansSerif,
    Monospace,
}

impl MathScriptStyle {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "roman" => Some(Self::Roman),
            "script" => Some(Self::Script),
            "fraktur" => Some(Self::Fraktur),
            "double-struck" => Some(Self::DoubleStruck),
            "sans-serif" => Some(Self::SansSerif),
            "monospace" => Some(Self::Monospace),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Roman => "roman",
            Self::Script => "script",
            Self::Fraktur => "fraktur",
            Self::DoubleStruck => "double-struck",
            Self::SansSerif => "sans-serif",
            Self::Monospace => "monospace",
        }
    }
}

/// Text-bearing OfficeMath run.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathRun {
    pub text: String,
    pub properties: MathRunProperties,
    preservation: Preservation,
}

impl MathRun {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut properties = MathRunProperties::default();
        let mut text = String::new();
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("rPr") if modeled == 0 => {
                    properties = parse_run_properties(&child, &parsed.bindings)?;
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "rPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("t") => {
                    text.push_str(&element_text(&child)?);
                    preserve_modeled_child(&mut parsed.preservation, "t", &child, &parsed.bindings);
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            text,
            properties,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_container(writer, "r", &self.preservation, |writer| {
            let had_properties = has_preserved_modeled_child(&self.preservation, "rPr");
            let emit_properties = !self.properties.is_empty() || had_properties;
            if emit_properties && !had_properties {
                write_run_properties(writer, &self.properties)?;
            }
            let mut modeled = 0usize;
            write_raw_slot(writer, &self.preservation, modeled)?;
            if emit_properties && had_properties {
                write_run_properties(writer, &self.properties)?;
                modeled += 1;
                write_raw_slot(writer, &self.preservation, modeled)?;
            }
            write_math_text(writer, &self.text, &self.preservation)?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

impl MathRunProperties {
    fn is_empty(&self) -> bool {
        self.style.is_none()
            && self.normal.is_none()
            && self.literal.is_none()
            && self.script.is_none()
            && !self.break_before
            && self.break_alignment.is_none()
            && self.preservation == Preservation::default()
    }
}

/// Fraction bar form.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FractionType {
    #[default]
    Bar,
    Skewed,
    Linear,
    NoBar,
}

impl FractionType {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "bar" => Some(Self::Bar),
            "skw" => Some(Self::Skewed),
            "lin" => Some(Self::Linear),
            "noBar" => Some(Self::NoBar),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Bar => "bar",
            Self::Skewed => "skw",
            Self::Linear => "lin",
            Self::NoBar => "noBar",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathFraction {
    pub fraction_type: FractionType,
    pub numerator: MathArgument,
    pub denominator: MathArgument,
    preservation: Preservation,
}

impl MathFraction {
    pub fn new(numerator: MathArgument, denominator: MathArgument) -> Self {
        Self {
            fraction_type: FractionType::Bar,
            numerator,
            denominator,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut fraction_type = FractionType::Bar;
        let mut numerator = MathArgument::default();
        let mut denominator = MathArgument::default();
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("fPr") if modeled == 0 => {
                    if let Some(value) = child_property_value(&child, &parsed.bindings, "type")? {
                        fraction_type = FractionType::parse(&value).unwrap_or_default();
                    }
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "fPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("num") => {
                    numerator = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                Some("den") => {
                    denominator = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            fraction_type,
            numerator,
            denominator,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_container(writer, "f", &self.preservation, |writer| {
            let mut modeled = write_leading_property_container(
                writer,
                "fPr",
                &[Property::text("type", self.fraction_type.as_str())],
                &self.preservation,
                true,
            )?;
            self.numerator.write_xml(writer, "num")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)?;
            self.denominator.write_xml(writer, "den")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

/// Base and one script argument, used for subscript and superscript forms.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathScript {
    pub base: MathArgument,
    pub script: MathArgument,
    pub alignment: Option<bool>,
    preservation: Preservation,
}

impl MathScript {
    pub fn new(base: MathArgument, script: MathArgument) -> Self {
        Self {
            base,
            script,
            alignment: None,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let root = math_local_name(raw, inherited)?.unwrap_or_default();
        let script_tag = if root == "sSub" { "sub" } else { "sup" };
        let property_tag = if root == "sSub" { "sSubPr" } else { "sSupPr" };
        let mut base = MathArgument::default();
        let mut script = MathArgument::default();
        let mut alignment = None;
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some(local) if local == property_tag && modeled == 0 => {
                    alignment = child_on_off_value(&child, &parsed.bindings, "alnScr")?;
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        property_tag,
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("e") => {
                    base = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                Some(local) if local == script_tag => {
                    script = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            base,
            script,
            alignment,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let property_tag = if tag == "sSub" { "sSubPr" } else { "sSupPr" };
        let script_tag = if tag == "sSub" { "sub" } else { "sup" };
        write_container(writer, tag, &self.preservation, |writer| {
            let properties = self
                .alignment
                .map(|value| Property::boolean("alnScr", value))
                .into_iter()
                .collect::<Vec<_>>();
            let emit_properties = self.alignment.is_some()
                || has_preserved_modeled_child(&self.preservation, property_tag);
            let mut modeled = write_leading_property_container(
                writer,
                property_tag,
                &properties,
                &self.preservation,
                emit_properties,
            )?;
            self.base.write_xml(writer, "e")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)?;
            self.script.write_xml(writer, script_tag)?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathSubSuperscript {
    pub base: MathArgument,
    pub subscript: MathArgument,
    pub superscript: MathArgument,
    pub alignment: Option<bool>,
    preservation: Preservation,
}

impl MathSubSuperscript {
    pub fn new(base: MathArgument, subscript: MathArgument, superscript: MathArgument) -> Self {
        Self {
            base,
            subscript,
            superscript,
            alignment: None,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        parse_three_argument_script(raw, inherited, false).map(|value| Self {
            base: value.0,
            subscript: value.1,
            superscript: value.2,
            alignment: value.3,
            preservation: value.4,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_three_argument_script(
            writer,
            "sSubSup",
            "sSubSupPr",
            &self.base,
            &self.subscript,
            &self.superscript,
            self.alignment,
            &self.preservation,
            false,
        )
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathPreSubSuperscript {
    pub base: MathArgument,
    pub subscript: MathArgument,
    pub superscript: MathArgument,
    pub alignment: Option<bool>,
    preservation: Preservation,
}

impl MathPreSubSuperscript {
    pub fn new(base: MathArgument, subscript: MathArgument, superscript: MathArgument) -> Self {
        Self {
            base,
            subscript,
            superscript,
            alignment: None,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        parse_three_argument_script(raw, inherited, true).map(|value| Self {
            base: value.0,
            subscript: value.1,
            superscript: value.2,
            alignment: value.3,
            preservation: value.4,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_three_argument_script(
            writer,
            "sPre",
            "sPrePr",
            &self.base,
            &self.subscript,
            &self.superscript,
            self.alignment,
            &self.preservation,
            true,
        )
    }
}

type ThreeArgumentScript = (
    MathArgument,
    MathArgument,
    MathArgument,
    Option<bool>,
    Preservation,
);

fn parse_three_argument_script(
    raw: &[u8],
    inherited: &[(String, String)],
    pre: bool,
) -> Result<ThreeArgumentScript> {
    let mut parsed = parse_element(raw, inherited)?;
    let property_tag = if pre { "sPrePr" } else { "sSubSupPr" };
    let mut base = MathArgument::default();
    let mut subscript = MathArgument::default();
    let mut superscript = MathArgument::default();
    let mut alignment = None;
    let mut modeled = 0usize;
    for child in parsed.children {
        match math_local_name(&child, &parsed.bindings)?.as_deref() {
            Some(local) if local == property_tag && modeled == 0 => {
                alignment = child_on_off_value(&child, &parsed.bindings, "alnScr")?;
                preserve_modeled_child(
                    &mut parsed.preservation,
                    property_tag,
                    &child,
                    &parsed.bindings,
                );
                modeled += 1;
            }
            Some("e") => {
                base = MathArgument::from_raw(&child, &parsed.bindings)?;
                modeled += 1;
            }
            Some("sub") => {
                subscript = MathArgument::from_raw(&child, &parsed.bindings)?;
                modeled += 1;
            }
            Some("sup") => {
                superscript = MathArgument::from_raw(&child, &parsed.bindings)?;
                modeled += 1;
            }
            _ => parsed.preservation.raw_children.push((modeled, child)),
        }
    }
    Ok((base, subscript, superscript, alignment, parsed.preservation))
}

#[allow(clippy::too_many_arguments)]
fn write_three_argument_script<W: Write>(
    writer: &mut Writer<W>,
    root_tag: &str,
    property_tag: &str,
    base: &MathArgument,
    subscript: &MathArgument,
    superscript: &MathArgument,
    alignment: Option<bool>,
    preservation: &Preservation,
    pre: bool,
) -> Result<()> {
    write_container(writer, root_tag, preservation, |writer| {
        let properties = alignment
            .map(|value| Property::boolean("alnScr", value))
            .into_iter()
            .collect::<Vec<_>>();
        let emit_properties =
            alignment.is_some() || has_preserved_modeled_child(preservation, property_tag);
        let mut modeled = write_leading_property_container(
            writer,
            property_tag,
            &properties,
            preservation,
            emit_properties,
        )?;
        if pre {
            subscript.write_xml(writer, "sub")?;
            modeled += 1;
            write_raw_slot(writer, preservation, modeled)?;
            superscript.write_xml(writer, "sup")?;
            modeled += 1;
            write_raw_slot(writer, preservation, modeled)?;
            base.write_xml(writer, "e")?;
        } else {
            base.write_xml(writer, "e")?;
            modeled += 1;
            write_raw_slot(writer, preservation, modeled)?;
            subscript.write_xml(writer, "sub")?;
            modeled += 1;
            write_raw_slot(writer, preservation, modeled)?;
            superscript.write_xml(writer, "sup")?;
        }
        modeled += 1;
        write_raw_slot(writer, preservation, modeled)
    })
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathRadical {
    pub degree: MathArgument,
    pub base: MathArgument,
    pub hide_degree: bool,
    degree_present: bool,
    preservation: Preservation,
}

impl MathRadical {
    pub fn new(base: MathArgument) -> Self {
        Self {
            degree: MathArgument::default(),
            base,
            hide_degree: true,
            degree_present: false,
            preservation: Preservation::default(),
        }
    }

    pub fn with_degree(degree: MathArgument, base: MathArgument) -> Self {
        Self {
            degree,
            base,
            hide_degree: false,
            degree_present: true,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut degree = MathArgument::default();
        let mut base = MathArgument::default();
        let mut hide_degree = false;
        let mut degree_present = false;
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("radPr") if modeled == 0 => {
                    hide_degree =
                        child_on_off_value(&child, &parsed.bindings, "degHide")?.unwrap_or(false);
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "radPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("deg") => {
                    degree = MathArgument::from_raw(&child, &parsed.bindings)?;
                    degree_present = true;
                    modeled += 1;
                }
                Some("e") => {
                    base = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            degree,
            base,
            hide_degree,
            degree_present,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_container(writer, "rad", &self.preservation, |writer| {
            let mut modeled = write_leading_property_container_before_raw(
                writer,
                "radPr",
                &[Property::boolean("degHide", self.hide_degree)],
                &self.preservation,
                true,
            )?;
            let emit_degree = self.degree_present
                || !self.degree.expressions.is_empty()
                || self.degree.has_unsupported_content();
            if self.degree_present {
                write_raw_slot(writer, &self.preservation, modeled)?;
                self.degree.write_xml(writer, "deg")?;
                modeled += 1;
            } else if emit_degree {
                self.degree.write_xml(writer, "deg")?;
            }
            write_raw_slot(writer, &self.preservation, modeled)?;
            self.base.write_xml(writer, "e")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathJustification {
    Left,
    Center,
    Right,
    CenterGroup,
}

impl MathJustification {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "left" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" => Some(Self::Right),
            "centerGroup" => Some(Self::CenterGroup),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::CenterGroup => "centerGroup",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathParagraphProperties {
    pub justification: Option<MathJustification>,
    preservation: Preservation,
}

impl MathParagraphProperties {
    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        Ok(Self {
            justification: child_property_value(raw, inherited, "jc")?
                .and_then(|value| MathJustification::parse(&value)),
            preservation: preserve_unrecognized_property_children(raw, inherited, &["jc"])?,
        })
    }

    fn is_empty(&self) -> bool {
        self.justification.is_none() && self.preservation == Preservation::default()
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let properties = self
            .justification
            .map(|value| Property::text("jc", value.as_str()))
            .into_iter()
            .collect::<Vec<_>>();
        write_property_container_with_preservation(
            writer,
            "oMathParaPr",
            &properties,
            &self.preservation,
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MatrixBaseJustification {
    Top,
    Center,
    Bottom,
}

impl MatrixBaseJustification {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "bottom" => Some(Self::Bottom),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathMatrixProperties {
    pub base_justification: Option<MatrixBaseJustification>,
    pub row_spacing: Option<u16>,
    pub column_spacing: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathMatrixRow {
    pub cells: Vec<MathArgument>,
    preservation: Preservation,
}

impl MathMatrixRow {
    pub fn new(cells: Vec<MathArgument>) -> Self {
        Self {
            cells,
            preservation: Preservation::default(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathMatrix {
    pub properties: MathMatrixProperties,
    pub rows: Vec<MathMatrixRow>,
    preservation: Preservation,
}

impl MathMatrix {
    pub fn new(rows: Vec<MathMatrixRow>) -> Self {
        Self {
            rows,
            ..Self::default()
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut properties = MathMatrixProperties::default();
        let mut rows = Vec::new();
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("mPr") if modeled == 0 => {
                    properties.base_justification =
                        child_property_value(&child, &parsed.bindings, "baseJc")?
                            .and_then(|value| MatrixBaseJustification::parse(&value));
                    properties.row_spacing = child_property_value(&child, &parsed.bindings, "rSp")?
                        .and_then(|value| value.parse().ok());
                    properties.column_spacing =
                        child_property_value(&child, &parsed.bindings, "cSp")?
                            .and_then(|value| value.parse().ok());
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "mPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("mr") => {
                    rows.push(parse_matrix_row(&child, &parsed.bindings)?);
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            properties,
            rows,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.rows.is_empty() || self.rows.iter().any(|row| row.cells.is_empty()) {
            return Err(OxmlError::InvalidValue(
                "OfficeMath matrix requires at least one nonempty row".to_owned(),
            ));
        }
        write_container(writer, "m", &self.preservation, |writer| {
            let properties = [
                self.properties
                    .base_justification
                    .map(|value| Property::text("baseJc", value.as_str())),
                self.properties
                    .row_spacing
                    .map(|value| Property::unsigned("rSp", value.into())),
                self.properties
                    .column_spacing
                    .map(|value| Property::unsigned("cSp", value.into())),
            ]
            .into_iter()
            .flatten()
            .collect::<Vec<_>>();
            let emit_properties =
                !properties.is_empty() || has_preserved_modeled_child(&self.preservation, "mPr");
            let mut modeled = write_leading_property_container(
                writer,
                "mPr",
                &properties,
                &self.preservation,
                emit_properties,
            )?;
            for row in &self.rows {
                write_container(writer, "mr", &row.preservation, |writer| {
                    write_raw_slot(writer, &row.preservation, 0)?;
                    for (index, cell) in row.cells.iter().enumerate() {
                        cell.write_xml(writer, "e")?;
                        write_raw_slot(writer, &row.preservation, index + 1)?;
                    }
                    write_raw_tail(writer, &row.preservation, row.cells.len() + 1)
                })?;
                modeled += 1;
                write_raw_slot(writer, &self.preservation, modeled)?;
            }
            write_raw_tail(writer, &self.preservation, modeled + 1)
        })
    }
}

fn parse_matrix_row(raw: &[u8], inherited: &[(String, String)]) -> Result<MathMatrixRow> {
    let mut parsed = parse_element(raw, inherited)?;
    let mut cells = Vec::new();
    for child in parsed.children {
        if math_local_name(&child, &parsed.bindings)?.as_deref() == Some("e") {
            cells.push(MathArgument::from_raw(&child, &parsed.bindings)?);
        } else {
            parsed.preservation.raw_children.push((cells.len(), child));
        }
    }
    Ok(MathMatrixRow {
        cells,
        preservation: parsed.preservation,
    })
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathLimit {
    pub base: MathArgument,
    pub limit: MathArgument,
    preservation: Preservation,
}

impl MathLimit {
    pub fn new(base: MathArgument, limit: MathArgument) -> Self {
        Self {
            base,
            limit,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut base = MathArgument::default();
        let mut limit = MathArgument::default();
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some(property @ ("limLowPr" | "limUppPr")) if modeled == 0 => {
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        property,
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("e") => {
                    base = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                Some("lim") => {
                    limit = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            base,
            limit,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        write_container(writer, tag, &self.preservation, |writer| {
            let property_tag = if tag == "limLow" {
                "limLowPr"
            } else {
                "limUppPr"
            };
            let mut modeled = 0usize;
            write_raw_slot(writer, &self.preservation, modeled)?;
            if has_preserved_modeled_child(&self.preservation, property_tag) {
                write_property_container_from_parent(
                    writer,
                    property_tag,
                    &[],
                    &self.preservation,
                )?;
                modeled += 1;
                write_raw_slot(writer, &self.preservation, modeled)?;
            }
            self.base.write_xml(writer, "e")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)?;
            self.limit.write_xml(writer, "lim")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LimitLocation {
    SubSuperscript,
    UnderOver,
}

impl LimitLocation {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "subSup" => Some(Self::SubSuperscript),
            "undOvr" => Some(Self::UnderOver),
            _ => None,
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::SubSuperscript => "subSup",
            Self::UnderOver => "undOvr",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathNary {
    pub character: String,
    pub base: MathArgument,
    pub subscript: MathArgument,
    pub superscript: MathArgument,
    pub hide_subscript: bool,
    pub hide_superscript: bool,
    pub grow: Option<bool>,
    pub limit_location: Option<LimitLocation>,
    subscript_present: bool,
    superscript_present: bool,
    preservation: Preservation,
}

impl MathNary {
    pub fn new(character: impl Into<String>, base: MathArgument) -> Self {
        Self {
            character: character.into(),
            base,
            subscript: MathArgument::default(),
            superscript: MathArgument::default(),
            hide_subscript: true,
            hide_superscript: true,
            grow: None,
            limit_location: None,
            subscript_present: false,
            superscript_present: false,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut value = Self::new("∫", MathArgument::default());
        value.hide_subscript = false;
        value.hide_superscript = false;
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("naryPr") if modeled == 0 => {
                    value.character = child_property_value(&child, &parsed.bindings, "chr")?
                        .unwrap_or(value.character);
                    value.hide_subscript =
                        child_on_off_value(&child, &parsed.bindings, "subHide")?.unwrap_or(false);
                    value.hide_superscript =
                        child_on_off_value(&child, &parsed.bindings, "supHide")?.unwrap_or(false);
                    value.grow = child_on_off_value(&child, &parsed.bindings, "grow")?;
                    value.limit_location =
                        child_property_value(&child, &parsed.bindings, "limLoc")?
                            .and_then(|v| LimitLocation::parse(&v));
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "naryPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("sub") => {
                    value.subscript = MathArgument::from_raw(&child, &parsed.bindings)?;
                    value.subscript_present = true;
                    modeled += 1;
                }
                Some("sup") => {
                    value.superscript = MathArgument::from_raw(&child, &parsed.bindings)?;
                    value.superscript_present = true;
                    modeled += 1;
                }
                Some("e") => {
                    value.base = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        value.preservation = parsed.preservation;
        Ok(value)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.character.chars().count() != 1 {
            return Err(OxmlError::InvalidValue(
                "OfficeMath n-ary character must contain one Unicode scalar".to_owned(),
            ));
        }
        write_container(writer, "nary", &self.preservation, |writer| {
            let mut properties = vec![Property::text("chr", &self.character)];
            if let Some(value) = self.limit_location {
                properties.push(Property::text("limLoc", value.as_str()));
            }
            if let Some(value) = self.grow {
                properties.push(Property::boolean("grow", value));
            }
            properties.push(Property::boolean("subHide", self.hide_subscript));
            properties.push(Property::boolean("supHide", self.hide_superscript));
            let mut modeled = write_leading_property_container_before_raw(
                writer,
                "naryPr",
                &properties,
                &self.preservation,
                true,
            )?;
            let emit_subscript = self.subscript_present
                || !self.subscript.expressions.is_empty()
                || self.subscript.has_unsupported_content();
            if self.subscript_present {
                write_raw_slot(writer, &self.preservation, modeled)?;
                self.subscript.write_xml(writer, "sub")?;
                modeled += 1;
            } else if emit_subscript {
                self.subscript.write_xml(writer, "sub")?;
            }
            let emit_superscript = self.superscript_present
                || !self.superscript.expressions.is_empty()
                || self.superscript.has_unsupported_content();
            if self.superscript_present {
                write_raw_slot(writer, &self.preservation, modeled)?;
                self.superscript.write_xml(writer, "sup")?;
                modeled += 1;
            } else if emit_superscript {
                self.superscript.write_xml(writer, "sup")?;
            }
            write_raw_slot(writer, &self.preservation, modeled)?;
            self.base.write_xml(writer, "e")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathDelimiter {
    pub begin_character: String,
    pub end_character: String,
    pub separator_character: String,
    pub grow: Option<bool>,
    pub arguments: Vec<MathArgument>,
    preservation: Preservation,
}

impl MathDelimiter {
    pub fn new(
        begin: impl Into<String>,
        end: impl Into<String>,
        arguments: Vec<MathArgument>,
    ) -> Self {
        Self {
            begin_character: begin.into(),
            end_character: end.into(),
            separator_character: "|".to_owned(),
            grow: None,
            arguments,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut value = Self::new("(", ")", Vec::new());
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("dPr") if modeled == 0 => {
                    value.begin_character =
                        child_property_value(&child, &parsed.bindings, "begChr")?
                            .unwrap_or(value.begin_character);
                    value.end_character = child_property_value(&child, &parsed.bindings, "endChr")?
                        .unwrap_or(value.end_character);
                    value.separator_character =
                        child_property_value(&child, &parsed.bindings, "sepChr")?
                            .unwrap_or(value.separator_character);
                    value.grow = child_on_off_value(&child, &parsed.bindings, "grow")?;
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "dPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("e") => {
                    value
                        .arguments
                        .push(MathArgument::from_raw(&child, &parsed.bindings)?);
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        value.preservation = parsed.preservation;
        Ok(value)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.arguments.is_empty()
            || [
                &self.begin_character,
                &self.separator_character,
                &self.end_character,
            ]
            .iter()
            .any(|value| value.chars().count() > 1)
        {
            return Err(OxmlError::InvalidValue(
                "OfficeMath delimiter requires arguments and scalar characters".to_owned(),
            ));
        }
        write_container(writer, "d", &self.preservation, |writer| {
            let mut properties = vec![
                Property::text("begChr", &self.begin_character),
                Property::text("sepChr", &self.separator_character),
                Property::text("endChr", &self.end_character),
            ];
            if let Some(value) = self.grow {
                properties.push(Property::boolean("grow", value));
            }
            let mut modeled = write_leading_property_container(
                writer,
                "dPr",
                &properties,
                &self.preservation,
                true,
            )?;
            for argument in &self.arguments {
                argument.write_xml(writer, "e")?;
                modeled += 1;
                write_raw_slot(writer, &self.preservation, modeled)?;
            }
            write_raw_tail(writer, &self.preservation, modeled + 1)
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathAccent {
    pub character: String,
    pub base: MathArgument,
    preservation: Preservation,
}

impl MathAccent {
    pub fn new(character: impl Into<String>, base: MathArgument) -> Self {
        Self {
            character: character.into(),
            base,
            preservation: Preservation::default(),
        }
    }

    fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut parsed = parse_element(raw, inherited)?;
        let mut character = "̂".to_owned();
        let mut base = MathArgument::default();
        let mut modeled = 0usize;
        for child in parsed.children {
            match math_local_name(&child, &parsed.bindings)?.as_deref() {
                Some("accPr") if modeled == 0 => {
                    character =
                        child_property_value(&child, &parsed.bindings, "chr")?.unwrap_or(character);
                    preserve_modeled_child(
                        &mut parsed.preservation,
                        "accPr",
                        &child,
                        &parsed.bindings,
                    );
                    modeled += 1;
                }
                Some("e") => {
                    base = MathArgument::from_raw(&child, &parsed.bindings)?;
                    modeled += 1;
                }
                _ => parsed.preservation.raw_children.push((modeled, child)),
            }
        }
        Ok(Self {
            character,
            base,
            preservation: parsed.preservation,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.character.chars().count() != 1 {
            return Err(OxmlError::InvalidValue(
                "OfficeMath accent character must contain one Unicode scalar".to_owned(),
            ));
        }
        write_container(writer, "acc", &self.preservation, |writer| {
            let mut modeled = write_leading_property_container(
                writer,
                "accPr",
                &[Property::text("chr", &self.character)],
                &self.preservation,
                true,
            )?;
            self.base.write_xml(writer, "e")?;
            modeled += 1;
            write_raw_slot(writer, &self.preservation, modeled)
        })
    }
}

/// Document-wide defaults from `w:settings/m:mathPr`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MathProperties {
    pub math_font: Option<String>,
    pub justification: Option<MathJustification>,
    pub left_margin: Option<u32>,
    pub right_margin: Option<u32>,
    pub pre_spacing: Option<u32>,
    pub post_spacing: Option<u32>,
    pub inter_spacing: Option<u32>,
    pub intra_spacing: Option<u32>,
    pub wrap_indent: Option<u32>,
    pub small_fraction: Option<bool>,
    pub display_defaults: Option<bool>,
    pub integral_limit_location: Option<LimitLocation>,
    pub nary_limit_location: Option<LimitLocation>,
    preservation: Preservation,
}

impl MathProperties {
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether the settings subtree retains content outside the typed subset.
    pub fn has_unsupported_content(&self) -> bool {
        property_preservation_has_unsupported_content(&self.preservation, "mathPr")
    }

    pub(crate) fn from_raw(raw: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        require_safe_fixed_math_prefix(raw, inherited)?;
        let mut parsed = parse_element(raw, inherited)?;
        require_root(raw, &parsed.bindings, b"mathPr")?;
        let mut value = Self::default();
        let mut modeled = 0usize;
        let mut seen = Vec::new();
        let mut last_rank = None;
        for child in parsed.children {
            let local = math_local_name(&child, &parsed.bindings)?;
            let Some(name) = local.as_deref() else {
                let key = property_raw_key("mathPr", None, &mut last_rank, false);
                parsed.preservation.property_raw_children.push((key, child));
                continue;
            };
            let duplicate = seen.iter().any(|seen_name| seen_name == name);
            if !supported_properties("mathPr").contains(&name)
                || duplicate
                || !property_leaf_is_valid("mathPr", name, &child, &parsed.bindings)?
            {
                let key = property_raw_key("mathPr", Some(name), &mut last_rank, duplicate);
                parsed.preservation.property_raw_children.push((key, child));
                continue;
            }
            property_raw_key("mathPr", Some(name), &mut last_rank, false);
            let parsed_value = root_value(&child, &parsed.bindings)?;
            let recognized = match name {
                "mathFont" => {
                    value.math_font = parsed_value;
                    true
                }
                "defJc" => {
                    value.justification =
                        parsed_value.as_deref().and_then(MathJustification::parse);
                    true
                }
                "lMargin" => {
                    value.left_margin = parse_u32(parsed_value);
                    true
                }
                "rMargin" => {
                    value.right_margin = parse_u32(parsed_value);
                    true
                }
                "preSp" => {
                    value.pre_spacing = parse_u32(parsed_value);
                    true
                }
                "postSp" => {
                    value.post_spacing = parse_u32(parsed_value);
                    true
                }
                "interSp" => {
                    value.inter_spacing = parse_u32(parsed_value);
                    true
                }
                "intraSp" => {
                    value.intra_spacing = parse_u32(parsed_value);
                    true
                }
                "wrapIndent" => {
                    value.wrap_indent = parse_u32(parsed_value);
                    true
                }
                "smallFrac" => {
                    value.small_fraction = parse_present_on_off(parsed_value.as_deref());
                    true
                }
                "dispDef" => {
                    value.display_defaults = parse_present_on_off(parsed_value.as_deref());
                    true
                }
                "intLim" => {
                    value.integral_limit_location =
                        parsed_value.as_deref().and_then(LimitLocation::parse);
                    true
                }
                "naryLim" => {
                    value.nary_limit_location =
                        parsed_value.as_deref().and_then(LimitLocation::parse);
                    true
                }
                _ => false,
            };
            if recognized {
                seen.push(name.to_owned());
                preserve_modeled_child(&mut parsed.preservation, name, &child, &parsed.bindings);
                modeled += 1;
            } else {
                parsed.preservation.raw_children.push((modeled, child));
            }
        }
        value.preservation = parsed.preservation;
        Ok(value)
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut properties = Vec::new();
        if let Some(value) = &self.math_font {
            properties.push(Property::text("mathFont", value));
        }
        if let Some(value) = self.small_fraction {
            properties.push(Property::boolean("smallFrac", value));
        }
        if let Some(value) = self.display_defaults {
            properties.push(Property::boolean("dispDef", value));
        }
        if let Some(value) = self.left_margin {
            properties.push(Property::unsigned("lMargin", value.into()));
        }
        if let Some(value) = self.right_margin {
            properties.push(Property::unsigned("rMargin", value.into()));
        }
        if let Some(value) = self.justification {
            properties.push(Property::text("defJc", value.as_str()));
        }
        for (name, value) in [
            ("preSp", self.pre_spacing),
            ("postSp", self.post_spacing),
            ("interSp", self.inter_spacing),
            ("intraSp", self.intra_spacing),
            ("wrapIndent", self.wrap_indent),
        ] {
            if let Some(value) = value {
                properties.push(Property::unsigned(name, value.into()));
            }
        }
        if let Some(value) = self.integral_limit_location {
            properties.push(Property::text("intLim", value.as_str()));
        }
        if let Some(value) = self.nary_limit_location {
            properties.push(Property::text("naryLim", value.as_str()));
        }
        write_property_container_with_preservation(
            writer,
            "mathPr",
            &properties,
            &self.preservation,
        )
    }
}

pub(crate) fn is_math_element(name: &[u8], local: &[u8], scope: &[String]) -> bool {
    let (prefix, actual_local) = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or((&b""[..], name), |at| (&name[..at], &name[at + 1..]));
    actual_local == local
        && scope.iter().any(|binding| {
            binding
                .strip_prefix('\0')
                .and_then(|binding| binding.split_once('\0'))
                .is_some_and(|(candidate, namespace)| {
                    candidate.as_bytes() == prefix && namespace == M_NS
                })
        })
}

fn parse_run_properties(raw: &[u8], inherited: &[(String, String)]) -> Result<MathRunProperties> {
    let mut parsed = parse_element(raw, inherited)?;
    let mut value = MathRunProperties::default();
    let mut modeled = 0usize;
    let mut seen = Vec::new();
    let mut last_rank = None;
    for child in parsed.children {
        let local = math_local_name(&child, &parsed.bindings)?;
        let Some(name) = local.as_deref() else {
            let key = property_raw_key("rPr", None, &mut last_rank, false);
            parsed.preservation.property_raw_children.push((key, child));
            continue;
        };
        let duplicate = seen.iter().any(|seen_name| seen_name == name);
        if !supported_properties("rPr").contains(&name)
            || duplicate
            || !property_leaf_is_valid("rPr", name, &child, &parsed.bindings)?
        {
            let key = property_raw_key("rPr", Some(name), &mut last_rank, duplicate);
            parsed.preservation.property_raw_children.push((key, child));
            continue;
        }
        property_raw_key("rPr", Some(name), &mut last_rank, false);
        let parsed_value = root_value(&child, &parsed.bindings)?;
        let recognized = match name {
            "sty" => {
                value.style = parsed_value.as_deref().and_then(MathStyle::parse);
                true
            }
            "nor" => {
                value.normal = parse_present_on_off(parsed_value.as_deref());
                true
            }
            "lit" => {
                value.literal = parse_present_on_off(parsed_value.as_deref());
                true
            }
            "scr" => {
                value.script = parsed_value.as_deref().and_then(MathScriptStyle::parse);
                true
            }
            "brk" => {
                let alignment = root_attribute(&child, &parsed.bindings, b"alnAt")?;
                match alignment {
                    Some(alignment) => match alignment.parse::<u8>() {
                        Ok(0) | Err(_) => false,
                        Ok(alignment) => {
                            value.break_before = true;
                            value.break_alignment = Some(alignment);
                            true
                        }
                    },
                    None => {
                        value.break_before = true;
                        true
                    }
                }
            }
            _ => false,
        };
        if recognized {
            seen.push(name.to_owned());
            preserve_modeled_child(&mut parsed.preservation, name, &child, &parsed.bindings);
            modeled += 1;
        } else {
            parsed.preservation.raw_children.push((modeled, child));
        }
    }
    value.preservation = parsed.preservation;
    Ok(value)
}

fn write_run_properties<W: Write>(writer: &mut Writer<W>, value: &MathRunProperties) -> Result<()> {
    let mut properties = Vec::new();
    if let Some(literal) = value.literal {
        properties.push(Property::boolean("lit", literal));
    }
    if let Some(normal) = value.normal {
        properties.push(Property::boolean("nor", normal));
    }
    if let Some(script) = value.script {
        properties.push(Property::text("scr", script.as_str()));
    }
    if let Some(style) = value.style {
        properties.push(Property::text("sty", style.as_str()));
    }
    if let Some(alignment) = value.break_alignment {
        properties.push(Property::attribute("brk", "alnAt", alignment.to_string()));
    } else if value.break_before {
        properties.push(Property::marker("brk"));
    }
    write_property_container_with_preservation(writer, "rPr", &properties, &value.preservation)
}

enum PropertyValue<'a> {
    Text(&'a str),
    Owned(String),
}
struct Property<'a> {
    name: &'a str,
    attribute: Option<(&'a str, PropertyValue<'a>)>,
}
impl<'a> Property<'a> {
    fn text(name: &'a str, value: &'a str) -> Self {
        Self {
            name,
            attribute: Some(("val", PropertyValue::Text(value))),
        }
    }
    fn boolean(name: &'a str, value: bool) -> Self {
        Self {
            name,
            attribute: Some(("val", PropertyValue::Text(if value { "1" } else { "0" }))),
        }
    }
    fn unsigned(name: &'a str, value: u64) -> Self {
        Self {
            name,
            attribute: Some(("val", PropertyValue::Owned(value.to_string()))),
        }
    }
    fn attribute(name: &'a str, attribute: &'a str, value: String) -> Self {
        Self {
            name,
            attribute: Some((attribute, PropertyValue::Owned(value))),
        }
    }
    fn marker(name: &'a str) -> Self {
        Self {
            name,
            attribute: None,
        }
    }
    fn attribute_value(&self) -> Option<(&str, &str)> {
        let (name, value) = self.attribute.as_ref()?;
        let value = match value {
            PropertyValue::Text(value) => *value,
            PropertyValue::Owned(value) => value.as_str(),
        };
        Some((name, value))
    }
}

fn write_property_container<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    properties: &[Property<'_>],
) -> Result<()> {
    write_property_container_with_preservation(writer, tag, properties, &Preservation::default())
}

fn write_math_text<W: Write>(
    writer: &mut Writer<W>,
    text_value: &str,
    parent: &Preservation,
) -> Result<()> {
    let mut preservation = if let Some(source) = parent
        .modeled_children
        .iter()
        .find(|child| child.name == "t")
    {
        parse_element(&source.raw, &source.bindings)?.preservation
    } else {
        Preservation::default()
    };
    preservation
        .attributes
        .retain(|(name, _)| name != "xml:space");
    let mut start = math_start("t", &preservation, false);
    if text_value.starts_with(char::is_whitespace) || text_value.ends_with(char::is_whitespace) {
        start.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(start.borrow()))?;
    writer.write_event(Event::Text(BytesText::new(text_value)))?;
    writer.write_event(Event::End(BytesEnd::new("m:t")))?;
    Ok(())
}

fn preserve_modeled_child(
    preservation: &mut Preservation,
    name: &str,
    raw: &[u8],
    bindings: &[(String, String)],
) {
    preservation.modeled_children.push(PreservedChild {
        name: name.to_owned(),
        raw: raw.to_vec(),
        bindings: bindings.to_vec(),
    });
}

fn preservation_has_unsupported_content(preservation: &Preservation) -> bool {
    preservation
        .attributes
        .iter()
        .any(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
        || !preservation.raw_children.is_empty()
        || !preservation.property_raw_children.is_empty()
}

fn property_preservation_has_unsupported_content(
    preservation: &Preservation,
    container: &str,
) -> bool {
    preservation_has_unsupported_content(preservation)
        || preservation.modeled_children.iter().any(|child| {
            parse_element(&child.raw, &child.bindings).map_or(true, |parsed| {
                let modeled_attribute = if container == "rPr" && child.name == "brk" {
                    b"alnAt".as_slice()
                } else {
                    b"val".as_slice()
                };
                parsed.preservation.attributes.iter().any(|(name, _)| {
                    if name == "xmlns" || name.starts_with("xmlns:") {
                        return false;
                    }
                    expanded_attribute_name(name.as_bytes(), &parsed.bindings).is_none_or(
                        |(namespace, local)| {
                            namespace != M_NS || local.as_slice() != modeled_attribute
                        },
                    )
                }) || !parsed.children.is_empty()
            })
        })
}

fn property_container_has_unsupported_content(preservation: &Preservation, tag: &str) -> bool {
    let Some(source) = preservation
        .modeled_children
        .iter()
        .find(|child| child.name == tag)
    else {
        return false;
    };
    preserve_property_container(
        &source.raw,
        &source.bindings,
        tag,
        supported_properties(tag),
    )
    .map_or(true, |value| {
        property_preservation_has_unsupported_content(&value, tag)
    })
}

fn math_text_has_unsupported_content(preservation: &Preservation) -> bool {
    let Some(source) = preservation
        .modeled_children
        .iter()
        .find(|child| child.name == "t")
    else {
        return false;
    };
    parse_element(&source.raw, &source.bindings).map_or(true, |parsed| {
        parsed
            .preservation
            .attributes
            .iter()
            .any(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:") && name != "xml:space")
    })
}

fn script_has_unsupported_content(value: &MathScript, property_tag: &str) -> bool {
    preservation_has_unsupported_content(&value.preservation)
        || property_container_has_unsupported_content(&value.preservation, property_tag)
        || value.base.has_unsupported_content()
        || value.script.has_unsupported_content()
}

fn three_argument_script_has_unsupported_content(
    preservation: &Preservation,
    property_tag: &str,
    base: &MathArgument,
    subscript: &MathArgument,
    superscript: &MathArgument,
) -> bool {
    preservation_has_unsupported_content(preservation)
        || property_container_has_unsupported_content(preservation, property_tag)
        || base.has_unsupported_content()
        || subscript.has_unsupported_content()
        || superscript.has_unsupported_content()
}

fn limit_has_unsupported_content(value: &MathLimit, property_tag: &str) -> bool {
    preservation_has_unsupported_content(&value.preservation)
        || property_container_has_unsupported_content(&value.preservation, property_tag)
        || value.base.has_unsupported_content()
        || value.limit.has_unsupported_content()
}

fn has_preserved_modeled_child(preservation: &Preservation, name: &str) -> bool {
    preservation
        .modeled_children
        .iter()
        .any(|child| child.name == name)
}

fn write_leading_property_container<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    properties: &[Property<'_>],
    preservation: &Preservation,
    emit: bool,
) -> Result<usize> {
    let modeled =
        write_leading_property_container_before_raw(writer, tag, properties, preservation, emit)?;
    write_raw_slot(writer, preservation, modeled)?;
    Ok(modeled)
}

fn write_leading_property_container_before_raw<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    properties: &[Property<'_>],
    preservation: &Preservation,
    emit: bool,
) -> Result<usize> {
    let existed_in_source = has_preserved_modeled_child(preservation, tag);
    if emit && existed_in_source {
        write_raw_slot(writer, preservation, 0)?;
        write_property_container_from_parent(writer, tag, properties, preservation)?;
        Ok(1)
    } else {
        if emit {
            write_property_container_from_parent(writer, tag, properties, preservation)?;
        }
        Ok(0)
    }
}

fn write_property_container_from_parent<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    properties: &[Property<'_>],
    parent: &Preservation,
) -> Result<()> {
    let Some(source) = parent
        .modeled_children
        .iter()
        .find(|child| child.name == tag)
    else {
        return write_property_container(writer, tag, properties);
    };
    let preservation = preserve_property_container(
        &source.raw,
        &source.bindings,
        tag,
        supported_properties(tag),
    )?;
    write_property_container_with_preservation(writer, tag, properties, &preservation)
}

fn write_property_container_with_preservation<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    properties: &[Property<'_>],
    preservation: &Preservation,
) -> Result<()> {
    let declare_math = tag == "mathPr";
    let qname = format!("m:{tag}");
    let start = math_start(tag, preservation, declare_math);
    writer.write_event(Event::Start(start.borrow()))?;
    let result = if preservation.property_raw_children.is_empty() {
        (|| {
            write_raw_slot(writer, preservation, 0)?;
            for (index, property) in properties.iter().enumerate() {
                write_property(writer, property, preservation)?;
                if index + 1 < properties.len() {
                    write_raw_slot(writer, preservation, index + 1)?;
                }
            }
            write_raw_tail(writer, preservation, properties.len())
        })()
    } else {
        let full_order = full_property_order(tag);
        let mut emissions = preservation
            .property_raw_children
            .iter()
            .enumerate()
            .map(|(source_order, (key, raw))| (*key, source_order, None, Some(raw)))
            .collect::<Vec<_>>();
        emissions.extend(properties.iter().enumerate().map(|(index, property)| {
            let rank = full_order
                .iter()
                .position(|name| *name == property.name)
                .unwrap_or(full_order.len());
            (rank * 3 + 1, index, Some(property), None)
        }));
        emissions.sort_by_key(|(key, source_order, _, _)| (*key, *source_order));
        for (_, _, property, raw) in emissions {
            if let Some(source) = preservation
                .modeled_children
                .iter()
                .find(|child| property.is_some_and(|property| child.name == property.name))
            {
                write_property_leaf(
                    writer,
                    property.expect("source requires a property"),
                    &source.raw,
                    &source.bindings,
                )?;
            } else if let Some(property) = property {
                write_new_property(writer, property)?;
            } else if let Some(raw) = raw {
                writer.get_mut().write_all(raw)?;
            }
        }
        Ok(())
    };
    result?;
    writer.write_event(Event::End(BytesEnd::new(qname)))?;
    Ok(())
}

fn write_property<W: Write>(
    writer: &mut Writer<W>,
    property: &Property<'_>,
    preservation: &Preservation,
) -> Result<()> {
    if let Some(source) = preservation
        .modeled_children
        .iter()
        .find(|child| child.name == property.name)
    {
        write_property_leaf(writer, property, &source.raw, &source.bindings)
    } else {
        write_new_property(writer, property)
    }
}

fn write_new_property<W: Write>(writer: &mut Writer<W>, property: &Property<'_>) -> Result<()> {
    let mut child = BytesStart::new(format!("m:{}", property.name));
    if let Some((name, value)) = property.attribute_value() {
        child.push_attribute((format!("m:{name}").as_str(), value));
    }
    writer.write_event(Event::Empty(child))?;
    Ok(())
}

fn write_property_leaf<W: Write>(
    writer: &mut Writer<W>,
    property: &Property<'_>,
    raw: &[u8],
    inherited: &[(String, String)],
) -> Result<()> {
    let parsed = parse_element(raw, inherited)?;
    let mut preservation = parsed.preservation;
    let modeled_attribute = property
        .attribute_value()
        .map(|(name, _)| name.as_bytes())
        .unwrap_or(b"alnAt");
    preservation.attributes.retain(|(name, _)| {
        expanded_attribute_name(name.as_bytes(), &parsed.bindings)
            .is_none_or(|(namespace, local)| namespace != M_NS || local != modeled_attribute)
    });
    let mut start = math_start(property.name, &preservation, false);
    if let Some((name, value)) = property.attribute_value() {
        start.push_attribute((format!("m:{name}").as_str(), value));
    }
    if parsed.children.is_empty() {
        writer.write_event(Event::Empty(start))?;
    } else {
        writer.write_event(Event::Start(start.borrow()))?;
        for child in parsed.children {
            writer.get_mut().write_all(&child)?;
        }
        writer.write_event(Event::End(BytesEnd::new(format!("m:{}", property.name))))?;
    }
    Ok(())
}

fn write_raw_tail<W: Write>(
    writer: &mut Writer<W>,
    preservation: &Preservation,
    from_slot: usize,
) -> Result<()> {
    for (_, raw) in preservation
        .raw_children
        .iter()
        .filter(|(position, _)| *position >= from_slot)
    {
        writer.get_mut().write_all(raw)?;
    }
    Ok(())
}

fn supported_properties(tag: &str) -> &'static [&'static str] {
    match tag {
        "rPr" => &["lit", "nor", "scr", "sty", "brk"],
        "fPr" => &["type"],
        "sSubPr" | "sSupPr" | "sSubSupPr" | "sPrePr" => &["alnScr"],
        "radPr" => &["degHide"],
        "mPr" => &["baseJc", "rSp", "cSp"],
        "limLowPr" | "limUppPr" => &[],
        "naryPr" => &["chr", "limLoc", "grow", "subHide", "supHide"],
        "dPr" => &["begChr", "sepChr", "endChr", "grow"],
        "accPr" => &["chr"],
        "oMathParaPr" => &["jc"],
        "mathPr" => &[
            "mathFont",
            "smallFrac",
            "dispDef",
            "lMargin",
            "rMargin",
            "defJc",
            "preSp",
            "postSp",
            "interSp",
            "intraSp",
            "wrapIndent",
            "intLim",
            "naryLim",
        ],
        _ => &[],
    }
}

fn full_property_order(tag: &str) -> &'static [&'static str] {
    match tag {
        "rPr" => &["lit", "nor", "scr", "sty", "brk", "aln"],
        "fPr" => &["type", "ctrlPr"],
        "sSubPr" | "sSupPr" | "sSubSupPr" | "sPrePr" => &["alnScr", "ctrlPr"],
        "radPr" => &["degHide", "ctrlPr"],
        "mPr" => &[
            "baseJc", "plcHide", "rSpRule", "cGpRule", "rSp", "cSp", "cGp", "mcs", "ctrlPr",
        ],
        "limLowPr" | "limUppPr" => &["ctrlPr"],
        "naryPr" => &["chr", "limLoc", "grow", "subHide", "supHide", "ctrlPr"],
        "dPr" => &["begChr", "sepChr", "endChr", "grow", "shp", "ctrlPr"],
        "accPr" => &["chr", "ctrlPr"],
        "oMathParaPr" => &["jc"],
        "mathPr" => &[
            "mathFont",
            "brkBin",
            "brkBinSub",
            "smallFrac",
            "dispDef",
            "lMargin",
            "rMargin",
            "defJc",
            "preSp",
            "postSp",
            "interSp",
            "intraSp",
            "wrapIndent",
            "wrapRight",
            "intLim",
            "naryLim",
        ],
        _ => &[],
    }
}

fn property_raw_key(
    tag: &str,
    local: Option<&str>,
    last_rank: &mut Option<usize>,
    duplicate: bool,
) -> usize {
    if let Some(rank) = local.and_then(|local| {
        full_property_order(tag)
            .iter()
            .position(|name| *name == local)
    }) {
        *last_rank = Some(rank);
        if duplicate {
            rank * 3 + 2
        } else {
            rank * 3 + 1
        }
    } else {
        last_rank.map_or(0, |rank| rank * 3 + 2)
    }
}

fn preserve_property_container(
    raw: &[u8],
    inherited: &[(String, String)],
    tag: &str,
    recognized: &[&str],
) -> Result<Preservation> {
    let mut parsed = parse_element(raw, inherited)?;
    let mut seen = Vec::new();
    let mut last_rank = None;
    for child in parsed.children {
        let local = math_local_name(&child, &parsed.bindings)?;
        if let Some(local) = local
            .as_deref()
            .filter(|local| recognized.contains(local) && !seen.iter().any(|name| name == local))
        {
            seen.push(local.to_owned());
            property_raw_key(tag, Some(local), &mut last_rank, false);
            preserve_modeled_child(&mut parsed.preservation, local, &child, &parsed.bindings);
        } else {
            let duplicate = local
                .as_deref()
                .is_some_and(|local| seen.iter().any(|name| name == local));
            let key = property_raw_key(tag, local.as_deref(), &mut last_rank, duplicate);
            parsed.preservation.property_raw_children.push((key, child));
        }
    }
    Ok(parsed.preservation)
}

fn write_container<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    preservation: &Preservation,
    children: impl FnOnce(&mut Writer<W>) -> Result<()>,
) -> Result<()> {
    let qname = format!("m:{tag}");
    let start = math_start(tag, preservation, false);
    writer.write_event(Event::Start(start.borrow()))?;
    children(writer)?;
    writer.write_event(Event::End(BytesEnd::new(qname)))?;
    Ok(())
}

fn math_start(tag: &str, preservation: &Preservation, declare_math: bool) -> BytesStart<'static> {
    let mut start = BytesStart::new(format!("m:{tag}"));
    if declare_math {
        start.push_attribute(("xmlns:m", M_NS));
        for (prefix, namespace) in &preservation.inherited_bindings {
            let name = if prefix.is_empty() {
                "xmlns".to_owned()
            } else {
                format!("xmlns:{prefix}")
            };
            if prefix != "m"
                && prefix != "xml"
                && !preservation
                    .attributes
                    .iter()
                    .any(|(candidate, _)| candidate == &name)
            {
                start.push_attribute((name.as_str(), namespace.as_str()));
            }
        }
    }
    for (name, value) in &preservation.attributes {
        if name != "xmlns:m" {
            start.push_attribute((name.as_str(), value.as_str()));
        }
    }
    start.into_owned()
}

fn write_raw_slot<W: Write>(
    writer: &mut Writer<W>,
    preservation: &Preservation,
    slot: usize,
) -> Result<()> {
    for (_, raw) in preservation
        .raw_children
        .iter()
        .filter(|(position, _)| *position == slot)
    {
        writer.get_mut().write_all(raw)?;
    }
    Ok(())
}

fn parse_element(raw: &[u8], inherited: &[(String, String)]) -> Result<ParsedElement> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut bindings = inherited.to_vec();
    let mut preservation = Preservation {
        inherited_bindings: inherited.to_vec(),
        ..Preservation::default()
    };
    let mut children = Vec::new();
    let mut inside = false;
    let mut saw_root = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if !inside => {
                saw_root = true;
                merge_bindings(&mut bindings, &element)?;
                preservation.attributes = extra_attributes(&element)?;
                inside = true;
            }
            Event::Empty(element) if !inside => {
                saw_root = true;
                merge_bindings(&mut bindings, &element)?;
                preservation.attributes = extra_attributes(&element)?;
                break;
            }
            Event::Start(element) if inside => {
                children.push(capture_element(&mut reader, &element)?)
            }
            Event::Empty(element) if inside => children.push(capture_empty(&element)?),
            Event::Text(text) if inside && !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                children.push(capture_event(Event::Text(text.into_owned()))?)
            }
            Event::Comment(comment) if inside => {
                children.push(capture_event(Event::Comment(comment.into_owned()))?)
            }
            Event::PI(pi) if inside => children.push(capture_event(Event::PI(pi.into_owned()))?),
            Event::End(_) if inside => break,
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !saw_root {
        return Err(OxmlError::MissingElement("OfficeMath element".to_owned()));
    }
    Ok(ParsedElement {
        preservation,
        bindings,
        children,
    })
}

const MAX_OFFICEMATH_XML_DEPTH: usize = 128;

fn validate_xml_depth(raw: &[u8]) -> Result<()> {
    let mut reader = Reader::from_reader(raw);
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_OFFICEMATH_XML_DEPTH {
                    return Err(OxmlError::InvalidValue(format!(
                        "OfficeMath XML nesting exceeds {MAX_OFFICEMATH_XML_DEPTH} elements"
                    )));
                }
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof if depth == 0 => return Ok(()),
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "OfficeMath closing element".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn valid_expression_shape(raw: &[u8], inherited: &[(String, String)]) -> Result<bool> {
    let Some(root) = math_local_name(raw, inherited)? else {
        return Ok(true);
    };
    let parsed = parse_element(raw, inherited)?;
    let property_tag = match root.as_str() {
        "r" => Some("rPr"),
        "f" => Some("fPr"),
        "sSub" => Some("sSubPr"),
        "sSup" => Some("sSupPr"),
        "sSubSup" => Some("sSubSupPr"),
        "sPre" => Some("sPrePr"),
        "rad" => Some("radPr"),
        "m" => Some("mPr"),
        "limLow" => Some("limLowPr"),
        "limUpp" => Some("limUppPr"),
        "nary" => Some("naryPr"),
        "d" => Some("dPr"),
        "acc" => Some("accPr"),
        _ => None,
    };
    if let Some(property_tag) = property_tag {
        for child in &parsed.children {
            if math_local_name(child, &parsed.bindings)?.as_deref() == Some(property_tag)
                && !property_container_is_valid(child, &parsed.bindings, property_tag)?
            {
                return Ok(false);
            }
        }
    }
    if root == "r" {
        for child in &parsed.children {
            if math_local_name(child, &parsed.bindings)?.as_deref() == Some("t")
                && math_text_has_non_text_nodes(child)?
            {
                return Ok(false);
            }
        }
    }
    let allowed: &[&str] = match root.as_str() {
        "r" => &["rPr", "t"],
        "f" => &["fPr", "num", "den"],
        "sSub" => &["sSubPr", "e", "sub"],
        "sSup" => &["sSupPr", "e", "sup"],
        "sSubSup" => &["sSubSupPr", "e", "sub", "sup"],
        "sPre" => &["sPrePr", "sub", "sup", "e"],
        "rad" => &["radPr", "deg", "e"],
        "m" => &["mPr", "mr"],
        "limLow" => &["limLowPr", "e", "lim"],
        "limUpp" => &["limUppPr", "e", "lim"],
        "nary" => &["naryPr", "sub", "sup", "e"],
        "d" => &["dPr", "e"],
        "acc" => &["accPr", "e"],
        _ => return Ok(true),
    };
    let sequence = parsed
        .children
        .iter()
        .filter_map(|child| math_local_name(child, &parsed.bindings).transpose())
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .filter(|local| allowed.contains(&local.as_str()))
        .collect::<Vec<_>>();
    let valid = match root.as_str() {
        "r" => {
            let text_start = usize::from(sequence.first().is_some_and(|value| value == "rPr"));
            sequence.len() == text_start + 1 && sequence[text_start] == "t"
        }
        "f" => sequence_matches(&sequence, "fPr", &["num", "den"]),
        "sSub" => sequence_matches(&sequence, "sSubPr", &["e", "sub"]),
        "sSup" => sequence_matches(&sequence, "sSupPr", &["e", "sup"]),
        "sSubSup" => sequence_matches(&sequence, "sSubSupPr", &["e", "sub", "sup"]),
        "sPre" => sequence_matches(&sequence, "sPrePr", &["sub", "sup", "e"]),
        "rad" => sequence_matches_with_optional(&sequence, "radPr", &["deg"], &["e"]),
        "limLow" => sequence_matches(&sequence, "limLowPr", &["e", "lim"]),
        "limUpp" => sequence_matches(&sequence, "limUppPr", &["e", "lim"]),
        "nary" => sequence_matches_with_optional(&sequence, "naryPr", &["sub", "sup"], &["e"]),
        "acc" => sequence_matches(&sequence, "accPr", &["e"]),
        "d" => {
            let argument_start = usize::from(sequence.first().is_some_and(|value| value == "dPr"));
            sequence.len() > argument_start
                && sequence[argument_start..].iter().all(|value| value == "e")
        }
        "m" => {
            let row_start = usize::from(sequence.first().is_some_and(|value| value == "mPr"));
            let mut rows_are_valid = true;
            for child in &parsed.children {
                if math_local_name(child, &parsed.bindings)?.as_deref() == Some("mr") {
                    rows_are_valid &= valid_matrix_row(child, &parsed.bindings)?;
                }
            }
            sequence.len() > row_start
                && sequence[row_start..].iter().all(|value| value == "mr")
                && rows_are_valid
        }
        _ => true,
    };
    Ok(valid)
}

fn valid_officemath_paragraph_shape(raw: &[u8], inherited: &[(String, String)]) -> Result<bool> {
    let parsed = parse_element(raw, inherited)?;
    let sequence = parsed
        .children
        .iter()
        .filter_map(|child| math_local_name(child, &parsed.bindings).transpose())
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .filter(|local| local == "oMathParaPr" || local == "oMath")
        .collect::<Vec<_>>();
    let equation_start = usize::from(sequence.first().is_some_and(|value| value == "oMathParaPr"));
    if sequence.len() <= equation_start
        || !sequence[equation_start..]
            .iter()
            .all(|value| value == "oMath")
    {
        return Ok(false);
    }
    for child in &parsed.children {
        if math_local_name(child, &parsed.bindings)?.as_deref() == Some("oMathParaPr")
            && !property_container_is_valid(child, &parsed.bindings, "oMathParaPr")?
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn property_container_is_valid(
    raw: &[u8],
    inherited: &[(String, String)],
    tag: &str,
) -> Result<bool> {
    let parsed = parse_element(raw, inherited)?;
    let supported = supported_properties(tag);
    let mut last_index = None;
    for child in &parsed.children {
        let Some(local) = math_local_name(child, &parsed.bindings)? else {
            continue;
        };
        let Some(index) = supported.iter().position(|name| *name == local) else {
            continue;
        };
        if last_index.is_some_and(|previous| index <= previous)
            || !property_leaf_is_valid(tag, &local, child, &parsed.bindings)?
        {
            return Ok(false);
        }
        last_index = Some(index);
    }
    Ok(true)
}

fn property_leaf_is_valid(
    container: &str,
    property: &str,
    raw: &[u8],
    inherited: &[(String, String)],
) -> Result<bool> {
    let value = root_value(raw, inherited)?;
    Ok(match (container, property) {
        ("rPr", "lit" | "nor")
        | ("sSubPr" | "sSupPr" | "sSubSupPr" | "sPrePr", "alnScr")
        | ("radPr", "degHide")
        | ("naryPr", "grow" | "subHide" | "supHide")
        | ("dPr", "grow") => value
            .as_deref()
            .is_none_or(|value| parse_bool(value).is_some()),
        ("rPr", "scr") => value
            .as_deref()
            .is_some_and(|value| MathScriptStyle::parse(value).is_some()),
        ("rPr", "sty") => value
            .as_deref()
            .is_some_and(|value| MathStyle::parse(value).is_some()),
        ("rPr", "brk") => root_attribute(raw, inherited, b"alnAt")?
            .as_deref()
            .is_none_or(|value| value.parse::<u8>().is_ok_and(|value| value > 0)),
        ("fPr", "type") => value
            .as_deref()
            .is_some_and(|value| FractionType::parse(value).is_some()),
        ("mPr", "baseJc") => value
            .as_deref()
            .is_some_and(|value| MatrixBaseJustification::parse(value).is_some()),
        ("mPr", "rSp") => value
            .as_deref()
            .is_some_and(|value| value.parse::<u16>().is_ok()),
        ("mPr", "cSp") => value
            .as_deref()
            .is_some_and(|value| value.parse::<u32>().is_ok()),
        ("naryPr", "limLoc") => value
            .as_deref()
            .is_some_and(|value| LimitLocation::parse(value).is_some()),
        ("oMathParaPr", "jc") => value
            .as_deref()
            .is_some_and(|value| MathJustification::parse(value).is_some()),
        ("naryPr" | "accPr", "chr") | ("dPr", "begChr" | "sepChr" | "endChr") => value
            .as_deref()
            .is_none_or(|value| value.chars().count() <= 1),
        ("mathPr", "mathFont") => value.as_deref().is_some_and(|value| !value.is_empty()),
        ("mathPr", "smallFrac" | "dispDef") => value
            .as_deref()
            .is_none_or(|value| parse_bool(value).is_some()),
        (
            "mathPr",
            "lMargin" | "rMargin" | "preSp" | "postSp" | "interSp" | "intraSp" | "wrapIndent",
        ) => value
            .as_deref()
            .is_some_and(|value| value.parse::<u32>().is_ok()),
        ("mathPr", "defJc") => value
            .as_deref()
            .is_some_and(|value| MathJustification::parse(value).is_some()),
        ("mathPr", "intLim" | "naryLim") => value
            .as_deref()
            .is_some_and(|value| LimitLocation::parse(value).is_some()),
        _ => true,
    })
}

fn sequence_matches(sequence: &[String], optional: &str, required: &[&str]) -> bool {
    let start = usize::from(sequence.first().is_some_and(|value| value == optional));
    sequence.len() == start + required.len()
        && sequence[start..]
            .iter()
            .zip(required)
            .all(|(actual, expected)| actual == expected)
}

fn sequence_matches_with_optional(
    sequence: &[String],
    property: &str,
    optional: &[&str],
    required: &[&str],
) -> bool {
    let mut index = usize::from(sequence.first().is_some_and(|value| value == property));
    for child in optional {
        if sequence.get(index).is_some_and(|value| value == child) {
            index += 1;
        }
    }
    sequence.len() == index + required.len()
        && sequence[index..]
            .iter()
            .zip(required)
            .all(|(actual, expected)| actual == expected)
}

fn valid_matrix_row(raw: &[u8], inherited: &[(String, String)]) -> Result<bool> {
    let parsed = parse_element(raw, inherited)?;
    let cells = parsed
        .children
        .iter()
        .filter_map(|child| math_local_name(child, &parsed.bindings).transpose())
        .collect::<Result<Vec<_>>>()?;
    Ok(!cells.is_empty() && cells.iter().all(|local| local == "e"))
}

fn element_text(raw: &[u8]) -> Result<String> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut inside = false;
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) if !inside => inside = true,
            Event::Empty(_) if !inside => return Ok(text),
            Event::Start(_) if inside => {
                return Err(OxmlError::UnexpectedElement(
                    "nested element in OfficeMath text".to_owned(),
                ));
            }
            Event::Text(value) if inside => text.push_str(std::str::from_utf8(value.as_ref())?),
            Event::CData(value) if inside => text.push_str(std::str::from_utf8(value.as_ref())?),
            Event::GeneralRef(value) if inside => {
                let name = std::str::from_utf8(value.as_ref())?;
                let entity = format!("&{name};");
                let decoded = quick_xml::escape::unescape(&entity).map_err(|error| {
                    OxmlError::InvalidValue(format!("invalid OfficeMath text entity: {error}"))
                })?;
                text.push_str(&decoded);
            }
            Event::End(_) if inside => return Ok(text),
            Event::Empty(_) if inside => {
                return Err(OxmlError::UnexpectedElement(
                    "nested empty element in OfficeMath text".to_owned(),
                ));
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("OfficeMath text end".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn math_text_has_non_text_nodes(raw: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    let mut inside = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) if !inside => inside = true,
            Event::Empty(_) if !inside => return Ok(false),
            Event::Start(_) | Event::Empty(_) if inside => return Ok(true),
            Event::Comment(_) | Event::PI(_) if inside => return Ok(true),
            Event::End(_) if inside => return Ok(false),
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn capture_empty(element: &BytesStart<'_>) -> Result<Vec<u8>> {
    capture_event(Event::Empty(element.to_owned().into_owned()))
}
fn capture_event(event: Event<'static>) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_event(event)?;
    Ok(output)
}

fn merge_bindings(bindings: &mut Vec<(String, String)>, element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            Some("")
        } else {
            name.strip_prefix(b"xmlns:")
                .and_then(|value| std::str::from_utf8(value).ok())
        };
        if let Some(prefix) = prefix {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                .into_owned();
            bindings.retain(|(candidate, _)| candidate != prefix);
            bindings.push((prefix.to_owned(), value));
        }
    }
    Ok(())
}

pub(crate) fn fixed_math_prefix_is_safe(
    raw: &[u8],
    inherited: &[(String, String)],
) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    let mut root_m_binding = inherited
        .iter()
        .rev()
        .find(|(prefix, _)| prefix == "m")
        .map(|(_, namespace)| namespace.clone());
    let mut saw_root = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                let mut local_m_binding = None;
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    if attribute.key.as_ref() == b"xmlns:m" {
                        local_m_binding = Some(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    element.decoder(),
                                )?
                                .into_owned(),
                        );
                    }
                }
                if !saw_root {
                    if local_m_binding.is_some() {
                        root_m_binding = local_m_binding;
                    }
                    if root_m_binding.as_deref().is_some_and(|value| value != M_NS) {
                        return Ok(false);
                    }
                    saw_root = true;
                } else if local_m_binding
                    .as_deref()
                    .is_some_and(|value| value != M_NS)
                {
                    return Ok(false);
                }
            }
            Event::Eof => return Ok(true),
            _ => {}
        }
        buffer.clear();
    }
}

fn require_safe_fixed_math_prefix(raw: &[u8], inherited: &[(String, String)]) -> Result<()> {
    if fixed_math_prefix_is_safe(raw, inherited)? {
        Ok(())
    } else {
        Err(OxmlError::InvalidValue(
            "OfficeMath source conflicts with the canonical m prefix".to_owned(),
        ))
    }
}

fn extra_attributes(element: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = std::str::from_utf8(attribute.key.as_ref())?.to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        attributes.push((name, value));
    }
    Ok(attributes)
}

fn expanded_name(name: &[u8], bindings: &[(String, String)]) -> Option<(String, Vec<u8>)> {
    let (prefix, local) = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or((&b""[..], name), |at| (&name[..at], &name[at + 1..]));
    let prefix = std::str::from_utf8(prefix).ok()?;
    let namespace = bindings
        .iter()
        .rev()
        .find(|(candidate, _)| candidate == prefix)?
        .1
        .clone();
    Some((namespace, local.to_vec()))
}

fn expanded_attribute_name(
    name: &[u8],
    bindings: &[(String, String)],
) -> Option<(String, Vec<u8>)> {
    if !name.contains(&b':') {
        return None;
    }
    expanded_name(name, bindings)
}

fn root_name(raw: &[u8], inherited: &[(String, String)]) -> Result<Option<(String, Vec<u8>)>> {
    let mut reader = Reader::from_reader(raw);
    let mut bindings = inherited.to_vec();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                merge_bindings(&mut bindings, &element)?;
                return Ok(expanded_name(element.name().as_ref(), &bindings));
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn math_local_name(raw: &[u8], inherited: &[(String, String)]) -> Result<Option<String>> {
    Ok(root_name(raw, inherited)?.and_then(|(namespace, local)| {
        (namespace == M_NS).then(|| String::from_utf8_lossy(&local).into_owned())
    }))
}

fn require_root(raw: &[u8], inherited: &[(String, String)], expected: &[u8]) -> Result<()> {
    if root_name(raw, inherited)?
        .is_some_and(|(namespace, local)| namespace == M_NS && local == expected)
    {
        Ok(())
    } else {
        Err(OxmlError::MissingElement(format!(
            "OfficeMath {} root",
            String::from_utf8_lossy(expected)
        )))
    }
}

fn root_value(raw: &[u8], inherited: &[(String, String)]) -> Result<Option<String>> {
    root_attribute(raw, inherited, b"val")
}

fn root_attribute(
    raw: &[u8],
    inherited: &[(String, String)],
    expected: &[u8],
) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(raw);
    let mut bindings = inherited.to_vec();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                merge_bindings(&mut bindings, &element)?;
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    if expanded_attribute_name(attribute.key.as_ref(), &bindings)
                        .is_some_and(|(namespace, local)| namespace == M_NS && local == expected)
                    {
                        return Ok(Some(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    element.decoder(),
                                )?
                                .into_owned(),
                        ));
                    }
                }
                return Ok(None);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn child_property_value(
    raw: &[u8],
    inherited: &[(String, String)],
    property: &str,
) -> Result<Option<String>> {
    let parsed = parse_element(raw, inherited)?;
    for child in parsed.children {
        if math_local_name(&child, &parsed.bindings)?.as_deref() == Some(property) {
            return root_value(&child, &parsed.bindings);
        }
    }
    Ok(None)
}

fn child_on_off_value(
    raw: &[u8],
    inherited: &[(String, String)],
    property: &str,
) -> Result<Option<bool>> {
    let parsed = parse_element(raw, inherited)?;
    for child in parsed.children {
        if math_local_name(&child, &parsed.bindings)?.as_deref() == Some(property) {
            return Ok(parse_present_on_off(
                root_value(&child, &parsed.bindings)?.as_deref(),
            ));
        }
    }
    Ok(None)
}

fn preserve_unrecognized_property_children(
    raw: &[u8],
    inherited: &[(String, String)],
    recognized: &[&str],
) -> Result<Preservation> {
    let mut parsed = parse_element(raw, inherited)?;
    let mut seen = Vec::new();
    let mut last_rank = None;
    for child in parsed.children {
        let local = math_local_name(&child, &parsed.bindings)?;
        if let Some(local) = local
            .as_deref()
            .filter(|local| recognized.contains(local) && !seen.iter().any(|name| name == local))
        {
            seen.push(local.to_owned());
            property_raw_key("oMathParaPr", Some(local), &mut last_rank, false);
            preserve_modeled_child(&mut parsed.preservation, local, &child, &parsed.bindings);
        } else {
            let duplicate = local
                .as_deref()
                .is_some_and(|local| seen.iter().any(|name| name == local));
            let key = property_raw_key("oMathParaPr", local.as_deref(), &mut last_rank, duplicate);
            parsed.preservation.property_raw_children.push((key, child));
        }
    }
    Ok(parsed.preservation)
}

fn parse_bool(value: &str) -> Option<bool> {
    match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}
fn parse_present_on_off(value: Option<&str>) -> Option<bool> {
    match value {
        Some(value) => parse_bool(value),
        None => Some(true),
    }
}
fn parse_u32(value: Option<String>) -> Option<u32> {
    value.and_then(|value| value.parse().ok())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn complete_corpus() -> CT_OMath {
        let text = |value| MathArgument::text(value);
        CT_OMath::new(vec![
            MathRun::new("x").into(),
            MathExpression::Fraction(MathFraction::new(text("1"), text("2"))),
            MathExpression::Subscript(MathScript::new(text("x"), text("i"))),
            MathExpression::Superscript(MathScript::new(text("x"), text("2"))),
            MathExpression::SubSuperscript(MathSubSuperscript::new(
                text("x"),
                text("i"),
                text("2"),
            )),
            MathExpression::PreSubSuperscript(MathPreSubSuperscript::new(
                text("x"),
                text("i"),
                text("2"),
            )),
            MathExpression::Radical(MathRadical::with_degree(text("3"), text("x"))),
            MathExpression::Matrix(MathMatrix::new(vec![MathMatrixRow::new(vec![
                text("a"),
                text("b"),
            ])])),
            MathExpression::LowerLimit(MathLimit::new(text("lim"), text("0"))),
            MathExpression::UpperLimit(MathLimit::new(text("max"), text("n"))),
            MathExpression::Nary({
                let mut nary = MathNary::new("∑", text("x"));
                nary.subscript = text("i");
                nary.superscript = text("n");
                nary
            }),
            MathExpression::Delimiter(MathDelimiter::new("(", ")", vec![text("x")])),
            MathExpression::Accent(MathAccent::new("̂", text("x"))),
        ])
    }

    #[test]
    fn every_supported_officemath_construct_writes_schema_order_and_reparses() {
        let equation = complete_corpus();
        let xml = equation.to_xml().unwrap();
        let reparsed = CT_OMath::from_xml(&xml).unwrap();
        assert_eq!(reparsed.expressions.len(), equation.expressions.len());
        let xml = String::from_utf8(xml).unwrap();
        assert!(xml.find("<m:fPr>").unwrap() < xml.find("<m:num>").unwrap());
        assert!(xml.find("<m:num>").unwrap() < xml.find("<m:den>").unwrap());
        let nary = &xml[xml.find("<m:nary>").unwrap()..];
        assert!(nary.find("<m:naryPr>").unwrap() < nary.find("<m:sub>").unwrap());
    }

    #[test]
    fn officemath_reader_accepts_aliases_and_writer_uses_fixed_math_prefix() {
        let xml = format!(r#"<z:oMath xmlns:z="{M_NS}"><z:r><z:t>x</z:t></z:r></z:oMath>"#);
        let parsed = CT_OMath::from_xml(xml.as_bytes()).unwrap();
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(output.starts_with(&format!(r#"<m:oMath xmlns:m="{M_NS}""#)));
        assert!(output.contains("<m:r><m:t>x</m:t></m:r>"));
    }

    #[test]
    fn unsupported_officemath_siblings_survive_typed_mutation_byte_for_byte() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer" x:flag="kept"><x:before a="1"/><m:r><m:t>x</m:t></m:r><x:after> raw </x:after></m:oMath>"#
        );
        let mut parsed = CT_OMath::from_xml(xml.as_bytes()).unwrap();
        let MathExpression::Run(run) = &mut parsed.expressions[0] else {
            panic!("run")
        };
        run.text = "y".to_owned();
        let output = parsed.to_xml().unwrap();
        assert!(
            output
                .windows(b"<x:before a=\"1\"/>".len())
                .any(|value| value == b"<x:before a=\"1\"/>")
        );
        assert!(
            output
                .windows(b"<x:after> raw </x:after>".len())
                .any(|value| value == b"<x:after> raw </x:after>")
        );
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains("x:flag=\"kept\"")
        );
    }

    #[test]
    fn inserted_property_containers_do_not_shift_raw_argument_slots() {
        let cases = [
            (
                format!(
                    r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:f><x:before/><m:num/><x:middle/><m:den/><x:after/></m:f></m:oMath>"#
                ),
                vec![
                    "<m:fPr>",
                    "<x:before/>",
                    "<m:num>",
                    "<x:middle/>",
                    "<m:den>",
                    "<x:after/>",
                ],
            ),
            (
                format!(
                    r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:rad><x:before/><m:deg/><x:middle/><m:e/><x:after/></m:rad></m:oMath>"#
                ),
                vec![
                    "<m:radPr>",
                    "<x:before/>",
                    "<m:deg>",
                    "<x:middle/>",
                    "<m:e>",
                    "<x:after/>",
                ],
            ),
            (
                format!(
                    r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:nary><x:before/><m:sub/><x:middle-one/><m:sup/><x:middle-two/><m:e/><x:after/></m:nary></m:oMath>"#
                ),
                vec![
                    "<m:naryPr>",
                    "<x:before/>",
                    "<m:sub>",
                    "<x:middle-one/>",
                    "<m:sup>",
                    "<x:middle-two/>",
                    "<m:e>",
                    "<x:after/>",
                ],
            ),
            (
                format!(
                    r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:d><x:before/><m:e/><x:middle/><m:e/><x:after/></m:d></m:oMath>"#
                ),
                vec![
                    "<m:dPr>",
                    "<x:before/>",
                    "<m:e>",
                    "<x:middle/>",
                    "<m:e>",
                    "<x:after/>",
                ],
            ),
            (
                format!(
                    r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:acc><x:before/><m:e/><x:after/></m:acc></m:oMath>"#
                ),
                vec!["<m:accPr>", "<x:before/>", "<m:e>", "<x:after/>"],
            ),
        ];
        for (source, expected) in cases {
            let parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
            let first = parsed.to_xml().unwrap();
            assert_fragments_in_order(&first, &expected);
            let reopened = CT_OMath::from_xml(&first).unwrap();
            assert_fragments_in_order(&reopened.to_xml().unwrap(), &expected);
        }
    }

    #[test]
    fn unsupported_content_is_observable_without_exposing_preservation_storage() {
        let authored = CT_OMath::new(vec![MathRun::new("x").into()]);
        assert!(!authored.has_unsupported_content());
        assert!(!OfficeMath::Inline(authored).has_unsupported_content());

        let source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:f><m:fPr><x:property/></m:fPr><m:num><x:argument/></m:num><m:den/></m:f></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        assert!(parsed.has_unsupported_content());
        let MathExpression::Fraction(fraction) = &parsed.expressions[0] else {
            panic!("fraction")
        };
        assert!(parsed.expressions[0].has_unsupported_content());
        assert!(fraction.numerator.has_unsupported_content());
        assert!(!fraction.denominator.has_unsupported_content());

        let extended_text = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:r><m:t x:keep="yes">x</m:t></m:r></m:oMath>"#
        );
        let extended_text = CT_OMath::from_xml(extended_text.as_bytes()).unwrap();
        assert!(extended_text.expressions[0].has_unsupported_content());

        let spaced_text = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:r><m:t xml:space="preserve"> x </m:t></m:r></m:oMath>"#
        );
        let spaced_text = CT_OMath::from_xml(spaced_text.as_bytes()).unwrap();
        assert!(!spaced_text.expressions[0].has_unsupported_content());
    }

    #[test]
    fn existing_run_and_display_properties_keep_both_parent_raw_slots() {
        let run_source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:r><x:run-before/><m:rPr/><x:run-after-property/><m:t>x</m:t><x:run-after-text/></m:r></m:oMath>"#
        );
        let run = CT_OMath::from_xml(run_source.as_bytes()).unwrap();
        let first = run.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:run-before/>",
                "<m:rPr>",
                "<x:run-after-property/>",
                "<m:t>",
                "<x:run-after-text/>",
            ],
        );
        let reopened = CT_OMath::from_xml(&first).unwrap();
        assert_fragments_in_order(
            &reopened.to_xml().unwrap(),
            &[
                "<x:run-before/>",
                "<m:rPr>",
                "<x:run-after-property/>",
                "<m:t>",
                "<x:run-after-text/>",
            ],
        );

        let display_source = format!(
            r#"<m:oMathPara xmlns:m="{M_NS}" xmlns:x="urn:producer"><x:display-before/><m:oMathParaPr/><x:display-after-property/><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><x:display-after-equation/></m:oMathPara>"#
        );
        let display = CT_OMathPara::from_xml(display_source.as_bytes()).unwrap();
        let first = display.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:display-before/>",
                "<m:oMathParaPr>",
                "<x:display-after-property/>",
                "<m:oMath>",
                "<x:display-after-equation/>",
            ],
        );
        let reopened = CT_OMathPara::from_xml(&first).unwrap();
        assert_fragments_in_order(
            &reopened.to_xml().unwrap(),
            &[
                "<x:display-before/>",
                "<m:oMathParaPr>",
                "<x:display-after-property/>",
                "<m:oMath>",
                "<x:display-after-equation/>",
            ],
        );
    }

    #[test]
    fn comments_and_processing_instructions_inside_math_text_keep_the_run_opaque() {
        let source = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:r><m:t>x<!--kept--></m:t></m:r><m:r><m:t>y<?kept value?></m:t></m:r></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
        assert!(parsed.has_unsupported_content());
        let first = parsed.to_xml().unwrap();
        let first = String::from_utf8(first).unwrap();
        assert!(first.contains("<m:t>x<!--kept--></m:t>"));
        assert!(first.contains("<m:t>y<?kept value?></m:t>"));
        let reopened = CT_OMath::from_xml(first.as_bytes()).unwrap();
        assert!(reopened.expressions.is_empty());
        assert!(reopened.has_unsupported_content());
        let second = String::from_utf8(reopened.to_xml().unwrap()).unwrap();
        assert!(second.contains("<m:t>x<!--kept--></m:t>"));
        assert!(second.contains("<m:t>y<?kept value?></m:t>"));
    }

    #[test]
    fn unsupported_property_attributes_are_checked_against_their_actual_leaf() {
        let source = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:f><m:fPr><m:type m:val="bar" m:alnAt="4"/></m:fPr><m:num/><m:den/></m:f><m:r><m:rPr><m:brk m:alnAt="4" m:val="1"/></m:rPr><m:t>x</m:t></m:r></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        assert!(parsed.expressions[0].has_unsupported_content());
        assert!(parsed.expressions[1].has_unsupported_content());

        let supported = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num/><m:den/></m:f><m:r><m:rPr><m:brk m:alnAt="4"/></m:rPr><m:t>x</m:t></m:r></m:oMath>"#
        );
        let supported = CT_OMath::from_xml(supported.as_bytes()).unwrap();
        assert!(!supported.expressions[0].has_unsupported_content());
        assert!(!supported.expressions[1].has_unsupported_content());
    }

    #[test]
    fn optional_radical_degree_and_nary_limits_remain_typed_and_omitted() {
        let source = format!(
            r#"<m:oMath xmlns:m="{M_NS}">
                <m:rad><m:e/></m:rad>
                <m:nary><m:e/></m:nary>
                <m:nary><m:sub/><m:e/></m:nary>
                <m:nary><m:sup/><m:e/></m:nary>
                <m:nary><m:sub/><m:sup/><m:e/></m:nary>
            </m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        assert!(matches!(parsed.expressions[0], MathExpression::Radical(_)));
        assert!(
            parsed.expressions[1..]
                .iter()
                .all(|value| matches!(value, MathExpression::Nary(_)))
        );

        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(!output.contains("<m:deg>"));
        assert_eq!(output.matches("<m:sub>").count(), 2);
        assert_eq!(output.matches("<m:sup>").count(), 2);

        let reopened = CT_OMath::from_xml(output.as_bytes()).unwrap();
        assert!(matches!(
            reopened.expressions[0],
            MathExpression::Radical(_)
        ));
        assert!(
            reopened.expressions[1..]
                .iter()
                .all(|value| matches!(value, MathExpression::Nary(_)))
        );
    }

    #[test]
    fn authored_optional_arguments_precede_preserved_slots_without_rebasing_them() {
        let source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer">
                <m:rad><x:rad-slot/><m:e/></m:rad>
                <m:nary><x:nary-slot/><m:e/></m:nary>
            </m:oMath>"#
        );
        let mut parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        let MathExpression::Radical(radical) = &mut parsed.expressions[0] else {
            panic!("radical")
        };
        radical.degree = MathArgument::text("3");
        let MathExpression::Nary(nary) = &mut parsed.expressions[1] else {
            panic!("nary")
        };
        nary.subscript = MathArgument::text("i");

        let output = parsed.to_xml().unwrap();
        assert_fragments_in_order(&output, &["<m:radPr>", "<m:deg>", "<x:rad-slot/>", "<m:e>"]);
        assert_fragments_in_order(
            &output,
            &["<m:naryPr>", "<m:sub>", "<x:nary-slot/>", "<m:e>"],
        );
        let reopened = CT_OMath::from_xml(&output).unwrap();
        let second = reopened.to_xml().unwrap();
        assert_fragments_in_order(&second, &["<m:radPr>", "<m:deg>", "<x:rad-slot/>", "<m:e>"]);
        assert_fragments_in_order(
            &second,
            &["<m:naryPr>", "<m:sub>", "<x:nary-slot/>", "<m:e>"],
        );
    }

    #[test]
    fn officemath_corpus_parses_mutates_saves_and_reopens_without_losing_supported_or_raw_siblings()
    {
        let authored = String::from_utf8(complete_corpus().to_xml().unwrap()).unwrap();
        let source = authored
            .replacen(
                &format!(r#"<m:oMath xmlns:m="{M_NS}">"#),
                &format!(
                    r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><x:before keep="root"/>"#
                ),
                1,
            )
            .replacen("<m:fPr>", "<m:fPr><x:fraction keep=\"property\"/>", 1)
            .replacen("<m:num>", "<m:num><x:numerator keep=\"argument\"/>", 1)
            .replacen("</m:oMath>", "<x:after keep=\"root\"/></m:oMath>", 1);
        let mut parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        let MathExpression::Run(run) = &mut parsed.expressions[0] else {
            panic!("run")
        };
        run.text = "changed".to_owned();
        let second = parsed.to_xml().unwrap();
        let serialized = String::from_utf8(second.clone()).unwrap();
        for raw in [
            r#"<x:before keep="root"/>"#,
            r#"<x:fraction keep="property"/>"#,
            r#"<x:numerator keep="argument"/>"#,
            r#"<x:after keep="root"/>"#,
        ] {
            assert!(serialized.contains(raw));
        }
        let reopened = CT_OMath::from_xml(&second).unwrap();
        assert!(matches!(reopened.expressions[0], MathExpression::Run(_)));
        assert!(matches!(
            reopened.expressions[1],
            MathExpression::Fraction(_)
        ));
        assert!(matches!(
            reopened.expressions[2],
            MathExpression::Subscript(_)
        ));
        assert!(matches!(
            reopened.expressions[3],
            MathExpression::Superscript(_)
        ));
        assert!(matches!(
            reopened.expressions[4],
            MathExpression::SubSuperscript(_)
        ));
        assert!(matches!(
            reopened.expressions[5],
            MathExpression::PreSubSuperscript(_)
        ));
        assert!(matches!(
            reopened.expressions[6],
            MathExpression::Radical(_)
        ));
        assert!(matches!(reopened.expressions[7], MathExpression::Matrix(_)));
        assert!(matches!(
            reopened.expressions[8],
            MathExpression::LowerLimit(_)
        ));
        assert!(matches!(
            reopened.expressions[9],
            MathExpression::UpperLimit(_)
        ));
        assert!(matches!(reopened.expressions[10], MathExpression::Nary(_)));
        assert!(matches!(
            reopened.expressions[11],
            MathExpression::Delimiter(_)
        ));
        assert!(matches!(
            reopened.expressions[12],
            MathExpression::Accent(_)
        ));
        let MathExpression::Run(run) = &reopened.expressions[0] else {
            panic!("run")
        };
        assert_eq!(run.text, "changed");
        let reopened_serialized = String::from_utf8(reopened.to_xml().unwrap()).unwrap();
        assert_fragments_in_order(
            reopened_serialized.as_bytes(),
            &[
                r#"<x:before keep="root"/>"#,
                "<m:r>",
                "<m:f>",
                "<m:acc>",
                "</m:acc>",
                r#"<x:after keep="root"/>"#,
            ],
        );
        let fraction_start = reopened_serialized.find("<m:f>").unwrap();
        let fraction_end = reopened_serialized[fraction_start..]
            .find("</m:f>")
            .map(|offset| fraction_start + offset + "</m:f>".len())
            .unwrap();
        let fraction = &reopened_serialized[fraction_start..fraction_end];
        assert_fragments_in_order(
            fraction.as_bytes(),
            &[
                "<m:fPr>",
                r#"<x:fraction keep="property"/>"#,
                "<m:type",
                "</m:fPr>",
                "<m:num>",
                r#"<x:numerator keep="argument"/>"#,
                "<m:r>",
                "</m:num>",
                "<m:den>",
            ],
        );
    }

    #[test]
    fn shortening_repeated_math_children_emits_every_unreached_raw_slot() {
        let root_source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><x:root-zero/><m:r><m:t>A</m:t></m:r><x:root-one/><m:f><m:num><x:arg-zero/><m:r><m:t>N1</m:t></m:r><x:arg-one/><m:r><m:t>N2</m:t></m:r><x:arg-two/></m:num><m:den/></m:f><x:root-two/><m:r><m:t>Z</m:t></m:r><x:root-three/></m:oMath>"#
        );
        let mut root = CT_OMath::from_xml(root_source.as_bytes()).unwrap();
        root.expressions.pop();
        let MathExpression::Fraction(fraction) = &mut root.expressions[1] else {
            panic!("fraction")
        };
        fraction.numerator.expressions.pop();
        let first = root.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:root-zero/>",
                "<m:t>A</m:t>",
                "<x:root-one/>",
                "<m:f>",
                "<x:root-two/>",
                "<x:root-three/>",
            ],
        );
        assert_fragments_in_order(
            &first,
            &[
                "<m:num>",
                "<x:arg-zero/>",
                "<m:t>N1</m:t>",
                "<x:arg-one/>",
                "<x:arg-two/>",
                "</m:num>",
            ],
        );
        let reopened = CT_OMath::from_xml(&first).unwrap();
        let second = reopened.to_xml().unwrap();
        assert_fragments_in_order(
            &second,
            &[
                "<x:root-zero/>",
                "<m:t>A</m:t>",
                "<x:root-one/>",
                "<m:f>",
                "<x:root-two/>",
                "<x:root-three/>",
            ],
        );
        assert_fragments_in_order(
            &second,
            &[
                "<m:num>",
                "<x:arg-zero/>",
                "<m:t>N1</m:t>",
                "<x:arg-one/>",
                "<x:arg-two/>",
                "</m:num>",
            ],
        );

        let display_source = format!(
            r#"<m:oMathPara xmlns:m="{M_NS}" xmlns:x="urn:producer"><x:display-zero/><m:oMath><m:r><m:t>A</m:t></m:r></m:oMath><x:display-one/><m:oMath><m:r><m:t>B</m:t></m:r></m:oMath><x:display-two/></m:oMathPara>"#
        );
        let mut display = CT_OMathPara::from_xml(display_source.as_bytes()).unwrap();
        display.equations.pop();
        let first = display.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:display-zero/>",
                "<m:t>A</m:t>",
                "<x:display-one/>",
                "<x:display-two/>",
            ],
        );
        let reopened = CT_OMathPara::from_xml(&first).unwrap();
        assert_fragments_in_order(
            &reopened.to_xml().unwrap(),
            &[
                "<x:display-zero/>",
                "<m:t>A</m:t>",
                "<x:display-one/>",
                "<x:display-two/>",
            ],
        );

        let matrix_source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:m><x:matrix-zero/><m:mr><x:cell-zero/><m:e/><x:cell-one/><m:e/><x:cell-two/></m:mr><x:matrix-one/><m:mr><m:e/></m:mr><x:matrix-two/></m:m></m:oMath>"#
        );
        let mut matrix_equation = CT_OMath::from_xml(matrix_source.as_bytes()).unwrap();
        let MathExpression::Matrix(matrix) = &mut matrix_equation.expressions[0] else {
            panic!("matrix")
        };
        matrix.rows.pop();
        matrix.rows[0].cells.pop();
        let first = matrix_equation.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:matrix-zero/>",
                "<m:mr>",
                "<x:matrix-one/>",
                "<x:matrix-two/>",
            ],
        );
        assert_fragments_in_order(
            &first,
            &["<x:cell-zero/>", "<m:e>", "<x:cell-one/>", "<x:cell-two/>"],
        );
        let reopened = CT_OMath::from_xml(&first).unwrap();
        let second = reopened.to_xml().unwrap();
        assert_fragments_in_order(
            &second,
            &[
                "<x:matrix-zero/>",
                "<m:mr>",
                "<x:matrix-one/>",
                "<x:matrix-two/>",
            ],
        );
        assert_fragments_in_order(
            &second,
            &["<x:cell-zero/>", "<m:e>", "<x:cell-one/>", "<x:cell-two/>"],
        );

        let delimiter_source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:d><x:delimiter-zero/><m:e/><x:delimiter-one/><m:e/><x:delimiter-two/></m:d></m:oMath>"#
        );
        let mut delimiter_equation = CT_OMath::from_xml(delimiter_source.as_bytes()).unwrap();
        let MathExpression::Delimiter(delimiter) = &mut delimiter_equation.expressions[0] else {
            panic!("delimiter")
        };
        delimiter.arguments.pop();
        let first = delimiter_equation.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:delimiter-zero/>",
                "<m:e>",
                "<x:delimiter-one/>",
                "<x:delimiter-two/>",
            ],
        );
        let reopened = CT_OMath::from_xml(&first).unwrap();
        assert_fragments_in_order(
            &reopened.to_xml().unwrap(),
            &[
                "<x:delimiter-zero/>",
                "<m:e>",
                "<x:delimiter-one/>",
                "<x:delimiter-two/>",
            ],
        );
    }

    #[test]
    fn expression_vector_edits_keep_raw_slots_at_ordinal_boundaries() {
        let source = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><x:slot-zero/><m:r><m:t>A</m:t></m:r><x:slot-one/><m:r><m:t>B</m:t></m:r><x:slot-two/></m:oMath>"#
        );
        let mut parsed = CT_OMath::from_xml(source.as_bytes()).unwrap();
        parsed.expressions.swap(0, 1);
        parsed
            .expressions
            .insert(0, MathExpression::Run(MathRun::new("C")));
        let first = parsed.to_xml().unwrap();
        assert_fragments_in_order(
            &first,
            &[
                "<x:slot-zero/>",
                "<m:t>C</m:t>",
                "<x:slot-one/>",
                "<m:t>B</m:t>",
                "<x:slot-two/>",
                "<m:t>A</m:t>",
            ],
        );
        let reopened = CT_OMath::from_xml(&first).unwrap();
        assert_fragments_in_order(
            &reopened.to_xml().unwrap(),
            &[
                "<x:slot-zero/>",
                "<m:t>C</m:t>",
                "<x:slot-one/>",
                "<m:t>B</m:t>",
                "<x:slot-two/>",
                "<m:t>A</m:t>",
            ],
        );
    }

    fn assert_fragments_in_order(xml: &[u8], fragments: &[&str]) {
        let xml = std::str::from_utf8(xml).unwrap();
        let mut offset = 0usize;
        for fragment in fragments {
            let position = xml[offset..]
                .find(fragment)
                .unwrap_or_else(|| panic!("missing {fragment:?} after byte {offset} in {xml}"));
            offset += position + fragment.len();
        }
    }

    #[test]
    fn officemath_rejects_excessive_nesting_and_invalid_text_bytes() {
        let mut deeply_nested = format!(r#"<m:oMath xmlns:m="{M_NS}">"#).into_bytes();
        for _ in 0..MAX_OFFICEMATH_XML_DEPTH {
            deeply_nested.extend_from_slice(b"<x:a xmlns:x=\"urn:test\">");
        }
        for _ in 0..MAX_OFFICEMATH_XML_DEPTH {
            deeply_nested.extend_from_slice(b"</x:a>");
        }
        deeply_nested.extend_from_slice(b"</m:oMath>");
        assert!(matches!(
            CT_OMath::from_xml(&deeply_nested),
            Err(OxmlError::InvalidValue(_))
        ));

        let mut invalid_text = format!(r#"<m:oMath xmlns:m="{M_NS}"><m:r><m:t>"#).into_bytes();
        invalid_text.push(0xff);
        invalid_text.extend_from_slice(b"</m:t></m:r></m:oMath>");
        assert!(CT_OMath::from_xml(&invalid_text).is_err());
    }

    #[test]
    fn malformed_supported_construct_remains_raw_instead_of_becoming_a_default() {
        let malformed = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:f data="kept"><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(malformed.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(r#"<m:f data="kept"><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>"#)
        );
    }

    #[test]
    fn modeled_property_containers_preserve_unknown_attributes_and_children() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:f><m:fPr data="container"><m:type m:val="bar" data="leaf"><x:hint/></m:type><x:extension value="kept"/></m:fPr><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>"#
        );
        let mut parsed = CT_OMath::from_xml(xml.as_bytes()).unwrap();
        let MathExpression::Fraction(fraction) = &mut parsed.expressions[0] else {
            panic!("fraction")
        };
        fraction.fraction_type = FractionType::Linear;
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<m:fPr data="container">"#));
        assert!(output.contains(r#"<m:type data="leaf" m:val="lin"><x:hint/></m:type>"#));
        assert!(output.contains(r#"<x:extension value="kept"/>"#));
    }

    #[test]
    fn global_math_properties_preserve_modeled_leaf_extensions() {
        let xml = format!(
            r#"<z:mathPr xmlns:z="{M_NS}" xmlns:x="urn:producer" data="root"><z:mathFont z:val="Cambria Math" data="leaf"><x:hint/></z:mathFont><x:extension/></z:mathPr>"#
        );
        let mut properties = MathProperties::from_raw(xml.as_bytes(), &[]).unwrap();
        properties.math_font = Some("STIX Two Math".to_owned());
        let mut writer = Writer::new(Vec::new());
        properties.write_xml(&mut writer).unwrap();
        let output = String::from_utf8(writer.into_inner()).unwrap();
        assert!(output.starts_with(&format!(r#"<m:mathPr xmlns:m="{M_NS}""#)));
        assert!(output.contains("data=\"root\""));
        assert!(
            output.contains(
                r#"<m:mathFont data="leaf" m:val="STIX Two Math"><x:hint/></m:mathFont>"#
            )
        );
        assert!(output.contains("<x:extension/>"));
    }

    #[test]
    fn property_groups_write_schema_order_and_preserve_math_text_extensions() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:x="urn:producer"><m:r><m:rPr><m:lit/><m:nor m:val="0"/><m:scr m:val="fraktur"/><m:sty m:val="b"/><m:brk m:alnAt="4"/></m:rPr><m:t x:keep="yes">x</m:t></m:r><m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="subSup"/><m:grow/><m:subHide m:val="0"/><m:supHide m:val="0"/></m:naryPr><m:sub/><m:sup/><m:e/></m:nary><m:d><m:dPr><m:begChr m:val="["/><m:sepChr m:val="|"/><m:endChr m:val="]"/><m:grow/></m:dPr><m:e/></m:d></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(xml.as_bytes()).unwrap();
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        let run_properties = &output[output.find("<m:rPr>").unwrap()..];
        for pair in [
            ("lit", "nor"),
            ("nor", "scr"),
            ("scr", "sty"),
            ("sty", "brk"),
        ] {
            assert!(
                run_properties.find(&format!("<m:{}", pair.0)).unwrap()
                    < run_properties.find(&format!("<m:{}", pair.1)).unwrap()
            );
        }
        assert!(output.contains(r#"<m:brk m:alnAt="4"/>"#));
        assert!(output.contains(r#"<m:t x:keep="yes">x</m:t>"#));

        let nary = &output[output.find("<m:naryPr>").unwrap()..];
        for pair in [
            ("chr", "limLoc"),
            ("limLoc", "grow"),
            ("grow", "subHide"),
            ("subHide", "supHide"),
        ] {
            assert!(
                nary.find(&format!("<m:{}", pair.0)).unwrap()
                    < nary.find(&format!("<m:{}", pair.1)).unwrap()
            );
        }
        let delimiter = &output[output.find("<m:dPr>").unwrap()..];
        assert!(delimiter.find("<m:begChr").unwrap() < delimiter.find("<m:sepChr").unwrap());
        assert!(delimiter.find("<m:sepChr").unwrap() < delimiter.find("<m:endChr").unwrap());
        assert!(delimiter.find("<m:endChr").unwrap() < delimiter.find("<m:grow").unwrap());
    }

    #[test]
    fn matrix_base_justification_and_nary_absence_use_schema_defaults() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:m><m:mPr><m:baseJc m:val="top"/></m:mPr><m:mr><m:e/></m:mr></m:m><m:nary><m:sub/><m:sup/><m:e/></m:nary></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(xml.as_bytes()).unwrap();
        let MathExpression::Matrix(matrix) = &parsed.expressions[0] else {
            panic!("matrix")
        };
        assert_eq!(
            matrix.properties.base_justification,
            Some(MatrixBaseJustification::Top)
        );
        let MathExpression::Nary(nary) = &parsed.expressions[1] else {
            panic!("nary")
        };
        assert!(!nary.hide_subscript);
        assert!(!nary.hide_superscript);
        assert_eq!(nary.character, "∫");
    }

    #[test]
    fn duplicate_math_text_keeps_the_run_opaque() {
        let xml =
            format!(r#"<m:oMath xmlns:m="{M_NS}"><m:r><m:t>a</m:t><m:t>b</m:t></m:r></m:oMath>"#);
        let parsed = CT_OMath::from_xml(xml.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<m:r><m:t>a</m:t><m:t>b</m:t></m:r>"#));
    }

    #[test]
    fn malformed_and_duplicate_property_leaves_remain_raw() {
        let malformed_fraction = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:f><m:fPr><m:type m:val="other"/></m:fPr><m:num/><m:den/></m:f></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(malformed_fraction.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
        assert!(
            String::from_utf8(parsed.to_xml().unwrap())
                .unwrap()
                .contains(r#"<m:type m:val="other"/>"#)
        );

        let global = format!(
            r#"<m:mathPr xmlns:m="{M_NS}"><m:mathFont m:val="Cambria Math"/><m:mathFont m:val="duplicate"/><m:defJc m:val="sideways"/></m:mathPr>"#
        );
        let properties = MathProperties::from_raw(global.as_bytes(), &[]).unwrap();
        assert_eq!(properties.math_font.as_deref(), Some("Cambria Math"));
        assert_eq!(properties.justification, None);
        let mut writer = Writer::new(Vec::new());
        properties.write_xml(&mut writer).unwrap();
        let output = String::from_utf8(writer.into_inner()).unwrap();
        assert!(output.contains(r#"<m:mathFont m:val="duplicate"/>"#));
        assert!(output.contains(r#"<m:defJc m:val="sideways"/>"#));

        let negative_margin =
            format!(r#"<m:mathPr xmlns:m="{M_NS}"><m:lMargin m:val="-1"/></m:mathPr>"#);
        let properties = MathProperties::from_raw(negative_margin.as_bytes(), &[]).unwrap();
        assert_eq!(properties.left_margin, None);
        let mut writer = Writer::new(Vec::new());
        properties.write_xml(&mut writer).unwrap();
        assert!(
            String::from_utf8(writer.into_inner())
                .unwrap()
                .contains(r#"<m:lMargin m:val="-1"/>"#)
        );

        let negative_matrix_spacing = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:m><m:mPr><m:rSp m:val="-1"/></m:mPr><m:mr><m:e/></m:mr></m:m></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(negative_matrix_spacing.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
    }

    #[test]
    fn malformed_display_and_nested_or_truncated_text_fail_closed() {
        let empty_display = format!(r#"<m:oMathPara xmlns:m="{M_NS}"/>"#);
        assert!(CT_OMathPara::from_xml(empty_display.as_bytes()).is_err());
        assert!(
            OfficeMath::from_raw(empty_display.as_bytes(), &[])
                .unwrap()
                .is_none()
        );

        let nested = format!(
            r#"<m:oMath xmlns:m="{M_NS}"><m:r><m:t><x:empty xmlns:x="urn:test"/></m:t></m:r></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(nested.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
        assert!(parsed.has_unsupported_content());
        let first = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert_eq!(first, nested);
        let reopened = CT_OMath::from_xml(first.as_bytes()).unwrap();
        assert!(reopened.expressions.is_empty());
        assert!(reopened.has_unsupported_content());
        assert_eq!(
            String::from_utf8(reopened.to_xml().unwrap()).unwrap(),
            nested
        );
        let truncated = format!(r#"<m:oMath xmlns:m="{M_NS}"><m:r><m:t>x"#);
        assert!(CT_OMath::from_xml(truncated.as_bytes()).is_err());
    }

    #[test]
    fn standalone_output_replays_inherited_prefixes_needed_by_raw_xml() {
        let raw =
            format!(r#"<m:oMath xmlns:m="{M_NS}"><x:before/><m:r><m:t>x</m:t></m:r></m:oMath>"#);
        let inherited = vec![("x".to_owned(), "urn:producer".to_owned())];
        let parsed = CT_OMath::from_raw(raw.as_bytes(), &inherited).unwrap();
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"xmlns:x="urn:producer""#));
        assert!(output.contains("<x:before/>"));
    }

    #[test]
    fn conflicting_m_bindings_fail_closed_and_unqualified_attributes_stay_raw() {
        let aliased = format!(
            r#"<q:oMath xmlns:q="{M_NS}" xmlns:m="urn:producer"><m:opaque/><q:r><q:t>x</q:t></q:r></q:oMath>"#
        );
        assert!(CT_OMath::from_xml(aliased.as_bytes()).is_err());
        assert!(
            OfficeMath::from_raw(aliased.as_bytes(), &[])
                .unwrap()
                .is_none()
        );

        let unqualified = format!(
            r#"<oMath xmlns="{M_NS}" xmlns:m="{M_NS}"><f><fPr><type val="lin"/></fPr><num/><den/></f></oMath>"#
        );
        let parsed = CT_OMath::from_xml(unqualified.as_bytes()).unwrap();
        assert!(parsed.expressions.is_empty());
        assert!(
            String::from_utf8(parsed.to_xml().unwrap())
                .unwrap()
                .contains(r#"<type val="lin"/>"#)
        );
    }

    #[test]
    fn newly_authored_properties_precede_later_preserved_schema_children() {
        let raw = format!(
            r#"<m:oMath xmlns:m="{M_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><m:f><m:fPr><m:ctrlPr><w:rPr/></m:ctrlPr></m:fPr><m:num/><m:den/></m:f></m:oMath>"#
        );
        let parsed = CT_OMath::from_xml(raw.as_bytes()).unwrap();
        let output = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        let properties = &output[output.find("<m:fPr").unwrap()..];
        assert!(properties.find("<m:type").unwrap() < properties.find("<m:ctrlPr").unwrap());

        let raw = format!(
            r#"<m:mathPr xmlns:m="{M_NS}"><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/></m:mathPr>"#
        );
        let mut parsed = MathProperties::from_raw(raw.as_bytes(), &[]).unwrap();
        parsed.small_fraction = Some(true);
        let mut writer = Writer::new(Vec::new());
        parsed.write_xml(&mut writer).unwrap();
        let output = String::from_utf8(writer.into_inner()).unwrap();
        assert!(output.find("<m:mathFont").unwrap() < output.find("<m:brkBin").unwrap());
        assert!(output.find("<m:brkBin").unwrap() < output.find("<m:smallFrac").unwrap());
    }
}
