use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::f64::consts::PI;
use std::fmt;
use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::xml::{get_attr, local_name, matches_local_name};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::order::OrderedRawChildren;
use crate::preset_shape_data::preset_shape_definition;

const ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
const QUARTER_CIRCLE: f64 = 90.0 * ANGLE_UNITS_PER_DEGREE;
const MAX_ARC_SEGMENTS: usize = 4_096;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum GuideOp {
    MulDiv,
    AddSub,
    AddDiv,
    IfElse,
    Abs,
    At2,
    Cat2,
    Cos,
    Max,
    Min,
    Mod,
    Pin,
    Sat2,
    Sin,
    Sqrt,
    Tan,
    Val,
}

impl GuideOp {
    pub fn parse(token: &str) -> Result<Self, GeometryError> {
        match token {
            "*/" => Ok(Self::MulDiv),
            "+-" => Ok(Self::AddSub),
            "+/" => Ok(Self::AddDiv),
            "?:" => Ok(Self::IfElse),
            "abs" => Ok(Self::Abs),
            "at2" => Ok(Self::At2),
            "cat2" => Ok(Self::Cat2),
            "cos" => Ok(Self::Cos),
            "max" => Ok(Self::Max),
            "min" => Ok(Self::Min),
            "mod" => Ok(Self::Mod),
            "pin" => Ok(Self::Pin),
            "sat2" => Ok(Self::Sat2),
            "sin" => Ok(Self::Sin),
            "sqrt" => Ok(Self::Sqrt),
            "tan" => Ok(Self::Tan),
            "val" => Ok(Self::Val),
            _ => Err(GeometryError::UnknownGuideOperation(token.to_owned())),
        }
    }

    fn argument_count(self) -> usize {
        match self {
            Self::Abs | Self::Sqrt | Self::Val => 1,
            Self::At2 | Self::Cos | Self::Max | Self::Min | Self::Sin | Self::Tan => 2,
            Self::MulDiv
            | Self::AddSub
            | Self::AddDiv
            | Self::IfElse
            | Self::Cat2
            | Self::Mod
            | Self::Pin
            | Self::Sat2 => 3,
        }
    }

    fn token(self) -> &'static str {
        match self {
            Self::MulDiv => "*/",
            Self::AddSub => "+-",
            Self::AddDiv => "+/",
            Self::IfElse => "?:",
            Self::Abs => "abs",
            Self::At2 => "at2",
            Self::Cat2 => "cat2",
            Self::Cos => "cos",
            Self::Max => "max",
            Self::Min => "min",
            Self::Mod => "mod",
            Self::Pin => "pin",
            Self::Sat2 => "sat2",
            Self::Sin => "sin",
            Self::Sqrt => "sqrt",
            Self::Tan => "tan",
            Self::Val => "val",
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum GuideOperand {
    Literal(f64),
    Guide(String),
}

impl GuideOperand {
    pub fn parse(value: &str) -> Result<Self, GeometryError> {
        match value.parse::<f64>() {
            Ok(value) if value.is_finite() => Ok(Self::Literal(value)),
            Ok(_) => Err(GeometryError::NonFiniteValue(value.to_owned())),
            Err(_) if value.is_empty() => Err(GeometryError::EmptyGuideOperand),
            Err(_) => Ok(Self::Guide(value.to_owned())),
        }
    }
}

impl From<f64> for GuideOperand {
    fn from(value: f64) -> Self {
        Self::Literal(value)
    }
}

impl From<&str> for GuideOperand {
    fn from(value: &str) -> Self {
        Self::Guide(value.to_owned())
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Guide {
    pub name: String,
    pub op: GuideOp,
    pub args: Vec<GuideOperand>,
}

impl Guide {
    pub fn parse(name: impl Into<String>, formula: &str) -> Result<Self, GeometryError> {
        let mut parts = formula.split_ascii_whitespace();
        let token = parts.next().ok_or(GeometryError::EmptyGuideFormula)?;
        let op = GuideOp::parse(token)?;
        let args = parts
            .map(GuideOperand::parse)
            .collect::<Result<Vec<_>, _>>()?;
        let expected = op.argument_count();
        if args.len() != expected {
            return Err(GeometryError::WrongArgumentCount {
                operation: token.to_owned(),
                expected,
                actual: args.len(),
            });
        }
        Ok(Self {
            name: name.into(),
            op,
            args,
        })
    }

    fn formula(&self) -> String {
        let mut formula = self.op.token().to_owned();
        for argument in &self.args {
            formula.push(' ');
            match argument {
                GuideOperand::Literal(value) => formula.push_str(&value.to_string()),
                GuideOperand::Guide(name) => formula.push_str(name),
            }
        }
        formula
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum PathCommand {
    MoveTo {
        x: GuideOperand,
        y: GuideOperand,
    },
    LineTo {
        x: GuideOperand,
        y: GuideOperand,
    },
    CubicTo {
        x1: GuideOperand,
        y1: GuideOperand,
        x2: GuideOperand,
        y2: GuideOperand,
        x: GuideOperand,
        y: GuideOperand,
    },
    ArcTo {
        width_radius: GuideOperand,
        height_radius: GuideOperand,
        start_angle: GuideOperand,
        sweep_angle: GuideOperand,
    },
    Close,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum EvaluatedPathCommand {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    Close,
}

#[derive(Clone, Debug, PartialEq)]
pub enum GeometryError {
    Xml(String),
    UnexpectedElement(String),
    MissingAttribute {
        element: String,
        attribute: String,
    },
    InvalidAttribute {
        element: String,
        attribute: String,
        value: String,
    },
    MissingPathList,
    MissingPathDimensions,
    EmptyGuideFormula,
    EmptyGuideOperand,
    UnknownPreset(String),
    UnknownGuideOperation(String),
    WrongArgumentCount {
        operation: String,
        expected: usize,
        actual: usize,
    },
    UnknownGuide(String),
    DuplicateGuide(String),
    UnknownAdjustOverride(String),
    DivisionByZero,
    NonFiniteValue(String),
    PathHasNoCurrentPoint,
    InvalidArcRadius,
    ArcSweepTooLarge,
}

impl fmt::Display for GeometryError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xml(error) => formatter.write_str(error),
            Self::UnexpectedElement(element) => {
                write!(formatter, "unexpected custom geometry element: {element}")
            }
            Self::MissingAttribute { element, attribute } => {
                write!(formatter, "DrawingML {element} requires @{attribute}")
            }
            Self::InvalidAttribute {
                element,
                attribute,
                value,
            } => write!(
                formatter,
                "DrawingML {element} has invalid @{attribute}: {value}"
            ),
            Self::MissingPathList => formatter.write_str("custom geometry requires a path list"),
            Self::MissingPathDimensions => {
                formatter.write_str("custom geometry path requires width and height")
            }
            Self::EmptyGuideFormula => formatter.write_str("empty guide formula"),
            Self::EmptyGuideOperand => formatter.write_str("empty guide operand"),
            Self::UnknownPreset(preset) => {
                write!(formatter, "unknown DrawingML preset geometry: {preset}")
            }
            Self::UnknownGuideOperation(operation) => {
                write!(formatter, "unknown guide operation: {operation}")
            }
            Self::WrongArgumentCount {
                operation,
                expected,
                actual,
            } => write!(
                formatter,
                "guide operation {operation} expects {expected} arguments, got {actual}"
            ),
            Self::UnknownGuide(name) => write!(formatter, "unknown guide: {name}"),
            Self::DuplicateGuide(name) => write!(formatter, "duplicate guide: {name}"),
            Self::UnknownAdjustOverride(name) => {
                write!(formatter, "unknown adjust override: {name}")
            }
            Self::DivisionByZero => formatter.write_str("division by zero"),
            Self::NonFiniteValue(context) => write!(formatter, "non-finite value: {context}"),
            Self::PathHasNoCurrentPoint => {
                formatter.write_str("path command requires a current point")
            }
            Self::InvalidArcRadius => formatter.write_str("arc radii must be positive"),
            Self::ArcSweepTooLarge => formatter.write_str("arc sweep requires too many segments"),
        }
    }
}

impl Error for GeometryError {}

impl From<OxmlError> for GeometryError {
    fn from(error: OxmlError) -> Self {
        Self::Xml(error.to_string())
    }
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_AdjPoint2D {
    pub x: GuideOperand,
    pub y: GuideOperand,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_GeomRect {
    pub left: GuideOperand,
    pub top: GuideOperand,
    pub right: GuideOperand,
    pub bottom: GuideOperand,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub enum CT_Path2DCommand {
    MoveTo(CT_AdjPoint2D),
    LineTo(CT_AdjPoint2D),
    CubicTo {
        control_1: CT_AdjPoint2D,
        control_2: CT_AdjPoint2D,
        end: CT_AdjPoint2D,
    },
    ArcTo {
        width_radius: GuideOperand,
        height_radius: GuideOperand,
        start_angle: GuideOperand,
        sweep_angle: GuideOperand,
    },
    Close,
}

#[derive(Clone, Debug, PartialEq)]
struct PathCommandRecord {
    command: CT_Path2DCommand,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_Path2D {
    pub width: Option<f64>,
    pub height: Option<f64>,
    pub fill: Option<String>,
    pub stroke: Option<bool>,
    pub extrusion_ok: Option<bool>,
    commands: Vec<PathCommandRecord>,
    raw_children: OrderedRawChildren,
}

impl CT_Path2D {
    pub fn commands(&self) -> impl Iterator<Item = &CT_Path2DCommand> {
        self.commands.iter().map(|record| &record.command)
    }
}

#[derive(Clone, Debug, PartialEq)]
struct GuideList {
    guides: Vec<Guide>,
    guide_raw_children: Vec<OrderedRawChildren>,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Debug, PartialEq)]
struct PathList {
    paths: Vec<CT_Path2D>,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_CustomGeometry2D {
    adjust_values: Option<GuideList>,
    guides: Option<GuideList>,
    pub text_rectangle: Option<CT_GeomRect>,
    path_list: PathList,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_PresetGeometry2D {
    pub preset: String,
    adjust_values: Option<GuideList>,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct EvaluatedTextRectangle {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

#[derive(Clone, Debug, PartialEq)]
pub struct EvaluatedCustomGeometry {
    pub paths: Vec<Vec<EvaluatedPathCommand>>,
    pub text_rectangle: Option<EvaluatedTextRectangle>,
}

impl CT_CustomGeometry2D {
    /// Parses one complete `a:custGeom` element with any namespace prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self, GeometryError> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| GeometryError::Xml(error.to_string()))?
            {
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"custGeom") =>
                {
                    return Self::from_element(&mut reader, &element);
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"custGeom") =>
                {
                    return Err(GeometryError::MissingPathList);
                }
                Event::Start(element) | Event::Empty(element) => {
                    return Err(GeometryError::UnexpectedElement(
                        String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    ));
                }
                Event::Eof => {
                    return Err(GeometryError::UnexpectedElement("EOF".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Parses an `a:custGeom` after the caller consumed its start event.
    pub fn from_element(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
    ) -> Result<Self, GeometryError> {
        if !matches_local_name(start.name().as_ref(), b"custGeom") {
            return Err(GeometryError::UnexpectedElement(element_name(start)));
        }

        let mut adjust_values = None;
        let mut guides = None;
        let mut text_rectangle = None;
        let mut path_list = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0;
        let mut buffer = Vec::new();

        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| GeometryError::Xml(error.to_string()))?
            {
                Event::Start(element) => match local_name(element.name().as_ref()) {
                    b"avLst" if adjust_values.is_none() => {
                        adjust_values = Some(parse_guide_list(reader, &element, b"avLst")?);
                        boundary = boundary.max(1);
                    }
                    b"gdLst" if guides.is_none() => {
                        guides = Some(parse_guide_list(reader, &element, b"gdLst")?);
                        boundary = boundary.max(2);
                    }
                    b"rect" if text_rectangle.is_none() => {
                        text_rectangle = Some(parse_rect(reader, &element)?);
                        boundary = boundary.max(5);
                    }
                    b"pathLst" if path_list.is_none() => {
                        path_list = Some(parse_path_list(reader, &element)?);
                        boundary = boundary.max(6);
                    }
                    _ => raw_children.push(boundary, capture_element(reader, &element)?),
                },
                Event::Empty(element) => match local_name(element.name().as_ref()) {
                    b"avLst" if adjust_values.is_none() => {
                        adjust_values = Some(GuideList {
                            guides: Vec::new(),
                            guide_raw_children: Vec::new(),
                            raw_children: OrderedRawChildren::default(),
                        });
                        boundary = boundary.max(1);
                    }
                    b"gdLst" if guides.is_none() => {
                        guides = Some(GuideList {
                            guides: Vec::new(),
                            guide_raw_children: Vec::new(),
                            raw_children: OrderedRawChildren::default(),
                        });
                        boundary = boundary.max(2);
                    }
                    b"rect" if text_rectangle.is_none() => {
                        text_rectangle = Some(parse_empty_rect(&element)?);
                        boundary = boundary.max(5);
                    }
                    b"pathLst" if path_list.is_none() => {
                        path_list = Some(PathList {
                            paths: Vec::new(),
                            raw_children: OrderedRawChildren::default(),
                        });
                        boundary = boundary.max(6);
                    }
                    _ => raw_children.push(boundary, capture_empty_element(&element)?),
                },
                Event::End(element) if matches_local_name(element.name().as_ref(), b"custGeom") => {
                    break;
                }
                Event::Eof => {
                    return Err(GeometryError::Xml("missing closing a:custGeom".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }

        Ok(Self {
            adjust_values,
            guides,
            text_rectangle,
            path_list: path_list.ok_or(GeometryError::MissingPathList)?,
            raw_children,
        })
    }

    pub fn adjust_values(&self) -> &[Guide] {
        self.adjust_values
            .as_ref()
            .map_or(&[], |list| list.guides.as_slice())
    }

    pub fn guides(&self) -> &[Guide] {
        self.guides
            .as_ref()
            .map_or(&[], |list| list.guides.as_slice())
    }

    pub fn paths(&self) -> &[CT_Path2D] {
        &self.path_list.paths
    }

    /// Writes with the canonical `a:` prefix and DrawingML schema order.
    pub fn to_xml(&self) -> Result<Vec<u8>, GeometryError> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Start(BytesStart::new("a:custGeom")))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        emit_raw(&mut writer, self.raw_children.at(0))?;
        if let Some(list) = &self.adjust_values {
            write_guide_list(&mut writer, "a:avLst", list)?;
        }
        emit_raw(&mut writer, self.raw_children.at(1))?;
        if let Some(list) = &self.guides {
            write_guide_list(&mut writer, "a:gdLst", list)?;
        }
        emit_raw(&mut writer, self.raw_children.at(2))?;
        emit_raw(&mut writer, self.raw_children.at(3))?;
        emit_raw(&mut writer, self.raw_children.at(4))?;
        if let Some(rectangle) = &self.text_rectangle {
            write_rect(&mut writer, rectangle)?;
        }
        emit_raw(&mut writer, self.raw_children.at(5))?;
        write_path_list(&mut writer, &self.path_list)?;
        emit_raw(&mut writer, self.raw_children.at(6))?;
        writer
            .write_event(Event::End(BytesEnd::new("a:custGeom")))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        Ok(writer.into_inner())
    }

    /// Evaluates all paths and the text rectangle in each path's coordinate space.
    pub fn evaluate(
        &self,
        overrides: &BTreeMap<String, f64>,
    ) -> Result<EvaluatedCustomGeometry, GeometryError> {
        self.evaluate_with_default_dimensions(overrides, None, None)
    }

    /// Evaluates paths with the containing shape size on omitted coordinate axes.
    pub fn evaluate_with_size(
        &self,
        overrides: &BTreeMap<String, f64>,
        size: (f64, f64),
    ) -> Result<EvaluatedCustomGeometry, GeometryError> {
        let text_dimensions = self.paths().first().map_or(size, |path| {
            (path.width.unwrap_or(size.0), path.height.unwrap_or(size.1))
        });
        self.evaluate_with_default_dimensions(overrides, Some(size), Some(text_dimensions))
    }

    fn evaluate_with_default_dimensions(
        &self,
        overrides: &BTreeMap<String, f64>,
        default_dimensions: Option<(f64, f64)>,
        text_dimensions: Option<(f64, f64)>,
    ) -> Result<EvaluatedCustomGeometry, GeometryError> {
        let mut paths = Vec::with_capacity(self.path_list.paths.len());
        let mut text_rectangle = None;
        for (index, path) in self.path_list.paths.iter().enumerate() {
            let width = path
                .width
                .or(default_dimensions.map(|dimensions| dimensions.0))
                .ok_or(GeometryError::MissingPathDimensions)?;
            let height = path
                .height
                .or(default_dimensions.map(|dimensions| dimensions.1))
                .ok_or(GeometryError::MissingPathDimensions)?;
            let mut evaluator = GuideEvaluator::new(width, height)?;
            evaluator.apply_adjust_values(self.adjust_values(), overrides)?;
            evaluator.evaluate_guides(self.guides())?;
            let commands = path
                .commands
                .iter()
                .map(|record| record.command.to_evaluator_command())
                .collect::<Vec<_>>();
            paths.push(evaluator.evaluate_path(&commands)?);
            if index == 0 && default_dimensions.is_none() {
                text_rectangle = self
                    .text_rectangle
                    .as_ref()
                    .map(|rectangle| rectangle.evaluate(&evaluator))
                    .transpose()?;
            }
        }
        if let Some((width, height)) = text_dimensions {
            let mut evaluator = GuideEvaluator::new(width, height)?;
            evaluator.apply_adjust_values(self.adjust_values(), overrides)?;
            evaluator.evaluate_guides(self.guides())?;
            text_rectangle = self
                .text_rectangle
                .as_ref()
                .map(|rectangle| rectangle.evaluate(&evaluator))
                .transpose()?;
        }
        Ok(EvaluatedCustomGeometry {
            paths,
            text_rectangle,
        })
    }
}

impl CT_PresetGeometry2D {
    /// Creates preset geometry with a canonical empty adjustment list.
    pub fn new(preset: &str) -> Result<Self, GeometryError> {
        if preset_shape_definition(preset).is_none() {
            return Err(GeometryError::UnknownPreset(preset.to_owned()));
        }
        Ok(Self {
            preset: preset.to_owned(),
            adjust_values: Some(GuideList {
                guides: Vec::new(),
                guide_raw_children: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            }),
            raw_children: OrderedRawChildren::default(),
        })
    }

    /// Parses one complete `a:prstGeom` element with any namespace prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self, GeometryError> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| GeometryError::Xml(error.to_string()))?
            {
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"prstGeom") =>
                {
                    return Self::from_element(&mut reader, &element);
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"prstGeom") =>
                {
                    return Ok(Self {
                        preset: required_attr(&element, b"prst")?,
                        adjust_values: None,
                        raw_children: OrderedRawChildren::default(),
                    });
                }
                Event::Start(element) | Event::Empty(element) => {
                    return Err(GeometryError::UnexpectedElement(element_name(&element)));
                }
                Event::Eof => {
                    return Err(GeometryError::UnexpectedElement("EOF".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_element(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
    ) -> Result<Self, GeometryError> {
        let preset = required_attr(start, b"prst")?;
        let mut adjust_values = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| GeometryError::Xml(error.to_string()))?
            {
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"avLst")
                        && adjust_values.is_none() =>
                {
                    adjust_values = Some(parse_guide_list(reader, &element, b"avLst")?);
                    boundary = 1;
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"avLst")
                        && adjust_values.is_none() =>
                {
                    adjust_values = Some(GuideList {
                        guides: Vec::new(),
                        guide_raw_children: Vec::new(),
                        raw_children: OrderedRawChildren::default(),
                    });
                    boundary = 1;
                }
                Event::Start(element) => {
                    raw_children.push(boundary, capture_element(reader, &element)?)
                }
                Event::Empty(element) => {
                    raw_children.push(boundary, capture_empty_element(&element)?)
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), b"prstGeom") => {
                    break;
                }
                Event::Eof => {
                    return Err(GeometryError::Xml("missing closing a:prstGeom".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
        Ok(Self {
            preset,
            adjust_values,
            raw_children,
        })
    }

    pub fn adjust_values(&self) -> &[Guide] {
        self.adjust_values
            .as_ref()
            .map_or(&[], |list| list.guides.as_slice())
    }

    /// Inserts or replaces one named `val` adjustment guide.
    pub fn set_adjust_value(&mut self, name: &str, value: f64) -> Result<(), GeometryError> {
        if !value.is_finite() {
            return Err(GeometryError::NonFiniteValue(name.to_owned()));
        }
        let guide = Guide {
            name: name.to_owned(),
            op: GuideOp::Val,
            args: vec![GuideOperand::Literal(value)],
        };
        let list = self.adjust_values.get_or_insert_with(|| GuideList {
            guides: Vec::new(),
            guide_raw_children: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        });
        if let Some(index) = list.guides.iter().position(|guide| guide.name == name) {
            list.guides[index] = guide;
        } else {
            let trailing_boundary = list.guides.len();
            list.raw_children.shift_boundaries_from(trailing_boundary);
            list.guides.push(guide);
            list.guide_raw_children.push(OrderedRawChildren::default());
        }
        Ok(())
    }

    /// Returns the preset definition's own `avLst` guides in definition order.
    ///
    /// Only the `avLst` is read, so a preset whose other guides do not parse
    /// still reports its defaults. An unknown preset has none.
    pub fn default_adjust_values(&self) -> Result<Vec<Guide>, GeometryError> {
        let Some(xml) = preset_shape_definition(&self.preset) else {
            return Ok(Vec::new());
        };
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| GeometryError::Xml(error.to_string()))?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"avLst") => {
                    return Ok(parse_guide_list(&mut reader, &element, b"avLst")?.guides);
                }
                Event::Eof => return Ok(Vec::new()),
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Writes with the canonical `a:` prefix and DrawingML schema order.
    pub fn to_xml(&self) -> Result<Vec<u8>, GeometryError> {
        let mut writer = Writer::new(Vec::new());
        let mut start = BytesStart::new("a:prstGeom");
        start.push_attribute(("prst", self.preset.as_str()));
        if self.adjust_values.is_none() && self.raw_children.is_empty() {
            writer
                .write_event(Event::Empty(start))
                .map_err(|error| GeometryError::Xml(error.to_string()))?;
            return Ok(writer.into_inner());
        }
        writer
            .write_event(Event::Start(start))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        emit_raw(&mut writer, self.raw_children.at(0))?;
        if let Some(list) = &self.adjust_values {
            write_guide_list(&mut writer, "a:avLst", list)?;
        }
        emit_raw(&mut writer, self.raw_children.at(1))?;
        writer
            .write_event(Event::End(BytesEnd::new("a:prstGeom")))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        Ok(writer.into_inner())
    }

    /// Evaluates a known preset in the shape's coordinate space.
    pub fn evaluate(
        &self,
        size: (f64, f64),
    ) -> Result<Option<EvaluatedCustomGeometry>, GeometryError> {
        let Some(xml) = preset_shape_definition(&self.preset) else {
            return Ok(None);
        };
        let definition = CT_CustomGeometry2D::from_xml(xml)?;
        let mut override_evaluator = GuideEvaluator::new(size.0, size.1)?;
        override_evaluator.apply_adjust_values(self.adjust_values(), &BTreeMap::new())?;
        let overrides = self
            .adjust_values()
            .iter()
            .map(|guide| Ok((guide.name.clone(), override_evaluator.value(&guide.name)?)))
            .collect::<Result<BTreeMap<_, _>, GeometryError>>()?;
        let mut evaluated =
            definition.evaluate_with_default_dimensions(&overrides, Some(size), Some(size))?;
        for (commands, path) in evaluated.paths.iter_mut().zip(definition.paths()) {
            let scale_x = path.width.map_or(1.0, |width| size.0 / width);
            let scale_y = path.height.map_or(1.0, |height| size.1 / height);
            for command in commands {
                scale_evaluated_path_command(command, scale_x, scale_y);
            }
        }
        Ok(Some(evaluated))
    }
}

fn scale_evaluated_path_command(command: &mut EvaluatedPathCommand, scale_x: f64, scale_y: f64) {
    match command {
        EvaluatedPathCommand::MoveTo { x, y } | EvaluatedPathCommand::LineTo { x, y } => {
            *x *= scale_x;
            *y *= scale_y;
        }
        EvaluatedPathCommand::CubicTo {
            x1,
            y1,
            x2,
            y2,
            x,
            y,
        } => {
            *x1 *= scale_x;
            *y1 *= scale_y;
            *x2 *= scale_x;
            *y2 *= scale_y;
            *x *= scale_x;
            *y *= scale_y;
        }
        EvaluatedPathCommand::Close => {}
    }
}

impl CT_GeomRect {
    fn evaluate(
        &self,
        evaluator: &GuideEvaluator,
    ) -> Result<EvaluatedTextRectangle, GeometryError> {
        Ok(EvaluatedTextRectangle {
            left: evaluator.resolve(&self.left)?,
            top: evaluator.resolve(&self.top)?,
            right: evaluator.resolve(&self.right)?,
            bottom: evaluator.resolve(&self.bottom)?,
        })
    }
}

impl CT_Path2DCommand {
    fn to_evaluator_command(&self) -> PathCommand {
        match self {
            Self::MoveTo(point) => PathCommand::MoveTo {
                x: point.x.clone(),
                y: point.y.clone(),
            },
            Self::LineTo(point) => PathCommand::LineTo {
                x: point.x.clone(),
                y: point.y.clone(),
            },
            Self::CubicTo {
                control_1,
                control_2,
                end,
            } => PathCommand::CubicTo {
                x1: control_1.x.clone(),
                y1: control_1.y.clone(),
                x2: control_2.x.clone(),
                y2: control_2.y.clone(),
                x: end.x.clone(),
                y: end.y.clone(),
            },
            Self::ArcTo {
                width_radius,
                height_radius,
                start_angle,
                sweep_angle,
            } => PathCommand::ArcTo {
                width_radius: width_radius.clone(),
                height_radius: height_radius.clone(),
                start_angle: start_angle.clone(),
                sweep_angle: sweep_angle.clone(),
            },
            Self::Close => PathCommand::Close,
        }
    }
}

fn parse_guide_list(
    reader: &mut Reader<&[u8]>,
    _start: &BytesStart<'_>,
    end_name: &[u8],
) -> Result<GuideList, GeometryError> {
    let mut guides = Vec::new();
    let mut guide_raw_children = Vec::new();
    let mut raw_children = OrderedRawChildren::default();
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| GeometryError::Xml(error.to_string()))?
        {
            Event::Empty(element) if matches_local_name(element.name().as_ref(), b"gd") => {
                guides.push(parse_guide(&element)?);
                guide_raw_children.push(OrderedRawChildren::default());
            }
            Event::Start(element) if matches_local_name(element.name().as_ref(), b"gd") => {
                let guide = parse_guide(&element)?;
                let mut children = OrderedRawChildren::default();
                consume_leaf_children(reader, b"gd", &mut children, 0)?;
                guides.push(guide);
                guide_raw_children.push(children);
            }
            Event::Start(element) => {
                raw_children.push(guides.len(), capture_element(reader, &element)?)
            }
            Event::Empty(element) => {
                raw_children.push(guides.len(), capture_empty_element(&element)?)
            }
            Event::End(element) if matches_local_name(element.name().as_ref(), end_name) => break,
            Event::Eof => {
                return Err(GeometryError::Xml(format!(
                    "missing closing a:{}",
                    String::from_utf8_lossy(end_name)
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(GuideList {
        guides,
        guide_raw_children,
        raw_children,
    })
}

fn consume_leaf_children(
    reader: &mut Reader<&[u8]>,
    end_name: &[u8],
    raw_children: &mut OrderedRawChildren,
    boundary: usize,
) -> Result<(), GeometryError> {
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| GeometryError::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                raw_children.push(boundary, capture_element(reader, &element)?)
            }
            Event::Empty(element) => raw_children.push(boundary, capture_empty_element(&element)?),
            Event::End(element) if matches_local_name(element.name().as_ref(), end_name) => break,
            Event::Eof => {
                return Err(GeometryError::Xml(format!(
                    "missing closing a:{}",
                    String::from_utf8_lossy(end_name)
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn parse_guide(element: &BytesStart<'_>) -> Result<Guide, GeometryError> {
    Guide::parse(
        required_attr(element, b"name")?,
        &required_attr(element, b"fmla")?,
    )
}

fn parse_empty_rect(element: &BytesStart<'_>) -> Result<CT_GeomRect, GeometryError> {
    Ok(CT_GeomRect {
        left: required_operand(element, b"l")?,
        top: required_operand(element, b"t")?,
        right: required_operand(element, b"r")?,
        bottom: required_operand(element, b"b")?,
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_rect(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<CT_GeomRect, GeometryError> {
    let mut rectangle = parse_empty_rect(element)?;
    consume_leaf_children(reader, b"rect", &mut rectangle.raw_children, 0)?;
    Ok(rectangle)
}

fn parse_path_list(
    reader: &mut Reader<&[u8]>,
    _start: &BytesStart<'_>,
) -> Result<PathList, GeometryError> {
    let mut paths = Vec::new();
    let mut raw_children = OrderedRawChildren::default();
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| GeometryError::Xml(error.to_string()))?
        {
            Event::Start(element) if matches_local_name(element.name().as_ref(), b"path") => {
                paths.push(parse_path(reader, &element)?);
            }
            Event::Empty(element) if matches_local_name(element.name().as_ref(), b"path") => {
                paths.push(parse_empty_path(&element)?);
            }
            Event::Start(element) => {
                raw_children.push(paths.len(), capture_element(reader, &element)?)
            }
            Event::Empty(element) => {
                raw_children.push(paths.len(), capture_empty_element(&element)?)
            }
            Event::End(element) if matches_local_name(element.name().as_ref(), b"pathLst") => {
                break;
            }
            Event::Eof => {
                return Err(GeometryError::Xml("missing closing a:pathLst".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(PathList {
        paths,
        raw_children,
    })
}

fn parse_empty_path(element: &BytesStart<'_>) -> Result<CT_Path2D, GeometryError> {
    Ok(CT_Path2D {
        width: optional_f64(element, b"w")?,
        height: optional_f64(element, b"h")?,
        fill: get_attr(element, b"fill"),
        stroke: optional_bool(element, b"stroke")?,
        extrusion_ok: optional_bool(element, b"extrusionOk")?,
        commands: Vec::new(),
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_path(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<CT_Path2D, GeometryError> {
    let mut path = parse_empty_path(element)?;
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| GeometryError::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                if let Some(command) = parse_path_command(reader, &element)? {
                    path.commands.push(command);
                } else {
                    path.raw_children
                        .push(path.commands.len(), capture_element(reader, &element)?);
                }
            }
            Event::Empty(element) => {
                if let Some(command) = parse_empty_path_command(&element)? {
                    path.commands.push(command);
                } else {
                    path.raw_children
                        .push(path.commands.len(), capture_empty_element(&element)?);
                }
            }
            Event::End(element) if matches_local_name(element.name().as_ref(), b"path") => break,
            Event::Eof => {
                return Err(GeometryError::Xml("missing closing a:path".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(path)
}

fn parse_path_command(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<PathCommandRecord>, GeometryError> {
    match local_name(element.name().as_ref()) {
        b"moveTo" => parse_point_command(reader, b"moveTo", 1, |mut points| {
            CT_Path2DCommand::MoveTo(points.remove(0))
        })
        .map(Some),
        b"lnTo" => parse_point_command(reader, b"lnTo", 1, |mut points| {
            CT_Path2DCommand::LineTo(points.remove(0))
        })
        .map(Some),
        b"cubicBezTo" => parse_point_command(reader, b"cubicBezTo", 3, |mut points| {
            CT_Path2DCommand::CubicTo {
                control_1: points.remove(0),
                control_2: points.remove(0),
                end: points.remove(0),
            }
        })
        .map(Some),
        b"arcTo" => {
            let command = parse_arc(element)?;
            let mut raw_children = OrderedRawChildren::default();
            consume_leaf_children(reader, b"arcTo", &mut raw_children, 0)?;
            Ok(Some(PathCommandRecord {
                command,
                raw_children,
            }))
        }
        b"close" => {
            let mut raw_children = OrderedRawChildren::default();
            consume_leaf_children(reader, b"close", &mut raw_children, 0)?;
            Ok(Some(PathCommandRecord {
                command: CT_Path2DCommand::Close,
                raw_children,
            }))
        }
        _ => Ok(None),
    }
}

fn parse_empty_path_command(
    element: &BytesStart<'_>,
) -> Result<Option<PathCommandRecord>, GeometryError> {
    match local_name(element.name().as_ref()) {
        b"arcTo" => Ok(Some(PathCommandRecord {
            command: parse_arc(element)?,
            raw_children: OrderedRawChildren::default(),
        })),
        b"close" => Ok(Some(PathCommandRecord {
            command: CT_Path2DCommand::Close,
            raw_children: OrderedRawChildren::default(),
        })),
        b"moveTo" | b"lnTo" | b"cubicBezTo" => Err(GeometryError::Xml(format!(
            "DrawingML {} requires point children",
            element_name(element)
        ))),
        _ => Ok(None),
    }
}

fn parse_point_command(
    reader: &mut Reader<&[u8]>,
    end_name: &[u8],
    expected_points: usize,
    make_command: impl FnOnce(Vec<CT_AdjPoint2D>) -> CT_Path2DCommand,
) -> Result<PathCommandRecord, GeometryError> {
    let mut points = Vec::new();
    let mut raw_children = OrderedRawChildren::default();
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| GeometryError::Xml(error.to_string()))?
        {
            Event::Empty(element)
                if matches_local_name(element.name().as_ref(), b"pt")
                    && points.len() < expected_points =>
            {
                points.push(parse_point(&element)?);
            }
            Event::Start(element)
                if matches_local_name(element.name().as_ref(), b"pt")
                    && points.len() < expected_points =>
            {
                let mut point = parse_point(&element)?;
                consume_leaf_children(reader, b"pt", &mut point.raw_children, 0)?;
                points.push(point);
            }
            Event::Start(element) => {
                raw_children.push(points.len(), capture_element(reader, &element)?)
            }
            Event::Empty(element) => {
                raw_children.push(points.len(), capture_empty_element(&element)?)
            }
            Event::End(element) if matches_local_name(element.name().as_ref(), end_name) => break,
            Event::Eof => {
                return Err(GeometryError::Xml(format!(
                    "missing closing a:{}",
                    String::from_utf8_lossy(end_name)
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    if points.len() != expected_points {
        return Err(GeometryError::Xml(format!(
            "DrawingML {} requires {expected_points} point children",
            String::from_utf8_lossy(end_name)
        )));
    }
    Ok(PathCommandRecord {
        command: make_command(points),
        raw_children,
    })
}

fn parse_point(element: &BytesStart<'_>) -> Result<CT_AdjPoint2D, GeometryError> {
    Ok(CT_AdjPoint2D {
        x: required_operand(element, b"x")?,
        y: required_operand(element, b"y")?,
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_arc(element: &BytesStart<'_>) -> Result<CT_Path2DCommand, GeometryError> {
    Ok(CT_Path2DCommand::ArcTo {
        width_radius: required_operand(element, b"wR")?,
        height_radius: required_operand(element, b"hR")?,
        start_angle: required_operand(element, b"stAng")?,
        sweep_angle: required_operand(element, b"swAng")?,
    })
}

fn write_guide_list<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    list: &GuideList,
) -> Result<(), GeometryError> {
    if list.guides.is_empty() && list.raw_children.is_empty() {
        writer
            .write_event(Event::Empty(BytesStart::new(tag)))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        return Ok(());
    }
    writer
        .write_event(Event::Start(BytesStart::new(tag)))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    for (index, guide) in list.guides.iter().enumerate() {
        emit_raw(writer, list.raw_children.at(index))?;
        let formula = guide.formula();
        let mut element = BytesStart::new("a:gd");
        element.push_attribute(("name", guide.name.as_str()));
        element.push_attribute(("fmla", formula.as_str()));
        let children = &list.guide_raw_children[index];
        write_leaf_with_raw(writer, element, "a:gd", children)?;
    }
    emit_raw(writer, list.raw_children.at(list.guides.len()))?;
    writer
        .write_event(Event::End(BytesEnd::new(tag)))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    Ok(())
}

fn write_rect<W: Write>(
    writer: &mut Writer<W>,
    rectangle: &CT_GeomRect,
) -> Result<(), GeometryError> {
    let values = [
        operand_text(&rectangle.left),
        operand_text(&rectangle.top),
        operand_text(&rectangle.right),
        operand_text(&rectangle.bottom),
    ];
    let mut element = BytesStart::new("a:rect");
    element.push_attribute(("l", values[0].as_str()));
    element.push_attribute(("t", values[1].as_str()));
    element.push_attribute(("r", values[2].as_str()));
    element.push_attribute(("b", values[3].as_str()));
    if rectangle.raw_children.is_empty() {
        writer
            .write_event(Event::Empty(element))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
    } else {
        writer
            .write_event(Event::Start(element))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        emit_raw(writer, rectangle.raw_children.at(0))?;
        writer
            .write_event(Event::End(BytesEnd::new("a:rect")))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
    }
    Ok(())
}

fn write_path_list<W: Write>(writer: &mut Writer<W>, list: &PathList) -> Result<(), GeometryError> {
    writer
        .write_event(Event::Start(BytesStart::new("a:pathLst")))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    for (index, path) in list.paths.iter().enumerate() {
        emit_raw(writer, list.raw_children.at(index))?;
        write_path(writer, path)?;
    }
    emit_raw(writer, list.raw_children.at(list.paths.len()))?;
    writer
        .write_event(Event::End(BytesEnd::new("a:pathLst")))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    Ok(())
}

fn write_path<W: Write>(writer: &mut Writer<W>, path: &CT_Path2D) -> Result<(), GeometryError> {
    let width = path.width.map(|value| value.to_string());
    let height = path.height.map(|value| value.to_string());
    let mut element = BytesStart::new("a:path");
    if let Some(value) = width.as_deref() {
        element.push_attribute(("w", value));
    }
    if let Some(value) = height.as_deref() {
        element.push_attribute(("h", value));
    }
    if let Some(value) = path.fill.as_deref() {
        element.push_attribute(("fill", value));
    }
    if let Some(value) = path.stroke {
        element.push_attribute(("stroke", if value { "1" } else { "0" }));
    }
    if let Some(value) = path.extrusion_ok {
        element.push_attribute(("extrusionOk", if value { "1" } else { "0" }));
    }
    if path.commands.is_empty() && path.raw_children.is_empty() {
        writer
            .write_event(Event::Empty(element))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        return Ok(());
    }
    writer
        .write_event(Event::Start(element))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    for (index, command) in path.commands.iter().enumerate() {
        emit_raw(writer, path.raw_children.at(index))?;
        write_path_command(writer, command)?;
    }
    emit_raw(writer, path.raw_children.at(path.commands.len()))?;
    writer
        .write_event(Event::End(BytesEnd::new("a:path")))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    Ok(())
}

fn write_path_command<W: Write>(
    writer: &mut Writer<W>,
    record: &PathCommandRecord,
) -> Result<(), GeometryError> {
    match &record.command {
        CT_Path2DCommand::MoveTo(point) => {
            write_point_command(writer, "a:moveTo", std::slice::from_ref(point), record)?
        }
        CT_Path2DCommand::LineTo(point) => {
            write_point_command(writer, "a:lnTo", std::slice::from_ref(point), record)?
        }
        CT_Path2DCommand::CubicTo {
            control_1,
            control_2,
            end,
        } => write_point_command(
            writer,
            "a:cubicBezTo",
            &[control_1.clone(), control_2.clone(), end.clone()],
            record,
        )?,
        CT_Path2DCommand::ArcTo {
            width_radius,
            height_radius,
            start_angle,
            sweep_angle,
        } => {
            let values = [
                operand_text(width_radius),
                operand_text(height_radius),
                operand_text(start_angle),
                operand_text(sweep_angle),
            ];
            let mut element = BytesStart::new("a:arcTo");
            element.push_attribute(("wR", values[0].as_str()));
            element.push_attribute(("hR", values[1].as_str()));
            element.push_attribute(("stAng", values[2].as_str()));
            element.push_attribute(("swAng", values[3].as_str()));
            write_leaf_with_raw(writer, element, "a:arcTo", &record.raw_children)?;
        }
        CT_Path2DCommand::Close => {
            write_leaf_with_raw(
                writer,
                BytesStart::new("a:close"),
                "a:close",
                &record.raw_children,
            )?;
        }
    }
    Ok(())
}

fn write_point_command<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    points: &[CT_AdjPoint2D],
    record: &PathCommandRecord,
) -> Result<(), GeometryError> {
    writer
        .write_event(Event::Start(BytesStart::new(tag)))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    for (index, point) in points.iter().enumerate() {
        emit_raw(writer, record.raw_children.at(index))?;
        write_point(writer, point)?;
    }
    emit_raw(writer, record.raw_children.at(points.len()))?;
    writer
        .write_event(Event::End(BytesEnd::new(tag)))
        .map_err(|error| GeometryError::Xml(error.to_string()))?;
    Ok(())
}

fn write_point<W: Write>(
    writer: &mut Writer<W>,
    point: &CT_AdjPoint2D,
) -> Result<(), GeometryError> {
    let x = operand_text(&point.x);
    let y = operand_text(&point.y);
    let mut element = BytesStart::new("a:pt");
    element.push_attribute(("x", x.as_str()));
    element.push_attribute(("y", y.as_str()));
    write_leaf_with_raw(writer, element, "a:pt", &point.raw_children)?;
    Ok(())
}

fn write_leaf_with_raw<W: Write>(
    writer: &mut Writer<W>,
    element: BytesStart<'_>,
    tag: &str,
    raw_children: &OrderedRawChildren,
) -> Result<(), GeometryError> {
    if raw_children.is_empty() {
        writer
            .write_event(Event::Empty(element))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
    } else {
        writer
            .write_event(Event::Start(element))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
        emit_raw(writer, raw_children.at(0))?;
        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
    }
    Ok(())
}

fn emit_raw<'a, W: Write>(
    writer: &mut Writer<W>,
    children: impl Iterator<Item = &'a [u8]>,
) -> Result<(), GeometryError> {
    for child in children {
        writer
            .get_mut()
            .write_all(child)
            .map_err(|error| GeometryError::Xml(error.to_string()))?;
    }
    Ok(())
}

fn required_operand(
    element: &BytesStart<'_>,
    attribute: &[u8],
) -> Result<GuideOperand, GeometryError> {
    GuideOperand::parse(&required_attr(element, attribute)?)
}

fn required_attr(element: &BytesStart<'_>, attribute: &[u8]) -> Result<String, GeometryError> {
    get_attr(element, attribute).ok_or_else(|| GeometryError::MissingAttribute {
        element: element_name(element),
        attribute: String::from_utf8_lossy(attribute).into_owned(),
    })
}

fn optional_f64(element: &BytesStart<'_>, attribute: &[u8]) -> Result<Option<f64>, GeometryError> {
    get_attr(element, attribute)
        .map(|value| {
            value
                .parse::<f64>()
                .ok()
                .filter(|value| value.is_finite() && *value >= 0.0)
                .ok_or_else(|| invalid_attribute(element, attribute, value))
        })
        .transpose()
}

fn optional_bool(
    element: &BytesStart<'_>,
    attribute: &[u8],
) -> Result<Option<bool>, GeometryError> {
    get_attr(element, attribute)
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid_attribute(element, attribute, value)),
        })
        .transpose()
}

fn invalid_attribute(element: &BytesStart<'_>, attribute: &[u8], value: String) -> GeometryError {
    GeometryError::InvalidAttribute {
        element: element_name(element),
        attribute: String::from_utf8_lossy(attribute).into_owned(),
        value,
    }
}

fn element_name(element: &BytesStart<'_>) -> String {
    String::from_utf8_lossy(local_name(element.name().as_ref())).into_owned()
}

fn operand_text(operand: &GuideOperand) -> String {
    match operand {
        GuideOperand::Literal(value) => value.to_string(),
        GuideOperand::Guide(name) => name.clone(),
    }
}

#[derive(Clone, Debug)]
pub struct GuideEvaluator {
    values: BTreeMap<String, f64>,
}

impl GuideEvaluator {
    pub fn new(width: f64, height: f64) -> Result<Self, GeometryError> {
        ensure_finite(width, "shape width")?;
        ensure_finite(height, "shape height")?;

        let mut evaluator = Self {
            values: BTreeMap::new(),
        };
        evaluator.seed("w", width);
        evaluator.seed("h", height);
        evaluator.seed("l", 0.0);
        evaluator.seed("t", 0.0);
        evaluator.seed("r", width);
        evaluator.seed("b", height);
        evaluator.seed("hc", width / 2.0);
        evaluator.seed("vc", height / 2.0);
        evaluator.seed("ss", width.min(height));
        evaluator.seed("ls", width.max(height));

        for divisor in [2_u32, 3, 4, 5, 6, 8, 10, 12, 32] {
            evaluator.seed(&format!("wd{divisor}"), width / f64::from(divisor));
        }
        for divisor in [2_u32, 3, 4, 5, 6, 8, 10] {
            evaluator.seed(&format!("hd{divisor}"), height / f64::from(divisor));
        }
        for divisor in [2_u32, 4, 6, 8, 16, 32] {
            evaluator.seed(
                &format!("ssd{divisor}"),
                width.min(height) / f64::from(divisor),
            );
        }
        evaluator.seed("cd2", 180.0 * ANGLE_UNITS_PER_DEGREE);
        evaluator.seed("cd4", 90.0 * ANGLE_UNITS_PER_DEGREE);
        evaluator.seed("cd8", 45.0 * ANGLE_UNITS_PER_DEGREE);
        evaluator.seed("3cd4", 270.0 * ANGLE_UNITS_PER_DEGREE);
        evaluator.seed("3cd8", 135.0 * ANGLE_UNITS_PER_DEGREE);
        evaluator.seed("5cd8", 225.0 * ANGLE_UNITS_PER_DEGREE);
        evaluator.seed("7cd8", 315.0 * ANGLE_UNITS_PER_DEGREE);
        Ok(evaluator)
    }

    pub fn value(&self, name: &str) -> Result<f64, GeometryError> {
        self.values
            .get(name)
            .copied()
            .ok_or_else(|| GeometryError::UnknownGuide(name.to_owned()))
    }

    pub fn apply_adjust_values(
        &mut self,
        adjustments: &[Guide],
        overrides: &BTreeMap<String, f64>,
    ) -> Result<(), GeometryError> {
        let declared = adjustments
            .iter()
            .map(|guide| guide.name.as_str())
            .collect::<BTreeSet<_>>();
        if let Some(name) = overrides
            .keys()
            .find(|name| !declared.contains(name.as_str()))
        {
            return Err(GeometryError::UnknownAdjustOverride(name.clone()));
        }

        for adjustment in adjustments {
            let value = match overrides.get(&adjustment.name) {
                Some(value) => {
                    ensure_finite(*value, &adjustment.name)?;
                    *value
                }
                None => self.evaluate_operation(adjustment.op, &adjustment.args)?,
            };
            self.insert_named(&adjustment.name, value)?;
        }
        Ok(())
    }

    pub fn evaluate_guides(&mut self, guides: &[Guide]) -> Result<(), GeometryError> {
        let mut evaluated = BTreeSet::new();
        for guide in guides {
            let value = self.evaluate_operation(guide.op, &guide.args)?;
            if evaluated.insert(guide.name.clone()) {
                self.insert_named(&guide.name, value)?;
            } else {
                ensure_finite(value, &guide.name)?;
                self.values.insert(guide.name.clone(), value);
            }
        }
        Ok(())
    }

    pub fn evaluate_path(
        &self,
        commands: &[PathCommand],
    ) -> Result<Vec<EvaluatedPathCommand>, GeometryError> {
        let mut output = Vec::with_capacity(commands.len());
        let mut current = None;
        let mut subpath_start = None;

        for command in commands {
            match command {
                PathCommand::MoveTo { x, y } => {
                    let point = (self.resolve(x)?, self.resolve(y)?);
                    output.push(EvaluatedPathCommand::MoveTo {
                        x: point.0,
                        y: point.1,
                    });
                    current = Some(point);
                    subpath_start = Some(point);
                }
                PathCommand::LineTo { x, y } => {
                    current.ok_or(GeometryError::PathHasNoCurrentPoint)?;
                    let point = (self.resolve(x)?, self.resolve(y)?);
                    output.push(EvaluatedPathCommand::LineTo {
                        x: point.0,
                        y: point.1,
                    });
                    current = Some(point);
                }
                PathCommand::CubicTo {
                    x1,
                    y1,
                    x2,
                    y2,
                    x,
                    y,
                } => {
                    current.ok_or(GeometryError::PathHasNoCurrentPoint)?;
                    let command = EvaluatedPathCommand::CubicTo {
                        x1: self.resolve(x1)?,
                        y1: self.resolve(y1)?,
                        x2: self.resolve(x2)?,
                        y2: self.resolve(y2)?,
                        x: self.resolve(x)?,
                        y: self.resolve(y)?,
                    };
                    if let EvaluatedPathCommand::CubicTo { x, y, .. } = command {
                        current = Some((x, y));
                    }
                    output.push(command);
                }
                PathCommand::ArcTo {
                    width_radius,
                    height_radius,
                    start_angle,
                    sweep_angle,
                } => {
                    let start = current.ok_or(GeometryError::PathHasNoCurrentPoint)?;
                    let cubics = flatten_arc(
                        start,
                        self.resolve(width_radius)?,
                        self.resolve(height_radius)?,
                        self.resolve(start_angle)?,
                        self.resolve(sweep_angle)?,
                    )?;
                    if let Some(EvaluatedPathCommand::CubicTo { x, y, .. }) = cubics.last() {
                        current = Some((*x, *y));
                    }
                    output.extend(cubics);
                }
                PathCommand::Close => {
                    current.ok_or(GeometryError::PathHasNoCurrentPoint)?;
                    output.push(EvaluatedPathCommand::Close);
                    current = subpath_start;
                }
            }
        }
        Ok(output)
    }

    fn seed(&mut self, name: &str, value: f64) {
        self.values.insert(name.to_owned(), value);
    }

    fn insert_named(&mut self, name: &str, value: f64) -> Result<(), GeometryError> {
        ensure_finite(value, name)?;
        if self.values.contains_key(name) {
            return Err(GeometryError::DuplicateGuide(name.to_owned()));
        }
        self.values.insert(name.to_owned(), value);
        Ok(())
    }

    fn resolve(&self, operand: &GuideOperand) -> Result<f64, GeometryError> {
        match operand {
            GuideOperand::Literal(value) => {
                ensure_finite(*value, "literal")?;
                Ok(*value)
            }
            GuideOperand::Guide(name) => self.value(name),
        }
    }

    fn evaluate_operation(
        &self,
        op: GuideOp,
        operands: &[GuideOperand],
    ) -> Result<f64, GeometryError> {
        let expected = op.argument_count();
        if operands.len() != expected {
            return Err(GeometryError::WrongArgumentCount {
                operation: format!("{op:?}"),
                expected,
                actual: operands.len(),
            });
        }
        let args = operands
            .iter()
            .map(|operand| self.resolve(operand))
            .collect::<Result<Vec<_>, _>>()?;

        let value = match op {
            GuideOp::MulDiv => checked_div(args[0] * args[1], args[2])?,
            GuideOp::AddSub => args[0] + args[1] - args[2],
            GuideOp::AddDiv => checked_div(args[0] + args[1], args[2])?,
            GuideOp::IfElse => {
                if args[0] > 0.0 {
                    args[1]
                } else {
                    args[2]
                }
            }
            GuideOp::Abs => args[0].abs(),
            GuideOp::At2 => radians_to_angle(args[1].atan2(args[0])),
            GuideOp::Cat2 => args[0] * args[2].atan2(args[1]).cos(),
            GuideOp::Cos => args[0] * angle_to_radians(args[1]).cos(),
            GuideOp::Max => args[0].max(args[1]),
            GuideOp::Min => args[0].min(args[1]),
            GuideOp::Mod => args[0].hypot(args[1]).hypot(args[2]),
            GuideOp::Pin => {
                if args[1] < args[0] {
                    args[0]
                } else if args[1] > args[2] {
                    args[2]
                } else {
                    args[1]
                }
            }
            GuideOp::Sat2 => args[0] * args[2].atan2(args[1]).sin(),
            GuideOp::Sin => args[0] * angle_to_radians(args[1]).sin(),
            GuideOp::Sqrt => args[0].abs().sqrt(),
            GuideOp::Tan => args[0] * angle_to_radians(args[1]).tan(),
            GuideOp::Val => args[0],
        };
        ensure_finite(value, "guide result")?;
        Ok(value)
    }
}

fn checked_div(numerator: f64, denominator: f64) -> Result<f64, GeometryError> {
    if denominator == 0.0 {
        return Err(GeometryError::DivisionByZero);
    }
    Ok(numerator / denominator)
}

fn angle_to_radians(angle: f64) -> f64 {
    angle / ANGLE_UNITS_PER_DEGREE * PI / 180.0
}

fn radians_to_angle(radians: f64) -> f64 {
    radians * 180.0 / PI * ANGLE_UNITS_PER_DEGREE
}

fn ensure_finite(value: f64, context: &str) -> Result<(), GeometryError> {
    if value.is_finite() {
        Ok(())
    } else {
        Err(GeometryError::NonFiniteValue(context.to_owned()))
    }
}

fn flatten_arc(
    current: (f64, f64),
    width_radius: f64,
    height_radius: f64,
    start_angle: f64,
    sweep_angle: f64,
) -> Result<Vec<EvaluatedPathCommand>, GeometryError> {
    for (value, context) in [
        (current.0, "arc start x"),
        (current.1, "arc start y"),
        (width_radius, "arc width radius"),
        (height_radius, "arc height radius"),
        (start_angle, "arc start angle"),
        (sweep_angle, "arc sweep angle"),
    ] {
        ensure_finite(value, context)?;
    }
    if width_radius <= 0.0 || height_radius <= 0.0 {
        return Err(GeometryError::InvalidArcRadius);
    }
    if sweep_angle == 0.0 {
        return Ok(Vec::new());
    }

    let segment_count = (sweep_angle.abs() / QUARTER_CIRCLE).ceil();
    if segment_count > MAX_ARC_SEGMENTS as f64 {
        return Err(GeometryError::ArcSweepTooLarge);
    }
    let segment_count = segment_count as usize;
    let segment_sweep = sweep_angle / segment_count as f64;
    let start_radians = angle_to_radians(start_angle);
    let center = (
        current.0 - width_radius * start_radians.cos(),
        current.1 - height_radius * start_radians.sin(),
    );
    let mut cubics = Vec::with_capacity(segment_count);

    for index in 0..segment_count {
        let angle_1 = angle_to_radians(start_angle + segment_sweep * index as f64);
        let angle_2 = angle_to_radians(start_angle + segment_sweep * (index + 1) as f64);
        let alpha = 4.0 / 3.0 * ((angle_2 - angle_1) / 4.0).tan();
        let point_1 = (
            center.0 + width_radius * angle_1.cos(),
            center.1 + height_radius * angle_1.sin(),
        );
        let point_2 = (
            center.0 + width_radius * angle_2.cos(),
            center.1 + height_radius * angle_2.sin(),
        );
        let tangent_1 = (-width_radius * angle_1.sin(), height_radius * angle_1.cos());
        let tangent_2 = (-width_radius * angle_2.sin(), height_radius * angle_2.cos());
        let command = EvaluatedPathCommand::CubicTo {
            x1: point_1.0 + alpha * tangent_1.0,
            y1: point_1.1 + alpha * tangent_1.1,
            x2: point_2.0 - alpha * tangent_2.0,
            y2: point_2.1 - alpha * tangent_2.1,
            x: point_2.0,
            y: point_2.1,
        };
        if let EvaluatedPathCommand::CubicTo {
            x1,
            y1,
            x2,
            y2,
            x,
            y,
        } = command
        {
            for value in [x1, y1, x2, y2, x, y] {
                ensure_finite(value, "arc cubic coordinate")?;
            }
        }
        cubics.push(command);
    }
    Ok(cubics)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preset_constructor_accepts_generated_names_and_rejects_unknown_names() {
        assert_eq!(
            CT_PresetGeometry2D::new("triangle").unwrap().preset,
            "triangle"
        );
        assert_eq!(
            CT_PresetGeometry2D::new("not-a-preset").unwrap_err(),
            GeometryError::UnknownPreset("not-a-preset".to_owned())
        );
    }

    #[test]
    fn preset_adjustment_setter_inserts_and_replaces_named_values() {
        let mut geometry = CT_PresetGeometry2D::from_xml(
            br#"<x:prstGeom xmlns:x="urn:a" prst="roundRect"><x:avLst><x:gd name="adj" fmla="val 12000"><ext:raw xmlns:ext="urn:ext"/></x:gd><ext:tail xmlns:ext="urn:ext"/></x:avLst><ext:after xmlns:ext="urn:ext"/></x:prstGeom>"#,
        )
        .unwrap();

        geometry.set_adjust_value("adj", 25_000.0).unwrap();
        geometry.set_adjust_value("adj2", 7_500.5).unwrap();
        let xml = String::from_utf8(geometry.to_xml().unwrap()).unwrap();

        assert_eq!(geometry.adjust_values().len(), 2);
        assert_eq!(geometry.adjust_values()[0].name, "adj");
        assert_eq!(geometry.adjust_values()[0].args, vec![literal(25_000.0)]);
        assert_eq!(geometry.adjust_values()[1].name, "adj2");
        assert_eq!(geometry.adjust_values()[1].args, vec![literal(7_500.5)]);
        assert_eq!(xml.matches("name=\"adj\"").count(), 1);
        assert_eq!(xml.matches("name=\"adj2\"").count(), 1);
        assert!(xml.contains("<ext:raw xmlns:ext=\"urn:ext\"/>"));
        assert!(xml.contains("<a:gd name=\"adj2\" fmla=\"val 7500.5\"/><ext:tail"));
        assert!(geometry.set_adjust_value("bad", f64::INFINITY).is_err());
    }

    fn literal(value: f64) -> GuideOperand {
        GuideOperand::Literal(value)
    }

    fn guide(name: &str) -> GuideOperand {
        GuideOperand::Guide(name.to_owned())
    }

    fn assert_close(actual: f64, expected: f64) {
        assert!((actual - expected).abs() < 1.0e-9, "{actual} != {expected}");
    }

    #[test]
    fn hand_written_custom_geometry_guides_produce_expected_path_coordinates() {
        let adjustments = [Guide::parse("adj1", "val 25000").unwrap()];
        let guides = [
            Guide::parse("x1", "*/ w adj1 100000").unwrap(),
            Guide::parse("y1", "+/ hd2 0 2").unwrap(),
            Guide::parse("x2", "+- r 0 x1").unwrap(),
            Guide::parse("y2", "?: adj1 75 25").unwrap(),
        ];
        let overrides = BTreeMap::from([("adj1".to_owned(), 20_000.0)]);
        let mut evaluator = GuideEvaluator::new(100.0, 100.0).unwrap();
        assert_eq!(evaluator.value("wd32").unwrap(), 3.125);
        assert_eq!(evaluator.value("wd12").unwrap(), 100.0 / 12.0);
        assert_eq!(evaluator.value("hd8").unwrap(), 12.5);
        assert_eq!(evaluator.value("hd10").unwrap(), 10.0);
        assert_eq!(evaluator.value("ssd32").unwrap(), 3.125);
        assert_eq!(evaluator.value("3cd8").unwrap(), 8_100_000.0);
        evaluator
            .apply_adjust_values(&adjustments, &overrides)
            .unwrap();
        evaluator.evaluate_guides(&guides).unwrap();

        let path = evaluator
            .evaluate_path(&[
                PathCommand::MoveTo {
                    x: guide("x1"),
                    y: guide("y1"),
                },
                PathCommand::LineTo {
                    x: guide("x2"),
                    y: guide("y2"),
                },
            ])
            .unwrap();
        assert_eq!(
            path,
            [
                EvaluatedPathCommand::MoveTo { x: 20.0, y: 25.0 },
                EvaluatedPathCommand::LineTo { x: 80.0, y: 75.0 },
            ]
        );
    }

    #[test]
    fn ordinary_guides_replace_in_order_without_relaxing_adjustment_validation() {
        let mut evaluator = GuideEvaluator::new(100.0, 100.0).unwrap();
        evaluator
            .evaluate_guides(&[
                Guide::parse("connsiteX0", "val 10").unwrap(),
                Guide::parse("connsiteX0", "val 20").unwrap(),
            ])
            .unwrap();
        assert_eq!(evaluator.value("connsiteX0").unwrap(), 20.0);

        let duplicate_adjustments = [
            Guide::parse("adj", "val 10").unwrap(),
            Guide::parse("adj", "val 20").unwrap(),
        ];
        let error = GuideEvaluator::new(100.0, 100.0)
            .unwrap()
            .apply_adjust_values(&duplicate_adjustments, &BTreeMap::new())
            .unwrap_err();
        assert_eq!(error, GeometryError::DuplicateGuide("adj".to_owned()));

        let error = GuideEvaluator::new(100.0, 100.0)
            .unwrap()
            .apply_adjust_values(
                &[Guide::parse("adj", "val 10").unwrap()],
                &BTreeMap::from([("unknown".to_owned(), 20.0)]),
            )
            .unwrap_err();
        assert_eq!(
            error,
            GeometryError::UnknownAdjustOverride("unknown".to_owned())
        );
    }

    #[test]
    fn all_seventeen_formula_tokens_parse_and_evaluate_with_drawingml_argument_order() {
        let cases = [
            ("*/ 6 7 2", GuideOp::MulDiv, 21.0),
            ("+- 10 4 3", GuideOp::AddSub, 11.0),
            ("+/ 10 4 2", GuideOp::AddDiv, 7.0),
            ("?: 1 5 6", GuideOp::IfElse, 5.0),
            ("abs -7", GuideOp::Abs, 7.0),
            ("at2 1 1", GuideOp::At2, 2_700_000.0),
            ("cat2 10 3 4", GuideOp::Cat2, 6.0),
            ("cos 10 3600000", GuideOp::Cos, 5.0),
            ("max 3 7", GuideOp::Max, 7.0),
            ("min 3 7", GuideOp::Min, 3.0),
            ("mod 3 4 12", GuideOp::Mod, 13.0),
            ("pin 0 15 10", GuideOp::Pin, 10.0),
            ("sat2 10 3 4", GuideOp::Sat2, 8.0),
            ("sin 10 1800000", GuideOp::Sin, 5.0),
            ("sqrt -9", GuideOp::Sqrt, 3.0),
            ("tan 10 2700000", GuideOp::Tan, 10.0),
            ("val 42", GuideOp::Val, 42.0),
        ];
        let guides = cases
            .iter()
            .enumerate()
            .map(|(index, (formula, expected_op, _))| {
                let guide = Guide::parse(format!("g{index}"), formula).unwrap();
                assert_eq!(guide.op, *expected_op);
                guide
            })
            .collect::<Vec<_>>();
        let mut evaluator = GuideEvaluator::new(100.0, 80.0).unwrap();
        evaluator.evaluate_guides(&guides).unwrap();

        for (index, (_, _, expected)) in cases.iter().enumerate() {
            assert_close(evaluator.value(&format!("g{index}")).unwrap(), *expected);
        }
    }

    #[test]
    fn arc_to_is_flattened_to_finite_cubics_with_matching_endpoints() {
        let evaluator = GuideEvaluator::new(100.0, 100.0).unwrap();
        let path = evaluator
            .evaluate_path(&[
                PathCommand::MoveTo {
                    x: literal(10.0),
                    y: literal(0.0),
                },
                PathCommand::ArcTo {
                    width_radius: literal(10.0),
                    height_radius: literal(5.0),
                    start_angle: literal(0.0),
                    sweep_angle: literal(27_000_000.0),
                },
            ])
            .unwrap();
        assert_eq!(path.len(), 6);
        assert!(path.iter().skip(1).all(|command| matches!(
            command,
            EvaluatedPathCommand::CubicTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y
            } if [x1, y1, x2, y2, x, y].iter().all(|value| value.is_finite())
        )));
        let EvaluatedPathCommand::CubicTo { x, y, .. } = path[5] else {
            panic!("arc did not end in a cubic command")
        };
        assert_close(x, 0.0);
        assert_close(y, 5.0);
        let EvaluatedPathCommand::CubicTo { x, y, .. } = path[1] else {
            panic!("arc did not start with a cubic command")
        };
        assert_close(x, 0.0);
        assert_close(y, 5.0);
    }

    #[test]
    fn office_mod_and_negative_sqrt_semantics_produce_finite_values() {
        let guides = [
            Guide::parse("norm", "mod 3 4 12").unwrap(),
            Guide::parse("root", "sqrt -9").unwrap(),
        ];
        let mut evaluator = GuideEvaluator::new(10.0, 10.0).unwrap();
        evaluator.evaluate_guides(&guides).unwrap();
        assert_eq!(evaluator.value("norm").unwrap(), 13.0);
        assert_eq!(evaluator.value("root").unwrap(), 3.0);
    }

    #[test]
    fn division_by_zero_returns_an_error_instead_of_non_finite_coordinates() {
        let mut evaluator = GuideEvaluator::new(10.0, 10.0).unwrap();
        let error = evaluator
            .evaluate_guides(&[Guide::parse("bad", "*/ 1 2 0").unwrap()])
            .unwrap_err();
        assert_eq!(error, GeometryError::DivisionByZero);
        assert_eq!(error.to_string(), "division by zero");
    }

    #[test]
    fn corpus_custom_geometry_round_trips_and_evaluates_to_a_closed_path() {
        let xml = br#"<z:custGeom xmlns:z="http://schemas.openxmlformats.org/drawingml/2006/main"><z:avLst><z:gd name="adj" fmla="val 25000"/></z:avLst><z:gdLst><z:gd name="x1" fmla="*/ w adj 100000"/><z:gd name="x2" fmla="+- r 0 x1"/></z:gdLst><z:rect l="x1" t="t" r="x2" b="b"/><z:pathLst><z:path w="100" h="100"><z:moveTo><z:pt x="l" y="t"/></z:moveTo><z:lnTo><z:pt x="r" y="t"/></z:lnTo><z:cubicBezTo><z:pt x="r" y="t"/><z:pt x="r" y="b"/><z:pt x="x2" y="b"/></z:cubicBezTo><z:close/></z:path></z:pathLst></z:custGeom>"#;
        let geometry = CT_CustomGeometry2D::from_xml(xml).unwrap();
        let evaluated = geometry.evaluate(&BTreeMap::new()).unwrap();

        assert_eq!(geometry.adjust_values().len(), 1);
        assert_eq!(geometry.guides().len(), 2);
        assert_eq!(geometry.paths().len(), 1);
        assert_eq!(evaluated.paths.len(), 1);
        assert_eq!(
            evaluated.paths[0].last(),
            Some(&EvaluatedPathCommand::Close)
        );
        assert_eq!(
            evaluated.text_rectangle,
            Some(EvaluatedTextRectangle {
                left: 25.0,
                top: 0.0,
                right: 75.0,
                bottom: 100.0,
            })
        );

        let written = geometry.to_xml().unwrap();
        let reparsed = CT_CustomGeometry2D::from_xml(&written).unwrap();
        assert_eq!(reparsed, geometry);
    }

    #[test]
    fn custom_geometry_reads_any_prefix_and_writes_fixed_a_prefix_in_schema_order() {
        let xml = br#"<q:custGeom><q:avLst/><q:gdLst/><q:ahLst/><q:cxnLst/><q:rect l="l" t="t" r="r" b="b"/><q:pathLst><q:path w="100" h="80" fill="none" stroke="false" extrusionOk="true"><q:moveTo><q:pt x="0" y="0"/></q:moveTo><q:arcTo wR="10" hR="5" stAng="0" swAng="5400000"/><q:close/></q:path></q:pathLst></q:custGeom>"#;
        let geometry = CT_CustomGeometry2D::from_xml(xml).unwrap();
        let written = String::from_utf8(geometry.to_xml().unwrap()).unwrap();

        assert!(written.starts_with("<a:custGeom>"));
        assert!(written.contains("<a:avLst/>"));
        assert!(written.contains("<a:gdLst/>"));
        assert!(written.contains("<a:rect l=\"l\" t=\"t\" r=\"r\" b=\"b\"/>"));
        assert!(written.contains("<a:arcTo wR=\"10\" hR=\"5\" stAng=\"0\" swAng=\"5400000\"/>"));
        assert!(written.contains("<a:close/>"));
        let av = written.find("<a:avLst").unwrap();
        let gd = written.find("<a:gdLst").unwrap();
        let ah = written.find("<q:ahLst/>").unwrap();
        let cxn = written.find("<q:cxnLst/>").unwrap();
        let rect = written.find("<a:rect").unwrap();
        let paths = written.find("<a:pathLst").unwrap();
        assert!(av < gd && gd < ah && ah < cxn && cxn < rect && rect < paths);
    }

    #[test]
    fn empty_custom_geometry_path_list_from_theme_defaults_round_trips() {
        let geometry = CT_CustomGeometry2D::from_xml(
            br#"<q:custGeom><q:avLst/><q:gdLst/><q:ahLst/><q:cxnLst/><q:rect l="0" t="0" r="0" b="0"/><q:pathLst/></q:custGeom>"#,
        )
        .unwrap();
        assert!(geometry.paths().is_empty());
        let written = geometry.to_xml().unwrap();
        assert_eq!(CT_CustomGeometry2D::from_xml(&written).unwrap(), geometry);
    }

    #[test]
    fn unknown_custom_geometry_children_round_trip_byte_for_byte_in_place() {
        let xml = br#"<a:custGeom><u:before/><a:avLst><u:avBefore/><a:gd name="adj" fmla="val 25000"><u:insideGuide/></a:gd><u:avAfter/></a:avLst><u:middle u:id="7"><u:child/></u:middle><a:pathLst><u:pathBefore/><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"><u:insidePoint/></a:pt><u:insideMove/></a:moveTo><u:between/><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path><u:pathAfter/></a:pathLst><u:after/></a:custGeom>"#;
        let written = String::from_utf8(
            CT_CustomGeometry2D::from_xml(xml)
                .unwrap()
                .to_xml()
                .unwrap(),
        )
        .unwrap();

        for raw in [
            "<u:before/>",
            "<u:avBefore/>",
            "<u:insideGuide/>",
            "<u:avAfter/>",
            "<u:middle u:id=\"7\"><u:child/></u:middle>",
            "<u:pathBefore/>",
            "<u:insidePoint/>",
            "<u:insideMove/>",
            "<u:between/>",
            "<u:pathAfter/>",
            "<u:after/>",
        ] {
            assert!(written.contains(raw), "missing raw subtree {raw}");
        }
        assert!(written.find("<u:before/>").unwrap() < written.find("<a:avLst").unwrap());
        assert!(written.find("<u:avBefore/>").unwrap() < written.find("<a:gd ").unwrap());
        assert!(written.find("<a:gd ").unwrap() < written.find("<u:avAfter/>").unwrap());
        assert!(written.find("<u:insideGuide/>").unwrap() < written.find("</a:gd>").unwrap());
        assert!(written.find("<u:insidePoint/>").unwrap() < written.find("</a:pt>").unwrap());
        assert!(written.find("<u:insideMove/>").unwrap() < written.find("</a:moveTo>").unwrap());
        assert!(written.find("</a:moveTo>").unwrap() < written.find("<u:between/>").unwrap());
    }

    #[test]
    fn malformed_custom_geometry_returns_an_error_without_panicking() {
        let malformed: [&[u8]; 3] = [
            br#"<a:custGeom><a:avLst><a:gd fmla="val 1"/></a:avLst><a:pathLst><a:path w="1" h="1"/></a:pathLst></a:custGeom>"#,
            br#"<a:custGeom><a:gdLst><a:gd name="bad" fmla="nope 1"/></a:gdLst><a:pathLst><a:path w="1" h="1"/></a:pathLst></a:custGeom>"#,
            br#"<a:custGeom><a:pathLst><a:path w="1" h="1">"#,
        ];

        for xml in malformed {
            let result = std::panic::catch_unwind(|| CT_CustomGeometry2D::from_xml(xml));
            assert!(result.is_ok(), "malformed XML panicked");
            assert!(result.unwrap().is_err(), "malformed XML was accepted");
        }
    }

    #[test]
    fn rectangle_preset_evaluates_to_expected_bounds_and_text_rect() {
        let preset = CT_PresetGeometry2D::from_xml(br#"<q:prstGeom prst="rect"/>"#).unwrap();
        let evaluated = preset.evaluate((120.0, 80.0)).unwrap().unwrap();

        assert_eq!(
            evaluated.paths[0],
            [
                EvaluatedPathCommand::MoveTo { x: 0.0, y: 0.0 },
                EvaluatedPathCommand::LineTo { x: 120.0, y: 0.0 },
                EvaluatedPathCommand::LineTo { x: 120.0, y: 80.0 },
                EvaluatedPathCommand::LineTo { x: 0.0, y: 80.0 },
                EvaluatedPathCommand::Close,
            ]
        );
        assert_eq!(
            evaluated.text_rectangle,
            Some(EvaluatedTextRectangle {
                left: 0.0,
                top: 0.0,
                right: 120.0,
                bottom: 80.0,
            })
        );
    }

    #[test]
    fn preset_adjustments_override_generated_defaults() {
        let default = CT_PresetGeometry2D::from_xml(
            br#"<a:prstGeom prst="trapezoid"><a:avLst/></a:prstGeom>"#,
        )
        .unwrap()
        .evaluate((200.0, 100.0))
        .unwrap()
        .unwrap();
        let adjusted = CT_PresetGeometry2D::from_xml(
            br#"<a:prstGeom prst="trapezoid"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>"#,
        )
        .unwrap()
        .evaluate((200.0, 100.0))
        .unwrap()
        .unwrap();

        assert_eq!(
            default.paths[0][1],
            EvaluatedPathCommand::LineTo { x: 25.0, y: 0.0 }
        );
        assert_eq!(
            adjusted.paths[0][1],
            EvaluatedPathCommand::LineTo { x: 50.0, y: 0.0 }
        );
    }
}
