use std::fmt;
use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::units::{Angle, Percent1000};
use oxml_core::xml::{R_NS, get_attr, local_name, matches_local_name};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::color::{ColorChoice, ColorError};
use crate::namespace::reject_conflicting_a_prefix;
use crate::order::OrderedRawChildren;

/// Errors produced while parsing or writing DrawingML fills.
#[derive(Debug)]
pub enum FillError {
    Xml(OxmlError),
    Color(ColorError),
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
    GradientStopOutOfRange(i32),
}

impl fmt::Display for FillError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xml(error) => error.fmt(formatter),
            Self::Color(error) => error.fmt(formatter),
            Self::UnexpectedElement(element) => {
                write!(formatter, "unexpected DrawingML fill element: {element}")
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
            Self::GradientStopOutOfRange(value) => {
                write!(
                    formatter,
                    "DrawingML gradient stop is outside 0 to 100000: {value}"
                )
            }
        }
    }
}

impl std::error::Error for FillError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Xml(error) => Some(error),
            Self::Color(error) => Some(error),
            _ => None,
        }
    }
}

impl From<OxmlError> for FillError {
    fn from(error: OxmlError) -> Self {
        Self::Xml(error)
    }
}

impl From<ColorError> for FillError {
    fn from(error: ColorError) -> Self {
        Self::Color(error)
    }
}

pub type Result<T> = std::result::Result<T, FillError>;

/// One of the five DrawingML fill families.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum Fill {
    NoFill(NoFill),
    Solid(SolidFill),
    Gradient(GradientFill),
    Pattern(PatternFill),
    Blip(BlipFill),
}

impl Fill {
    /// Parses one complete fill element with any namespace prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) => return Self::from_element(&mut reader, &element),
                Event::Empty(element) => return Self::from_empty_element(&element),
                Event::Eof => {
                    return Err(FillError::Xml(OxmlError::MissingElement(
                        "DrawingML fill".to_owned(),
                    )));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Parses a fill after the caller has consumed its start event.
    pub fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        match local_name(start.name().as_ref()) {
            b"noFill" => Ok(Self::NoFill(NoFill::from_element(reader, start)?)),
            b"solidFill" => Ok(Self::Solid(SolidFill::from_element(reader, start)?)),
            b"gradFill" => Ok(Self::Gradient(GradientFill::from_element(reader, start)?)),
            b"pattFill" => Ok(Self::Pattern(PatternFill::from_element(reader, start)?)),
            b"blipFill" => Ok(Self::Blip(BlipFill::from_element(reader, start)?)),
            _ => Err(unexpected(start)),
        }
    }

    /// Parses a self-closing fill element.
    pub fn from_empty_element(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        match local_name(start.name().as_ref()) {
            b"noFill" => Ok(Self::NoFill(NoFill::default())),
            b"solidFill" => Ok(Self::Solid(SolidFill::default())),
            b"gradFill" => Ok(Self::Gradient(GradientFill::from_start(start)?)),
            b"pattFill" => Ok(Self::Pattern(PatternFill::from_start(start)?)),
            b"blipFill" => Ok(Self::Blip(BlipFill::from_start(start)?)),
            _ => Err(unexpected(start)),
        }
    }

    /// Writes this fill with canonical DrawingML prefixes and schema order.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer)?;
        Ok(writer.into_inner())
    }

    /// Writes this fill into an existing XML writer.
    pub fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::NoFill(fill) => fill.write_xml(writer),
            Self::Solid(fill) => fill.write_xml(writer),
            Self::Gradient(fill) => fill.write_xml(writer),
            Self::Pattern(fill) => fill.write_xml(writer),
            Self::Blip(fill) => fill.write_xml(writer),
        }
    }
}

/// An `a:noFill` value, including preserved extension children.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct NoFill {
    raw_children: OrderedRawChildren,
}

impl NoFill {
    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        Ok(Self {
            raw_children: capture_all_children(reader, local_name(start.name().as_ref()))?,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.raw_children.is_empty() {
            write_empty(writer, BytesStart::new("a:noFill"))
        } else {
            write_start(writer, BytesStart::new("a:noFill"))?;
            emit_raw(writer, self.raw_children.at(0))?;
            write_end(writer, "a:noFill")
        }
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

/// A solid fill with its DrawingML colour choice.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct SolidFill {
    pub color: Option<ColorChoice>,
    raw_children: OrderedRawChildren,
}

impl SolidFill {
    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut value = Self::default();
        read_color_children(
            reader,
            local_name(start.name().as_ref()),
            &mut value.color,
            &mut value.raw_children,
        )?;
        Ok(value)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.color.is_none() && self.raw_children.is_empty() {
            return write_empty(writer, BytesStart::new("a:solidFill"));
        }
        write_start(writer, BytesStart::new("a:solidFill"))?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(color) = &self.color {
            color.to_xml(writer)?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        write_end(writer, "a:solidFill")
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

/// One gradient stop. Stops remain in document order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct GradientStop {
    pub position: Percent1000,
    pub color: Option<ColorChoice>,
    raw_children: OrderedRawChildren,
}

impl GradientStop {
    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        let position = required_i32(start, b"pos")?;
        validate_gradient_position(position)?;
        let mut stop = Self {
            position: Percent1000(position),
            color: None,
            raw_children: OrderedRawChildren::default(),
        };
        read_color_children(reader, b"gs", &mut stop.color, &mut stop.raw_children)?;
        Ok(stop)
    }

    fn from_empty(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        let position = required_i32(start, b"pos")?;
        validate_gradient_position(position)?;
        Ok(Self {
            position: Percent1000(position),
            color: None,
            raw_children: OrderedRawChildren::default(),
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        validate_gradient_position(self.position.0)?;
        let position = self.position.0.to_string();
        let mut start = BytesStart::new("a:gs");
        start.push_attribute(("pos", position.as_str()));
        if self.color.is_none() && self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(color) = &self.color {
            color.to_xml(writer)?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        write_end(writer, "a:gs")
    }
}

/// Gradient geometry for either a linear or path gradient.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum GradientGeometry {
    Linear(LinearGradient),
    Path(PathGradient),
}

/// Linear gradient angle and scaling behavior.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct LinearGradient {
    pub angle: Angle,
    pub scaled: Option<bool>,
    raw_children: OrderedRawChildren,
}

/// A DrawingML path-gradient shape.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PathGradientKind {
    Shape,
    Circle,
    Rectangle,
}

impl PathGradientKind {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "shape" => Some(Self::Shape),
            "circle" => Some(Self::Circle),
            "rect" => Some(Self::Rectangle),
            _ => None,
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Shape => "shape",
            Self::Circle => "circle",
            Self::Rectangle => "rect",
        }
    }
}

/// Path gradient shape and optional focal rectangle.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PathGradient {
    pub kind: PathGradientKind,
    pub fill_to_rect: Option<RelativeRect>,
    raw_children: OrderedRawChildren,
}

/// A rectangle expressed as optional percentages from each edge.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct RelativeRect {
    pub left: Option<Percent1000>,
    pub top: Option<Percent1000>,
    pub right: Option<Percent1000>,
    pub bottom: Option<Percent1000>,
    raw_children: OrderedRawChildren,
}

/// A gradient fill with stops in source order.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct GradientFill {
    pub flip: Option<String>,
    pub rotate_with_shape: Option<bool>,
    pub stops: Vec<GradientStop>,
    pub geometry: Option<GradientGeometry>,
    pub tile_rect: Option<RelativeRect>,
    raw_children: OrderedRawChildren,
    stop_list_raw_children: OrderedRawChildren,
}

impl GradientFill {
    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        Ok(Self {
            flip: optional_token(start, b"flip", &["none", "x", "y", "xy"])?,
            rotate_with_shape: optional_bool(start, b"rotWithShape")?,
            ..Self::default()
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut fill = Self::from_start(start)?;
        let mut buffer = Vec::new();
        let mut boundary = 0;
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"gsLst") => {
                    reject_conflicting_a_prefix(&element)?;
                    fill.read_stop_list(reader)?;
                    boundary = 1;
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"lin") => {
                    fill.geometry = Some(GradientGeometry::Linear(parse_linear_element(
                        reader, &element,
                    )?));
                    boundary = 2;
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"lin") => {
                    fill.geometry = Some(GradientGeometry::Linear(parse_linear(&element)?));
                    boundary = 2;
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"path") => {
                    fill.geometry = Some(GradientGeometry::Path(parse_path(reader, &element)?));
                    boundary = 2;
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"path") => {
                    fill.geometry = Some(GradientGeometry::Path(parse_empty_path(&element)?));
                    boundary = 2;
                }
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"tileRect") =>
                {
                    fill.tile_rect = Some(parse_rect_element(reader, &element)?);
                    boundary = 3;
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"tileRect") =>
                {
                    fill.tile_rect = Some(parse_rect(&element)?);
                    boundary = 3;
                }
                Event::Start(element) => fill
                    .raw_children
                    .push(boundary, capture_element(reader, &element)?),
                Event::Empty(element) => fill
                    .raw_children
                    .push(boundary, capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"gradFill") => {
                    break;
                }
                Event::Eof => return Err(missing_end("gradFill")),
                _ => {}
            }
            buffer.clear();
        }
        Ok(fill)
    }

    fn read_stop_list(&mut self, reader: &mut Reader<&[u8]>) -> Result<()> {
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"gs") => {
                    self.stops
                        .push(GradientStop::from_element(reader, &element)?);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"gs") => {
                    self.stops.push(GradientStop::from_empty(&element)?);
                }
                Event::Start(element) => self
                    .stop_list_raw_children
                    .push(self.stops.len(), capture_element(reader, &element)?),
                Event::Empty(element) => self
                    .stop_list_raw_children
                    .push(self.stops.len(), capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"gsLst") => {
                    break;
                }
                Event::Eof => return Err(missing_end("gsLst")),
                _ => {}
            }
            buffer.clear();
        }
        Ok(())
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:gradFill");
        if let Some(flip) = self.flip.as_deref() {
            start.push_attribute(("flip", flip));
        }
        let rotate = self.rotate_with_shape.map(bool_text);
        if let Some(rotate) = rotate {
            start.push_attribute(("rotWithShape", rotate));
        }
        if self.stops.is_empty()
            && self.geometry.is_none()
            && self.tile_rect.is_none()
            && self.stop_list_raw_children.is_empty()
            && self.raw_children.is_empty()
        {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        if !self.stops.is_empty() || !self.stop_list_raw_children.is_empty() {
            write_start(writer, BytesStart::new("a:gsLst"))?;
            for boundary in 0..=self.stops.len() {
                emit_raw(writer, self.stop_list_raw_children.at(boundary))?;
                if let Some(stop) = self.stops.get(boundary) {
                    stop.write_xml(writer)?;
                }
            }
            write_end(writer, "a:gsLst")?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        if let Some(geometry) = &self.geometry {
            write_geometry(writer, geometry)?;
        }
        emit_raw(writer, self.raw_children.at(2))?;
        if let Some(rect) = &self.tile_rect {
            write_rect(writer, "a:tileRect", rect)?;
        }
        emit_raw(writer, self.raw_children.at(3))?;
        write_end(writer, "a:gradFill")
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

/// Pattern fill foreground, background, and preset identifier.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct PatternFill {
    pub preset: Option<String>,
    pub foreground: Option<ColorChoice>,
    pub background: Option<ColorChoice>,
    raw_children: OrderedRawChildren,
    foreground_raw_children: OrderedRawChildren,
    background_raw_children: OrderedRawChildren,
}

impl PatternFill {
    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        Ok(Self {
            preset: get_attr(start, b"prst"),
            ..Self::default()
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut fill = Self::from_start(start)?;
        let mut buffer = Vec::new();
        let mut boundary = 0;
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"fgClr") => {
                    reject_conflicting_a_prefix(&element)?;
                    (fill.foreground, fill.foreground_raw_children) =
                        parse_color_container(reader, b"fgClr")?;
                    boundary = 1;
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"bgClr") => {
                    reject_conflicting_a_prefix(&element)?;
                    (fill.background, fill.background_raw_children) =
                        parse_color_container(reader, b"bgClr")?;
                    boundary = 2;
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"fgClr") => {
                    reject_conflicting_a_prefix(&element)?;
                    boundary = 1
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"bgClr") => {
                    reject_conflicting_a_prefix(&element)?;
                    boundary = 2
                }
                Event::Start(element) => fill
                    .raw_children
                    .push(boundary, capture_element(reader, &element)?),
                Event::Empty(element) => fill
                    .raw_children
                    .push(boundary, capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"pattFill") => {
                    break;
                }
                Event::Eof => return Err(missing_end("pattFill")),
                _ => {}
            }
            buffer.clear();
        }
        Ok(fill)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:pattFill");
        if let Some(preset) = self.preset.as_deref() {
            start.push_attribute(("prst", preset));
        }
        if self.foreground.is_none()
            && self.background.is_none()
            && self.foreground_raw_children.is_empty()
            && self.background_raw_children.is_empty()
            && self.raw_children.is_empty()
        {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        if self.foreground.is_some() || !self.foreground_raw_children.is_empty() {
            write_color_container(
                writer,
                "a:fgClr",
                self.foreground.as_ref(),
                &self.foreground_raw_children,
            )?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        if self.background.is_some() || !self.background_raw_children.is_empty() {
            write_color_container(
                writer,
                "a:bgClr",
                self.background.as_ref(),
                &self.background_raw_children,
            )?;
        }
        emit_raw(writer, self.raw_children.at(2))?;
        write_end(writer, "a:pattFill")
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

/// Picture relationship identifiers and preserved image effects.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct Blip {
    pub embed: Option<String>,
    pub link: Option<String>,
    raw_children: OrderedRawChildren,
}

impl Blip {
    fn from_start(start: &BytesStart<'_>) -> Self {
        Self {
            embed: get_attr(start, b"embed"),
            link: get_attr(start, b"link"),
            raw_children: OrderedRawChildren::default(),
        }
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut blip = Self::from_start(start);
        blip.raw_children = capture_all_children(reader, b"blip")?;
        Ok(blip)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:blip");
        if self.embed.is_some() || self.link.is_some() {
            start.push_attribute(("xmlns:r", R_NS));
        }
        if let Some(embed) = self.embed.as_deref() {
            start.push_attribute(("r:embed", embed));
        }
        if let Some(link) = self.link.as_deref() {
            start.push_attribute(("r:link", link));
        }
        if self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        write_end(writer, "a:blip")
    }

    /// The opacity `a:alphaModFix` gives the picture, where 100000 is opaque.
    ///
    /// The element stays a preserved raw child, so this only reads it. An
    /// omitted `amt` is the schema default of 100%.
    pub fn alpha_modulation_fix(&self) -> Option<Percent1000> {
        self.raw_children.at(0).find_map(|raw| {
            let mut reader = Reader::from_reader(raw);
            loop {
                match reader.read_event() {
                    Ok(Event::Start(element) | Event::Empty(element)) => {
                        if !matches_local_name(element.name().as_ref(), b"alphaModFix") {
                            return None;
                        }
                        return match get_attr(&element, b"amt") {
                            Some(amount) => amount.parse().ok().map(Percent1000),
                            None => Some(Percent1000(100_000)),
                        };
                    }
                    Ok(Event::Eof) | Err(_) => return None,
                    Ok(_) => {}
                }
            }
        })
    }
}

/// Picture-fill placement mode.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum BlipMode {
    Stretch {
        fill_rect: Option<RelativeRect>,
        raw_children: OrderedRawChildren,
    },
    Tile(Tile),
}

/// Tiled picture-fill offsets, scales, flip, and alignment.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct Tile {
    pub translation_x: Option<i64>,
    pub translation_y: Option<i64>,
    pub scale_x: Option<Percent1000>,
    pub scale_y: Option<Percent1000>,
    pub flip: Option<String>,
    pub alignment: Option<String>,
    raw_children: OrderedRawChildren,
}

/// Picture fill with relationship ids but no package dependency.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct BlipFill {
    pub dpi: Option<u32>,
    pub rotate_with_shape: Option<bool>,
    pub blip: Option<Blip>,
    pub source_rect: Option<RelativeRect>,
    pub mode: Option<BlipMode>,
    raw_attributes: Vec<(String, String)>,
    raw_children: OrderedRawChildren,
}

impl BlipFill {
    /// Parses one complete picture-fill element with any namespace prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"blipFill") =>
                {
                    return Self::from_element(&mut reader, &element);
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"blipFill") =>
                {
                    return Self::from_start(&element);
                }
                Event::Start(element) | Event::Empty(element) => {
                    return Err(unexpected(&element));
                }
                Event::Eof => return Err(missing_end("blipFill")),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        Ok(Self {
            dpi: optional_parse(start, b"dpi")?,
            rotate_with_shape: optional_bool(start, b"rotWithShape")?,
            raw_attributes: capture_unmodelled_attributes(start, &[b"dpi", b"rotWithShape"])?,
            ..Self::default()
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut fill = Self::from_start(start)?;
        let mut buffer = Vec::new();
        let mut boundary = 0;
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"blip") => {
                    reject_conflicting_a_prefix(&element)?;
                    fill.blip = Some(Blip::from_element(reader, &element)?);
                    boundary = 1;
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"blip") => {
                    reject_conflicting_a_prefix(&element)?;
                    fill.blip = Some(Blip::from_start(&element));
                    boundary = 1;
                }
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"srcRect") =>
                {
                    fill.source_rect = Some(parse_rect_element(reader, &element)?);
                    boundary = 2;
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"srcRect") =>
                {
                    fill.source_rect = Some(parse_rect(&element)?);
                    boundary = 2;
                }
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"stretch") =>
                {
                    reject_conflicting_a_prefix(&element)?;
                    fill.mode = Some(parse_stretch(reader)?);
                    boundary = 3;
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"stretch") =>
                {
                    reject_conflicting_a_prefix(&element)?;
                    fill.mode = Some(BlipMode::Stretch {
                        fill_rect: None,
                        raw_children: OrderedRawChildren::default(),
                    });
                    boundary = 3;
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"tile") => {
                    fill.mode = Some(BlipMode::Tile(parse_tile(reader, &element)?));
                    boundary = 3;
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"tile") => {
                    fill.mode = Some(BlipMode::Tile(parse_empty_tile(&element)?));
                    boundary = 3;
                }
                Event::Start(element) => fill
                    .raw_children
                    .push(boundary, capture_element(reader, &element)?),
                Event::Empty(element) => fill
                    .raw_children
                    .push(boundary, capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"blipFill") => {
                    break;
                }
                Event::Eof => return Err(missing_end("blipFill")),
                _ => {}
            }
            buffer.clear();
        }
        Ok(fill)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.write_xml_as(writer, "a:blipFill")
    }

    /// Writes the picture fill under the caller's required root name.
    pub fn write_xml_as<W: Write>(&self, writer: &mut Writer<W>, name: &str) -> Result<()> {
        let mut start = BytesStart::new(name);
        let dpi = self.dpi.map(|value| value.to_string());
        if let Some(dpi) = dpi.as_deref() {
            start.push_attribute(("dpi", dpi));
        }
        if let Some(rotate) = self.rotate_with_shape.map(bool_text) {
            start.push_attribute(("rotWithShape", rotate));
        }
        push_raw_attributes(&mut start, &self.raw_attributes);
        if self.blip.is_none()
            && self.source_rect.is_none()
            && self.mode.is_none()
            && self.raw_children.is_empty()
        {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(blip) = &self.blip {
            blip.write_xml(writer)?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        if let Some(rect) = &self.source_rect {
            write_rect(writer, "a:srcRect", rect)?;
        }
        emit_raw(writer, self.raw_children.at(2))?;
        if let Some(mode) = &self.mode {
            write_blip_mode(writer, mode)?;
        }
        emit_raw(writer, self.raw_children.at(3))?;
        write_end(writer, name)
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

fn capture_unmodelled_attributes(
    start: &BytesStart<'_>,
    modelled: &[&[u8]],
) -> Result<Vec<(String, String)>> {
    let mut raw = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(OxmlError::from)?;
        let key = attribute.key.as_ref();
        if !key.contains(&b':') && modelled.contains(&key) {
            continue;
        }
        let name = std::str::from_utf8(key)
            .map_err(OxmlError::from)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
            .map_err(OxmlError::from)?
            .into_owned();
        raw.push((name, value));
    }
    Ok(raw)
}

fn push_raw_attributes(start: &mut BytesStart<'_>, attributes: &[(String, String)]) {
    for (name, value) in attributes {
        start.push_attribute((name.as_str(), value.as_str()));
    }
}

fn parse_linear(start: &BytesStart<'_>) -> Result<LinearGradient> {
    reject_conflicting_a_prefix(start)?;
    Ok(LinearGradient {
        angle: Angle(required_i32(start, b"ang")?),
        scaled: optional_bool(start, b"scaled")?,
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_linear_element(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<LinearGradient> {
    let mut linear = parse_linear(start)?;
    linear.raw_children = capture_all_children(reader, b"lin")?;
    Ok(linear)
}

fn parse_empty_path(start: &BytesStart<'_>) -> Result<PathGradient> {
    reject_conflicting_a_prefix(start)?;
    let value = required_attr(start, b"path")?;
    let kind = PathGradientKind::parse(&value).ok_or_else(|| invalid(start, b"path", value))?;
    Ok(PathGradient {
        kind,
        fill_to_rect: None,
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_path(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<PathGradient> {
    let mut path = parse_empty_path(start)?;
    let mut buffer = Vec::new();
    let mut boundary = 0;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) if matches_local_name(element.name().as_ref(), b"fillToRect") => {
                path.fill_to_rect = Some(parse_rect_element(reader, &element)?);
                boundary = 1;
            }
            Event::Empty(element) if matches_local_name(element.name().as_ref(), b"fillToRect") => {
                path.fill_to_rect = Some(parse_rect(&element)?);
                boundary = 1;
            }
            Event::Start(element) => path
                .raw_children
                .push(boundary, capture_element(reader, &element)?),
            Event::Empty(element) => path
                .raw_children
                .push(boundary, capture_empty_element(&element)?),
            Event::End(element) if matches_local_name(element.name().as_ref(), b"path") => break,
            Event::Eof => return Err(missing_end("path")),
            _ => {}
        }
        buffer.clear();
    }
    Ok(path)
}

fn write_geometry<W: Write>(writer: &mut Writer<W>, geometry: &GradientGeometry) -> Result<()> {
    match geometry {
        GradientGeometry::Linear(linear) => {
            let angle = linear.angle.0.to_string();
            let mut start = BytesStart::new("a:lin");
            start.push_attribute(("ang", angle.as_str()));
            if let Some(scaled) = linear.scaled.map(bool_text) {
                start.push_attribute(("scaled", scaled));
            }
            if linear.raw_children.is_empty() {
                return write_empty(writer, start);
            }
            write_start(writer, start)?;
            emit_raw(writer, linear.raw_children.at(0))?;
            write_end(writer, "a:lin")
        }
        GradientGeometry::Path(path) => {
            let mut start = BytesStart::new("a:path");
            start.push_attribute(("path", path.kind.as_str()));
            if path.fill_to_rect.is_none() && path.raw_children.is_empty() {
                return write_empty(writer, start);
            }
            write_start(writer, start)?;
            emit_raw(writer, path.raw_children.at(0))?;
            if let Some(rect) = &path.fill_to_rect {
                write_rect(writer, "a:fillToRect", rect)?;
            }
            emit_raw(writer, path.raw_children.at(1))?;
            write_end(writer, "a:path")
        }
    }
}

fn parse_stretch(reader: &mut Reader<&[u8]>) -> Result<BlipMode> {
    let mut fill_rect = None;
    let mut raw_children = OrderedRawChildren::default();
    let mut boundary = 0;
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) if matches_local_name(element.name().as_ref(), b"fillRect") => {
                fill_rect = Some(parse_rect_element(reader, &element)?);
                boundary = 1;
            }
            Event::Empty(element) if matches_local_name(element.name().as_ref(), b"fillRect") => {
                fill_rect = Some(parse_rect(&element)?);
                boundary = 1;
            }
            Event::Start(element) => {
                raw_children.push(boundary, capture_element(reader, &element)?)
            }
            Event::Empty(element) => raw_children.push(boundary, capture_empty_element(&element)?),
            Event::End(element) if matches_local_name(element.name().as_ref(), b"stretch") => break,
            Event::Eof => return Err(missing_end("stretch")),
            _ => {}
        }
        buffer.clear();
    }
    Ok(BlipMode::Stretch {
        fill_rect,
        raw_children,
    })
}

fn parse_empty_tile(start: &BytesStart<'_>) -> Result<Tile> {
    reject_conflicting_a_prefix(start)?;
    Ok(Tile {
        translation_x: optional_parse(start, b"tx")?,
        translation_y: optional_parse(start, b"ty")?,
        scale_x: optional_parse::<i32>(start, b"sx")?.map(Percent1000),
        scale_y: optional_parse::<i32>(start, b"sy")?.map(Percent1000),
        flip: optional_token(start, b"flip", &["none", "x", "y", "xy"])?,
        alignment: optional_token(
            start,
            b"algn",
            &["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"],
        )?,
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_tile(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Tile> {
    let mut tile = parse_empty_tile(start)?;
    tile.raw_children = capture_all_children(reader, b"tile")?;
    Ok(tile)
}

fn write_blip_mode<W: Write>(writer: &mut Writer<W>, mode: &BlipMode) -> Result<()> {
    match mode {
        BlipMode::Stretch {
            fill_rect,
            raw_children,
        } => {
            if fill_rect.is_none() && raw_children.is_empty() {
                return write_empty(writer, BytesStart::new("a:stretch"));
            }
            write_start(writer, BytesStart::new("a:stretch"))?;
            emit_raw(writer, raw_children.at(0))?;
            if let Some(rect) = fill_rect {
                write_rect(writer, "a:fillRect", rect)?;
            }
            emit_raw(writer, raw_children.at(1))?;
            write_end(writer, "a:stretch")
        }
        BlipMode::Tile(tile) => {
            let mut start = BytesStart::new("a:tile");
            let tx = tile.translation_x.map(|value| value.to_string());
            let ty = tile.translation_y.map(|value| value.to_string());
            let sx = tile.scale_x.map(|value| value.0.to_string());
            let sy = tile.scale_y.map(|value| value.0.to_string());
            for (name, value) in [
                ("tx", tx.as_deref()),
                ("ty", ty.as_deref()),
                ("sx", sx.as_deref()),
                ("sy", sy.as_deref()),
            ] {
                if let Some(value) = value {
                    start.push_attribute((name, value));
                }
            }
            if let Some(flip) = tile.flip.as_deref() {
                start.push_attribute(("flip", flip));
            }
            if let Some(alignment) = tile.alignment.as_deref() {
                start.push_attribute(("algn", alignment));
            }
            if tile.raw_children.is_empty() {
                return write_empty(writer, start);
            }
            write_start(writer, start)?;
            emit_raw(writer, tile.raw_children.at(0))?;
            write_end(writer, "a:tile")
        }
    }
}

fn parse_rect(start: &BytesStart<'_>) -> Result<RelativeRect> {
    reject_conflicting_a_prefix(start)?;
    Ok(RelativeRect {
        left: optional_parse::<i32>(start, b"l")?.map(Percent1000),
        top: optional_parse::<i32>(start, b"t")?.map(Percent1000),
        right: optional_parse::<i32>(start, b"r")?.map(Percent1000),
        bottom: optional_parse::<i32>(start, b"b")?.map(Percent1000),
        raw_children: OrderedRawChildren::default(),
    })
}

fn parse_rect_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<RelativeRect> {
    let mut rect = parse_rect(start)?;
    rect.raw_children = capture_all_children(reader, local_name(start.name().as_ref()))?;
    Ok(rect)
}

fn write_rect<W: Write>(writer: &mut Writer<W>, name: &str, rect: &RelativeRect) -> Result<()> {
    let left = rect.left.map(|value| value.0.to_string());
    let top = rect.top.map(|value| value.0.to_string());
    let right = rect.right.map(|value| value.0.to_string());
    let bottom = rect.bottom.map(|value| value.0.to_string());
    let mut start = BytesStart::new(name);
    for (name, value) in [
        ("l", left.as_deref()),
        ("t", top.as_deref()),
        ("r", right.as_deref()),
        ("b", bottom.as_deref()),
    ] {
        if let Some(value) = value {
            start.push_attribute((name, value));
        }
    }
    if rect.raw_children.is_empty() {
        return write_empty(writer, start);
    }
    write_start(writer, start)?;
    emit_raw(writer, rect.raw_children.at(0))?;
    write_end(writer, name)
}

fn parse_color_container(
    reader: &mut Reader<&[u8]>,
    end_name: &[u8],
) -> Result<(Option<ColorChoice>, OrderedRawChildren)> {
    let mut color = None;
    let mut raw_children = OrderedRawChildren::default();
    let mut boundary = 0;
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) if is_color(element.name().as_ref()) => {
                color = Some(ColorChoice::from_xml(reader, &element)?);
                boundary = 1;
            }
            Event::Empty(element) if is_color(element.name().as_ref()) => {
                color = Some(ColorChoice::from_empty_xml(&element)?);
                boundary = 1;
            }
            Event::Start(element) => {
                raw_children.push(boundary, capture_element(reader, &element)?);
            }
            Event::Empty(element) => {
                raw_children.push(boundary, capture_empty_element(&element)?);
            }
            Event::End(element) if matches_local_name(element.name().as_ref(), end_name) => break,
            Event::Eof => return Err(missing_end(&String::from_utf8_lossy(end_name))),
            _ => {}
        }
        buffer.clear();
    }
    Ok((color, raw_children))
}

fn write_color_container<W: Write>(
    writer: &mut Writer<W>,
    name: &str,
    color: Option<&ColorChoice>,
    raw_children: &OrderedRawChildren,
) -> Result<()> {
    write_start(writer, BytesStart::new(name))?;
    emit_raw(writer, raw_children.at(0))?;
    if let Some(color) = color {
        color.to_xml(writer)?;
    }
    emit_raw(writer, raw_children.at(1))?;
    write_end(writer, name)
}

fn read_color_children(
    reader: &mut Reader<&[u8]>,
    end_name: &[u8],
    color: &mut Option<ColorChoice>,
    raw_children: &mut OrderedRawChildren,
) -> Result<()> {
    let mut buffer = Vec::new();
    let mut boundary = 0;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) if is_color(element.name().as_ref()) => {
                *color = Some(ColorChoice::from_xml(reader, &element)?);
                boundary = 1;
            }
            Event::Empty(element) if is_color(element.name().as_ref()) => {
                *color = Some(ColorChoice::from_empty_xml(&element)?);
                boundary = 1;
            }
            Event::Start(element) => {
                raw_children.push(boundary, capture_element(reader, &element)?)
            }
            Event::Empty(element) => raw_children.push(boundary, capture_empty_element(&element)?),
            Event::End(element) if matches_local_name(element.name().as_ref(), end_name) => break,
            Event::Eof => return Err(missing_end(&String::from_utf8_lossy(end_name))),
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn capture_all_children(reader: &mut Reader<&[u8]>, end_name: &[u8]) -> Result<OrderedRawChildren> {
    let mut raw_children = OrderedRawChildren::default();
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) => raw_children.push(0, capture_element(reader, &element)?),
            Event::Empty(element) => raw_children.push(0, capture_empty_element(&element)?),
            Event::End(element) if matches_local_name(element.name().as_ref(), end_name) => break,
            Event::Eof => return Err(missing_end(&String::from_utf8_lossy(end_name))),
            _ => {}
        }
        buffer.clear();
    }
    Ok(raw_children)
}

fn is_color(name: &[u8]) -> bool {
    matches!(
        local_name(name),
        b"srgbClr" | b"schemeClr" | b"sysClr" | b"prstClr"
    )
}

fn validate_gradient_position(value: i32) -> Result<()> {
    if (0..=100_000).contains(&value) {
        Ok(())
    } else {
        Err(FillError::GradientStopOutOfRange(value))
    }
}

fn required_attr(start: &BytesStart<'_>, name: &[u8]) -> Result<String> {
    get_attr(start, name).ok_or_else(|| FillError::MissingAttribute {
        element: String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned(),
        attribute: String::from_utf8_lossy(name).into_owned(),
    })
}

fn required_i32(start: &BytesStart<'_>, name: &[u8]) -> Result<i32> {
    let value = required_attr(start, name)?;
    value.parse().map_err(|_| invalid(start, name, value))
}

fn optional_parse<T: std::str::FromStr>(start: &BytesStart<'_>, name: &[u8]) -> Result<Option<T>> {
    get_attr(start, name)
        .map(|value| value.parse().map_err(|_| invalid(start, name, value)))
        .transpose()
}

fn optional_bool(start: &BytesStart<'_>, name: &[u8]) -> Result<Option<bool>> {
    get_attr(start, name)
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(start, name, value)),
        })
        .transpose()
}

fn optional_token(start: &BytesStart<'_>, name: &[u8], allowed: &[&str]) -> Result<Option<String>> {
    get_attr(start, name)
        .map(|value| {
            if allowed.contains(&value.as_str()) {
                Ok(value)
            } else {
                Err(invalid(start, name, value))
            }
        })
        .transpose()
}

fn invalid(start: &BytesStart<'_>, attribute: &[u8], value: String) -> FillError {
    FillError::InvalidAttribute {
        element: String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned(),
        attribute: String::from_utf8_lossy(attribute).into_owned(),
        value,
    }
}

fn unexpected(start: &BytesStart<'_>) -> FillError {
    FillError::UnexpectedElement(String::from_utf8_lossy(start.name().as_ref()).into_owned())
}

fn missing_end(name: &str) -> FillError {
    FillError::Xml(OxmlError::MissingElement(format!("closing a:{name}")))
}

const fn bool_text(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}

fn emit_raw<'a, W: Write>(
    writer: &mut Writer<W>,
    children: impl Iterator<Item = &'a [u8]>,
) -> Result<()> {
    for child in children {
        writer.get_mut().write_all(child).map_err(OxmlError::from)?;
    }
    Ok(())
}

fn write_start<W: Write>(writer: &mut Writer<W>, start: BytesStart<'_>) -> Result<()> {
    writer
        .write_event(Event::Start(start))
        .map_err(OxmlError::from)?;
    Ok(())
}

fn write_empty<W: Write>(writer: &mut Writer<W>, start: BytesStart<'_>) -> Result<()> {
    writer
        .write_event(Event::Empty(start))
        .map_err(OxmlError::from)?;
    Ok(())
}

fn write_end<W: Write>(writer: &mut Writer<W>, name: &str) -> Result<()> {
    writer
        .write_event(Event::End(BytesEnd::new(name)))
        .map_err(OxmlError::from)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{Fill, FillError, GradientGeometry};

    #[test]
    fn every_fill_form_round_trips_and_gradient_stops_keep_document_order() {
        let cases: &[&[u8]] = &[
            br#"<a:noFill/>"#,
            br#"<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>"#,
            br#"<a:gradFill flip="xy" rotWithShape="1"><a:gsLst><a:gs pos="75000"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="25000"><a:schemeClr val="accent2"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/><a:tileRect l="1000"/></a:gradFill>"#,
            br#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="10000" t="10000" r="10000" b="10000"/></a:path></a:gradFill>"#,
            br#"<a:pattFill prst="pct10"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:schemeClr val="bg1"/></a:bgClr></a:pattFill>"#,
            br#"<a:blipFill dpi="96" rotWithShape="0"><a:blip r:embed="rId7"/><a:srcRect l="1000" t="2000"/><a:stretch><a:fillRect r="3000"/></a:stretch></a:blipFill>"#,
            br#"<a:blipFill><a:blip r:link="rId8"/><a:tile tx="10" ty="20" sx="50000" sy="75000" flip="x" algn="ctr"/></a:blipFill>"#,
        ];

        for xml in cases {
            let parsed = Fill::from_xml(xml).unwrap();
            let written = parsed.to_xml().unwrap();
            assert_eq!(Fill::from_xml(&written).unwrap(), parsed);
        }

        let Fill::Gradient(gradient) = Fill::from_xml(cases[2]).unwrap() else {
            panic!("expected gradient")
        };
        assert_eq!(
            gradient
                .stops
                .iter()
                .map(|stop| stop.position.0)
                .collect::<Vec<_>>(),
            vec![75_000, 25_000]
        );
        assert!(matches!(
            gradient.geometry,
            Some(GradientGeometry::Linear(_))
        ));
    }

    #[test]
    fn fill_forms_read_any_prefix_and_write_fixed_a_prefix_in_schema_order() {
        let fill = Fill::from_xml(br#"<z:blipFill><z:blip r:embed="rId4"/><z:srcRect l="10"/><z:stretch><z:fillRect t="20"/></z:stretch></z:blipFill>"#).unwrap();
        assert_eq!(fill.to_xml().unwrap(), br#"<a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId4"/><a:srcRect l="10"/><a:stretch><a:fillRect t="20"/></a:stretch></a:blipFill>"#);

        let gradient = Fill::from_xml(br#"<q:gradFill><q:gsLst><q:gs pos="0"><q:srgbClr val="ABCDEF"/></q:gs></q:gsLst><q:path path="rect"/></q:gradFill>"#).unwrap();
        assert_eq!(gradient.to_xml().unwrap(), br#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst><a:path path="rect"/></a:gradFill>"#);
    }

    #[test]
    fn blip_relationship_prefix_is_declared_in_canonical_output() {
        let fill = Fill::from_xml(
            br#"<a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></a:blipFill>"#,
        )
        .unwrap();
        let written = String::from_utf8(fill.to_xml().unwrap()).unwrap();

        assert!(written.contains(
            "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\"/>"
        ));
    }

    #[test]
    fn blip_reports_alpha_modulation_fix_from_its_preserved_children() {
        let blip = |children: &str| {
            let xml =
                format!(r#"<z:blipFill><z:blip r:embed="rId1">{children}</z:blip></z:blipFill>"#);
            super::BlipFill::from_xml(xml.as_bytes())
                .unwrap()
                .blip
                .expect("blip")
        };
        assert_eq!(
            blip(r#"<a:alphaModFix amt="30000"/>"#).alpha_modulation_fix(),
            Some(super::Percent1000(30_000))
        );
        assert_eq!(
            blip(r#"<x:ext/><q:alphaModFix/>"#).alpha_modulation_fix(),
            Some(super::Percent1000(100_000))
        );
        assert_eq!(
            blip(r#"<a:lum bright="10000"/>"#).alpha_modulation_fix(),
            None
        );
        assert_eq!(blip("").alpha_modulation_fix(), None);
    }

    #[test]
    fn unknown_fill_children_round_trip_byte_for_byte_in_place() {
        let xml = br#"<z:gradFill><x:before x:id="1"><x:item>one &amp; two</x:item><!--note--></x:before><z:gsLst><y:first/><z:gs pos="0"><z:srgbClr val="000000"><x:transform value="kept"/></z:srgbClr></z:gs><y:last/></z:gsLst><x:middle/><z:lin ang="0"/><x:after/></z:gradFill>"#;
        let written = Fill::from_xml(xml).unwrap().to_xml().unwrap();
        assert_eq!(written, br#"<a:gradFill><x:before x:id="1"><x:item>one &amp; two</x:item><!--note--></x:before><a:gsLst><y:first/><a:gs pos="0"><a:srgbClr val="000000"><x:transform value="kept"/></a:srgbClr></a:gs><y:last/></a:gsLst><x:middle/><a:lin ang="0"/><x:after/></a:gradFill>"#);

        let blip = br#"<z:blipFill><z:blip r:embed="rId1"><a:alphaModFix amt="50000"/><x:ext x:v="raw"/></z:blip><x:between/><z:tile><x:tileExt/></z:tile></z:blipFill>"#;
        assert_eq!(Fill::from_xml(blip).unwrap().to_xml().unwrap(), br#"<a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"><a:alphaModFix amt="50000"/><x:ext x:v="raw"/></a:blip><x:between/><a:tile><x:tileExt/></a:tile></a:blipFill>"#);

        let pattern = br#"<z:pattFill><z:fgClr><x:before/><z:srgbClr val="102030"/><x:after x:v="kept"/></z:fgClr></z:pattFill>"#;
        assert_eq!(Fill::from_xml(pattern).unwrap().to_xml().unwrap(), br#"<a:pattFill><a:fgClr><x:before/><a:srgbClr val="102030"/><x:after x:v="kept"/></a:fgClr></a:pattFill>"#);

        let extension_only_pattern =
            br#"<z:pattFill><z:fgClr><x:extension x:v="kept"/></z:fgClr></z:pattFill>"#;
        assert_eq!(
            Fill::from_xml(extension_only_pattern)
                .unwrap()
                .to_xml()
                .unwrap(),
            br#"<a:pattFill><a:fgClr><x:extension x:v="kept"/></a:fgClr></a:pattFill>"#
        );

        let gradient_leaf_extensions = br#"<z:gradFill><z:lin ang="0"><x:linExt/></z:lin><z:tileRect l="10"><x:tileRectExt/></z:tileRect></z:gradFill>"#;
        assert_eq!(Fill::from_xml(gradient_leaf_extensions).unwrap().to_xml().unwrap(), br#"<a:gradFill><a:lin ang="0"><x:linExt/></a:lin><a:tileRect l="10"><x:tileRectExt/></a:tileRect></a:gradFill>"#);

        let extension_only_stop_list =
            br#"<z:gradFill><z:gsLst><x:extension x:v="kept"/></z:gsLst></z:gradFill>"#;
        assert_eq!(
            Fill::from_xml(extension_only_stop_list)
                .unwrap()
                .to_xml()
                .unwrap(),
            br#"<a:gradFill><a:gsLst><x:extension x:v="kept"/></a:gsLst></a:gradFill>"#
        );

        let path_leaf_extension = br#"<z:gradFill><z:path path="rect"><z:fillToRect><x:pathRectExt/></z:fillToRect></z:path></z:gradFill>"#;
        assert_eq!(Fill::from_xml(path_leaf_extension).unwrap().to_xml().unwrap(), br#"<a:gradFill><a:path path="rect"><a:fillToRect><x:pathRectExt/></a:fillToRect></a:path></a:gradFill>"#);

        let blip_leaf_extensions = br#"<z:blipFill><z:srcRect><x:sourceExt/></z:srcRect><z:stretch><z:fillRect><x:fillExt/></z:fillRect></z:stretch></z:blipFill>"#;
        assert_eq!(Fill::from_xml(blip_leaf_extensions).unwrap().to_xml().unwrap(), br#"<a:blipFill><a:srcRect><x:sourceExt/></a:srcRect><a:stretch><a:fillRect><x:fillExt/></a:fillRect></a:stretch></a:blipFill>"#);
    }

    #[test]
    fn malformed_fill_values_return_errors_without_panicking() {
        let cases: &[&[u8]] = &[
            br#"<a:gradFill rotWithShape="maybe"/>"#,
            br#"<a:gradFill><a:gsLst><a:gs pos="100001"/></a:gsLst></a:gradFill>"#,
            br#"<a:gradFill><a:gsLst><a:gs/></a:gsLst></a:gradFill>"#,
            br#"<a:gradFill><a:lin ang="ninety"/></a:gradFill>"#,
            br#"<a:gradFill><a:path path="triangle"/></a:gradFill>"#,
            br#"<a:blipFill dpi="many"/>"#,
            br#"<a:blipFill><a:srcRect l="half"/></a:blipFill>"#,
            br#"<a:blipFill><a:tile sx="large"/></a:blipFill>"#,
            br#"<a:gradFill flip="diagonal"/>"#,
            br#"<a:blipFill><a:tile algn="middle"/></a:blipFill>"#,
        ];
        for xml in cases {
            let result = std::panic::catch_unwind(|| Fill::from_xml(xml));
            assert!(
                result.is_ok(),
                "parser panicked for {}",
                String::from_utf8_lossy(xml)
            );
            assert!(
                result.unwrap().is_err(),
                "malformed fill parsed successfully"
            );
        }
        assert!(matches!(
            Fill::from_xml(cases[1]),
            Err(FillError::GradientStopOutOfRange(100_001))
        ));
    }
}
