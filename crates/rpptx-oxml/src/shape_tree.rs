use std::collections::{HashMap, HashSet};
use std::io::Write;
use std::ops::Range;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_drawing::fill::Fill;
use oxml_drawing::geometry::CT_PresetGeometry2D;
use oxml_drawing::namespace::A_NS;
use oxml_drawing::order::OrderedRawChildren;
use oxml_drawing::shape_props::CT_ShapeProperties;
use oxml_drawing::style_ref::{FontReference, StyleMatrixReference, StyleReference};
use oxml_drawing::text::CT_TextBody;
use oxml_drawing::xfrm::CT_Transform2D;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer, XmlVersion};

use crate::connector::CT_ConnectionShape;
use crate::graphic_frame::CT_GraphicFrame;
use crate::namespace::{
    FIXED_SHAPE_TREE_PREFIXES, MC_NS, NamespaceBindings, P_NS, R_NS, all_attributes,
    non_visual_drawing_id, non_visual_drawing_name, root_attributes, self_contained_attributes,
    set_non_visual_drawing_name,
};
use crate::picture::CT_Picture;
use crate::placeholder::{ApplicationProperties, CT_Placeholder, parse_application_properties};

pub type Result<T> = std::result::Result<T, OxmlError>;
type RawAttributes = Vec<(String, String)>;

/// One shape-tree child in document and z-order.
// The public PresentationML model intentionally stores CT_Picture directly.
#[allow(clippy::large_enum_variant)]
#[derive(Clone, Debug, PartialEq)]
pub enum ShapeTreeChild {
    Shape(CT_Shape),
    Picture(CT_Picture),
    GraphicFrame(Box<CT_GraphicFrame>),
    GroupShape(Box<CT_GroupShape>),
    Connector(CT_ConnectionShape),
    AlternateContent(Box<CT_AlternateContent>),
}

impl ShapeTreeChild {
    /// Returns the child's `p:cNvPr/@id`, when the child owns one.
    pub fn non_visual_id(&self) -> Option<u32> {
        match self {
            Self::Shape(shape) => shape.non_visual_id(),
            Self::Picture(picture) => picture.non_visual_id(),
            Self::GraphicFrame(frame) => frame.non_visual_id(),
            Self::GroupShape(group) => group.non_visual_id(),
            Self::Connector(connector) => connector.non_visual_id(),
            Self::AlternateContent(alternate) => alternate
                .chart_choice()
                .and_then(CT_GraphicFrame::non_visual_id),
        }
    }

    /// Returns the decoded producer-facing non-visual shape name.
    pub fn non_visual_name(&self) -> Option<String> {
        match self {
            Self::Shape(shape) => shape.non_visual_name(),
            Self::Picture(picture) => picture.non_visual_name(),
            Self::GraphicFrame(frame) => frame.non_visual_name(),
            Self::GroupShape(group) => group.non_visual_name(),
            Self::Connector(connector) => connector.non_visual_name(),
            Self::AlternateContent(alternate) => alternate
                .chart_choice()
                .and_then(CT_GraphicFrame::non_visual_name),
        }
        .map(str::to_owned)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Shape(shape) => shape.write_xml_internal(writer, false)?,
            Self::Picture(picture) => picture.write_xml_internal(writer, false)?,
            Self::GraphicFrame(frame) => frame.write_xml_internal(writer, false)?,
            Self::Connector(connector) => connector.write_xml_internal(writer, false)?,
            Self::AlternateContent(alternate) => alternate.write_xml(writer)?,
            Self::GroupShape(group) => group.write_xml_internal(writer, false)?,
        }
        Ok(())
    }
}

/// Allocates shape ids that are unused across a complete shape tree.
#[derive(Clone, Debug)]
pub struct ShapeIdAllocator {
    occupied: HashSet<u32>,
    next: u32,
}

impl ShapeIdAllocator {
    /// Scans typed and preserved shape-tree content for occupied ids.
    pub fn scan(tree: &CT_ShapeTree) -> Self {
        let mut occupied = HashSet::new();
        collect_non_visual_ids(&tree.children, &mut occupied);
        if let Ok(xml) = tree.to_xml() {
            collect_preserved_non_visual_ids(&xml, &mut occupied);
        }
        Self { occupied, next: 2 }
    }

    /// Returns and reserves the next free id, starting at 2.
    pub fn allocate(&mut self) -> u32 {
        loop {
            let candidate = self.next;
            self.next = self.next.checked_add(1).unwrap_or(2);
            if self.occupied.insert(candidate) {
                return candidate;
            }
        }
    }
}

#[derive(Debug)]
struct ShapeIdOccurrence {
    range: Range<usize>,
    id: u32,
    defines_shape: bool,
}

/// Assigns fresh non-visual shape ids and rewrites connector endpoints.
pub fn rewrite_shape_ids(raw: &[u8]) -> Result<Vec<u8>> {
    let occurrences = shape_id_occurrences(raw)?;
    let occupied = occurrences
        .iter()
        .filter(|occurrence| occurrence.defines_shape)
        .map(|occurrence| occurrence.id)
        .collect::<HashSet<_>>();
    let mut next = 2u32;
    let mut map = HashMap::new();
    for occurrence in occurrences
        .iter()
        .filter(|occurrence| occurrence.defines_shape)
    {
        if occurrence.id == 1 || map.contains_key(&occurrence.id) {
            continue;
        }
        while occupied.contains(&next) {
            next = next.checked_add(1).ok_or_else(shape_ids_exhausted)?;
        }
        map.insert(occurrence.id, next);
        next = next.checked_add(1).ok_or_else(shape_ids_exhausted)?;
    }
    if map.is_empty() {
        return Ok(raw.to_vec());
    }

    let mut rewritten = Vec::with_capacity(raw.len());
    let mut copied_through = 0usize;
    for occurrence in occurrences {
        let Some(target) = map.get(&occurrence.id) else {
            continue;
        };
        if occurrence.range.start < copied_through || occurrence.range.end > raw.len() {
            return Err(OxmlError::InvalidValue(
                "shape-id replacement ranges overlap".to_owned(),
            ));
        }
        rewritten.extend_from_slice(&raw[copied_through..occurrence.range.start]);
        rewritten.extend_from_slice(target.to_string().as_bytes());
        copied_through = occurrence.range.end;
    }
    rewritten.extend_from_slice(&raw[copied_through..]);
    Ok(rewritten)
}

fn shape_id_occurrences(raw: &[u8]) -> Result<Vec<ShapeIdOccurrence>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut scopes = vec![NamespaceBindings::default()];
    let mut occurrences = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let scope = scopes
                    .last()
                    .ok_or_else(|| OxmlError::InvalidValue("missing namespace scope".to_owned()))?
                    .with_start(&element)?;
                collect_shape_id_occurrence(raw, event_start, &element, &scope, &mut occurrences)?;
                scopes.push(scope);
            }
            Event::Empty(element) => {
                let scope = scopes
                    .last()
                    .ok_or_else(|| OxmlError::InvalidValue("missing namespace scope".to_owned()))?
                    .with_start(&element)?;
                collect_shape_id_occurrence(raw, event_start, &element, &scope, &mut occurrences)?;
            }
            Event::End(_) => {
                if scopes.len() == 1 {
                    return Err(OxmlError::InvalidValue(
                        "shape XML has an unmatched closing tag".to_owned(),
                    ));
                }
                scopes.pop();
            }
            Event::Eof => {
                if scopes.len() != 1 {
                    return Err(OxmlError::InvalidValue(
                        "shape XML ended before its root closed".to_owned(),
                    ));
                }
                return Ok(occurrences);
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_shape_id_occurrence(
    raw: &[u8],
    event_start: usize,
    element: &BytesStart<'_>,
    scope: &NamespaceBindings,
    occurrences: &mut Vec<ShapeIdOccurrence>,
) -> Result<()> {
    let element_name = element.name();
    let name = local_name(element_name.as_ref());
    let uri = scope.element_uri(element_name.as_ref());
    let defines_shape = uri == Some(P_NS) && name == b"cNvPr";
    let connector_endpoint = uri == Some(A_NS) && matches!(name, b"stCxn" | b"endCxn");
    let presentation_shape_reference = uri == Some(P_NS)
        && matches!(
            name,
            b"spTgt" | b"inkTgt" | b"bldP" | b"bldDgm" | b"bldGraphic" | b"bldOleChart"
        );
    if !defines_shape && !connector_endpoint && !presentation_shape_reference {
        return Ok(());
    }
    for attribute in element.attributes() {
        let attribute = attribute?;
        let expected_attribute = if presentation_shape_reference {
            b"spid".as_slice()
        } else {
            b"id".as_slice()
        };
        if attribute.key.as_ref() != expected_attribute {
            continue;
        }
        let Ok(id) = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .parse()
        else {
            continue;
        };
        let value = attribute.value.as_ref();
        let relative_start = (value.as_ptr() as usize)
            .checked_sub(element.as_ptr() as usize)
            .filter(|start| {
                start
                    .checked_add(value.len())
                    .is_some_and(|end| end <= element.len())
            })
            .ok_or_else(|| {
                OxmlError::InvalidValue("shape-id attribute is outside its start tag".to_owned())
            })?;
        let start = event_start
            .checked_add(1)
            .and_then(|start| start.checked_add(relative_start))
            .ok_or_else(|| OxmlError::InvalidValue("shape-id byte range overflowed".to_owned()))?;
        let end = start
            .checked_add(value.len())
            .ok_or_else(|| OxmlError::InvalidValue("shape-id byte range overflowed".to_owned()))?;
        if raw.get(start..end) != Some(value) {
            return Err(OxmlError::InvalidValue(
                "shape-id byte range did not match the source".to_owned(),
            ));
        }
        occurrences.push(ShapeIdOccurrence {
            range: start..end,
            id,
            defines_shape,
        });
    }
    Ok(())
}

fn shape_ids_exhausted() -> OxmlError {
    OxmlError::InvalidValue("PowerPoint shape ids are exhausted".to_owned())
}

fn collect_preserved_non_visual_ids(xml: &[u8], occupied: &mut HashSet<u32>) {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element) | Event::Empty(element)))
                if is_presentationml_namespace(&namespace)
                    && local_name(element.name().as_ref()) == b"cNvPr" =>
            {
                if let Some(id) = all_attributes(&element)
                    .ok()
                    .and_then(|attributes| attributes.into_iter().find(|(name, _)| name == "id"))
                    .and_then(|(_, value)| value.parse().ok())
                {
                    occupied.insert(id);
                }
            }
            Ok((_, Event::Eof)) | Err(_) => return,
            _ => {}
        }
        buffer.clear();
    }
}

fn is_presentationml_namespace(namespace: &ResolveResult<'_>) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => *uri == P_NS.as_bytes(),
        ResolveResult::Unknown(prefix) => prefix == b"p",
        ResolveResult::Unbound => false,
    }
}

fn collect_non_visual_ids(children: &[ShapeTreeChild], occupied: &mut HashSet<u32>) {
    for child in children {
        if let Some(id) = child.non_visual_id() {
            occupied.insert(id);
        }
        match child {
            ShapeTreeChild::GroupShape(group) => {
                collect_non_visual_ids(&group.children, occupied);
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(fallback) = alternate.selected_fallback() {
                    collect_non_visual_ids(fallback, occupied);
                }
            }
            _ => {}
        }
    }
}

/// One preserved `mc:AlternateContent` subtree with its render fallback.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_AlternateContent {
    raw_xml: Vec<u8>,
    chart_choice: Option<Box<CT_GraphicFrame>>,
    selected_fallback: Option<Vec<ShapeTreeChild>>,
}

impl CT_AlternateContent {
    /// Returns the original subtree used as the sole serialisation source.
    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }

    /// Returns ordered typed members from the immediate `mc:Fallback` branch.
    pub fn selected_fallback(&self) -> Option<&[ShapeTreeChild]> {
        self.selected_fallback.as_deref()
    }

    /// Returns the first immediate chart-bearing compatibility choice.
    pub fn chart_choice(&self) -> Option<&CT_GraphicFrame> {
        self.chart_choice.as_deref()
    }

    /// Returns the first immediate typed picture in the paired fallback branch.
    pub fn picture_fallback(&self) -> Option<&CT_Picture> {
        self.selected_fallback.as_deref()?.iter().find_map(|child| {
            if let ShapeTreeChild::Picture(picture) = child {
                Some(picture)
            } else {
                None
            }
        })
    }

    fn from_fragment(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces =
                        NamespaceBindings::from_entries(inherited).with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"AlternateContent"
                        || namespaces.element_uri(start.name().as_ref()) != Some(MC_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    return Self::from_reader(&mut reader, &namespaces, xml);
                }
                Event::Empty(start) => {
                    let namespaces =
                        NamespaceBindings::from_entries(inherited).with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"AlternateContent"
                        || namespaces.element_uri(start.name().as_ref()) != Some(MC_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    return Ok(Self {
                        raw_xml: xml.to_vec(),
                        chart_choice: None,
                        selected_fallback: None,
                    });
                }
                Event::Eof => {
                    return Err(OxmlError::MissingElement("mc:AlternateContent".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_reader(
        reader: &mut Reader<&[u8]>,
        namespaces: &NamespaceBindings,
        raw_xml: &[u8],
    ) -> Result<Self> {
        let mut chart_choice = None;
        let mut selected_fallback = None;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let is_fallback = local_name(child.name().as_ref()) == b"Fallback"
                        && child_namespaces.element_uri(child.name().as_ref()) == Some(MC_NS);
                    let is_choice = local_name(child.name().as_ref()) == b"Choice"
                        && child_namespaces.element_uri(child.name().as_ref()) == Some(MC_NS);
                    let raw = capture_element(reader, &child)?;
                    if is_choice && chart_choice.is_none() {
                        chart_choice = parse_chart_choice(&raw, &child_namespaces.entries())?;
                    } else if is_fallback {
                        if selected_fallback.is_some() {
                            return Err(duplicate_fallback());
                        }
                        selected_fallback = Some(parse_compatibility_members(
                            &raw,
                            &child_namespaces.entries(),
                            b"Fallback",
                        )?);
                    }
                }
                Event::Empty(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    if local_name(child.name().as_ref()) == b"Fallback"
                        && child_namespaces.element_uri(child.name().as_ref()) == Some(MC_NS)
                    {
                        if selected_fallback.is_some() {
                            return Err(duplicate_fallback());
                        }
                        selected_fallback = Some(Vec::new());
                    }
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"AlternateContent" => {
                    return Ok(Self {
                        raw_xml: raw_xml.to_vec(),
                        chart_choice,
                        selected_fallback,
                    });
                }
                Event::Eof => {
                    return Err(OxmlError::MissingElement(
                        "closing mc:AlternateContent".to_owned(),
                    ));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.get_mut().write_all(&self.raw_xml)?;
        Ok(())
    }
}

fn parse_chart_choice(
    xml: &[u8],
    inherited: &[(String, String)],
) -> Result<Option<Box<CT_GraphicFrame>>> {
    let mut reader = Reader::from_reader(xml);
    let mut namespaces = NamespaceBindings::from_entries(inherited);
    let mut inside_choice = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if !inside_choice => {
                namespaces = namespaces.with_start(&start)?;
                if local_name(start.name().as_ref()) != b"Choice"
                    || namespaces.element_uri(start.name().as_ref()) != Some(MC_NS)
                {
                    return Err(unexpected(&start));
                }
                inside_choice = true;
            }
            Event::Start(child) if inside_choice => {
                let child_namespaces = namespaces.with_start(&child)?;
                let is_frame = local_name(child.name().as_ref()) == b"graphicFrame"
                    && child_namespaces.element_uri(child.name().as_ref()) == Some(P_NS);
                let raw = capture_element(&mut reader, &child)?;
                if is_frame && graphic_frame_projects_chart(&raw, &child_namespaces)? {
                    let frame = CT_GraphicFrame::from_fragment(&raw, &child_namespaces.entries())?;
                    return Ok(Some(Box::new(frame)));
                }
            }
            Event::Empty(child) if inside_choice => {
                let child_namespaces = namespaces.with_start(&child)?;
                if local_name(child.name().as_ref()) == b"graphicFrame"
                    && child_namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                {
                    let raw = capture_empty_element(&child)?;
                    if graphic_frame_projects_chart(&raw, &child_namespaces)? {
                        let frame =
                            CT_GraphicFrame::from_fragment(&raw, &child_namespaces.entries())?;
                        return Ok(Some(Box::new(frame)));
                    }
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"Choice" => return Ok(None),
            Event::Eof => return Err(OxmlError::MissingElement("closing mc:Choice".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn graphic_frame_projects_chart(raw: &[u8], inherited: &NamespaceBindings) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    let mut scopes = vec![inherited.clone()];
    let mut graphic_depth = None;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let depth = scopes.len() - 1;
                let namespaces = scopes
                    .last()
                    .expect("namespace root is retained")
                    .with_start(&start)?;
                if depth == 1
                    && local_name(start.name().as_ref()) == b"graphic"
                    && namespaces.element_uri(start.name().as_ref()) == Some(A_NS)
                {
                    graphic_depth = Some(depth);
                }
                let chart_data = depth == 2
                    && graphic_depth == Some(1)
                    && is_chart_graphic_data(&start, &namespaces)?;
                scopes.push(namespaces);
                if chart_data {
                    return Ok(true);
                }
            }
            Event::Empty(start) => {
                let depth = scopes.len() - 1;
                let namespaces = scopes
                    .last()
                    .expect("namespace root is retained")
                    .with_start(&start)?;
                if depth == 2
                    && graphic_depth == Some(1)
                    && is_chart_graphic_data(&start, &namespaces)?
                {
                    return Ok(true);
                }
            }
            Event::End(_) => {
                if graphic_depth == Some(scopes.len().saturating_sub(2)) {
                    graphic_depth = None;
                }
                if scopes.len() > 1 {
                    scopes.pop();
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn is_chart_graphic_data(
    start: &quick_xml::events::BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<bool> {
    Ok(local_name(start.name().as_ref()) == b"graphicData"
        && namespaces.element_uri(start.name().as_ref()) == Some(A_NS)
        && all_attributes(start)?.iter().any(|(name, value)| {
            name == "uri" && value == "http://schemas.openxmlformats.org/drawingml/2006/chart"
        }))
}

fn duplicate_fallback() -> OxmlError {
    OxmlError::InvalidValue(
        "mc:AlternateContent may contain at most one immediate mc:Fallback".to_owned(),
    )
}

/// A partial typed `p:sp` model that owns its placeholder identity.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_Shape {
    pub placeholder: Option<CT_Placeholder>,
    pub shape_properties: Box<CT_ShapeProperties>,
    pub text_body: Option<CT_TextBody>,
    raw: Box<ShapeRaw>,
}

/// The four ordered DrawingML references carried by an ordinary shape style.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_ShapeStyle {
    pub line_reference: StyleMatrixReference,
    pub fill_reference: StyleMatrixReference,
    pub effect_reference: StyleMatrixReference,
    pub font_reference: FontReference,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct ShapeRaw {
    raw_attributes: RawAttributes,
    non_visual_attributes: RawAttributes,
    non_visual_children: OrderedRawChildren,
    non_visual_drawing_properties: Vec<u8>,
    non_visual_name: Option<String>,
    non_visual_shape_properties: Vec<u8>,
    application_properties_attributes: RawAttributes,
    application_properties_raw_children: OrderedRawChildren,
    style: Option<Box<CT_ShapeStyle>>,
    raw_children: OrderedRawChildren,
}

impl CT_ShapeStyle {
    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Box<Self>> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"style"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    return Self::from_reader(&mut reader, &start, &namespaces);
                }
                Event::Empty(start) => {
                    return Err(OxmlError::MissingElement(format!(
                        "{} requires a:lnRef, a:fillRef, a:effectRef, and a:fontRef",
                        String::from_utf8_lossy(start.name().as_ref())
                    )));
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:style".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_reader(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        namespaces: &NamespaceBindings,
    ) -> Result<Box<Self>> {
        let mut line_reference = None;
        let mut fill_reference = None;
        let mut effect_reference = None;
        let mut font_reference = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0usize;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_element(reader, &child)?;
                    capture_style_child(
                        &name,
                        uri,
                        raw,
                        &mut line_reference,
                        &mut fill_reference,
                        &mut effect_reference,
                        &mut font_reference,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::Empty(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_empty_element(&child)?;
                    capture_style_child(
                        &name,
                        uri,
                        raw,
                        &mut line_reference,
                        &mut fill_reference,
                        &mut effect_reference,
                        &mut font_reference,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"style" => break,
                Event::Eof => {
                    return Err(OxmlError::MissingElement("closing p:style".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }

        Ok(Box::new(Self {
            line_reference: required(line_reference, "a:lnRef")?,
            fill_reference: required(fill_reference, "a:fillRef")?,
            effect_reference: required(effect_reference, "a:effectRef")?,
            font_reference: required(font_reference, "a:fontRef")?,
            raw_attributes: root_attributes(start, FIXED_SHAPE_TREE_PREFIXES)?,
            raw_children,
        }))
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:style");
        push_attributes(&mut start, &self.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        StyleReference::Line(self.line_reference.clone())
            .write_xml(writer)
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        emit_raw(writer, self.raw_children.at(1))?;
        StyleReference::Fill(self.fill_reference.clone())
            .write_xml(writer)
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        emit_raw(writer, self.raw_children.at(2))?;
        StyleReference::Effect(self.effect_reference.clone())
            .write_xml(writer)
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        emit_raw(writer, self.raw_children.at(3))?;
        StyleReference::Font(self.font_reference.clone())
            .write_xml(writer)
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        emit_raw(writer, self.raw_children.at(4))?;
        writer.write_event(Event::End(BytesEnd::new("p:style")))?;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn capture_style_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    line_reference: &mut Option<StyleMatrixReference>,
    fill_reference: &mut Option<StyleMatrixReference>,
    effect_reference: &mut Option<StyleMatrixReference>,
    font_reference: &mut Option<FontReference>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    let expected_boundary = match (uri, name) {
        (Some(A_NS), b"lnRef") => Some(0),
        (Some(A_NS), b"fillRef") => Some(1),
        (Some(A_NS), b"effectRef") => Some(2),
        (Some(A_NS), b"fontRef") => Some(3),
        _ => None,
    };
    let Some(expected_boundary) = expected_boundary else {
        raw_children.push(*boundary, raw);
        return Ok(());
    };
    if *boundary != expected_boundary {
        return Err(OxmlError::InvalidValue(
            "p:style children must be a:lnRef, a:fillRef, a:effectRef, and a:fontRef in order"
                .to_owned(),
        ));
    }

    let reference = StyleReference::from_xml(&raw)
        .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
    match reference {
        StyleReference::Line(reference) => *line_reference = Some(reference),
        StyleReference::Fill(reference) => *fill_reference = Some(reference),
        StyleReference::Effect(reference) => *effect_reference = Some(reference),
        StyleReference::Font(reference) => *font_reference = Some(reference),
    }
    *boundary += 1;
    Ok(())
}

/// The recursive group-shape form used inside a shape tree.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_GroupShape {
    pub children: Vec<ShapeTreeChild>,
    non_visual_group_properties: NonVisualGroupProperties,
    group_properties: GroupProperties,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

/// The ordered `p:spTree` root shared by slides, layouts, and masters.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_ShapeTree {
    pub children: Vec<ShapeTreeChild>,
    non_visual_group_properties: NonVisualGroupProperties,
    group_properties: GroupProperties,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct GroupProperties {
    transform: Option<CT_Transform2D>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct NonVisualGroupProperties {
    non_visual_id: Option<u32>,
    non_visual_name: Option<String>,
    drawing_properties_index: Option<usize>,
    raw_attributes: RawAttributes,
    raw_children: Vec<Vec<u8>>,
}

#[derive(Clone, Copy)]
enum GroupKind {
    ShapeTree,
    GroupShape,
}

impl GroupKind {
    const fn local_name(self) -> &'static [u8] {
        match self {
            Self::ShapeTree => b"spTree",
            Self::GroupShape => b"grpSp",
        }
    }

    const fn tag(self) -> &'static str {
        match self {
            Self::ShapeTree => "p:spTree",
            Self::GroupShape => "p:grpSp",
        }
    }
}

struct ParsedGroup {
    non_visual_group_properties: NonVisualGroupProperties,
    group_properties: GroupProperties,
    children: Vec<ShapeTreeChild>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

impl CT_Shape {
    /// Creates an ordinary preset shape with a transform and minimal text body.
    pub fn new_preset(
        id: u32,
        name: &str,
        preset: &str,
        transform: CT_Transform2D,
    ) -> Result<Self> {
        Self::new_with_shell(id, name, preset, transform, false)
    }

    /// Creates a textbox with a rectangular no-fill shell and required paragraph.
    pub fn new_textbox(id: u32, name: &str, transform: CT_Transform2D) -> Result<Self> {
        Self::new_with_shell(id, name, "rect", transform, true)
    }

    fn new_with_shell(
        id: u32,
        name: &str,
        preset: &str,
        transform: CT_Transform2D,
        textbox: bool,
    ) -> Result<Self> {
        let text_body = CT_TextBody::from_xml(
            br#"<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>"#,
        )
        .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        let fill = textbox
            .then(|| Fill::from_xml(br#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#))
            .transpose()
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        let mut shape_properties = CT_ShapeProperties::default();
        shape_properties.transform = Some(transform);
        shape_properties.preset_geometry = Some(
            CT_PresetGeometry2D::new(preset)
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
        );
        shape_properties.fill = fill;
        let mut shape = Self {
            placeholder: None,
            shape_properties: Box::new(shape_properties),
            text_body: Some(text_body),
            raw: Box::new(ShapeRaw {
                raw_attributes: Vec::new(),
                non_visual_attributes: Vec::new(),
                non_visual_children: OrderedRawChildren::default(),
                non_visual_drawing_properties: format!(r#"<p:cNvPr id="{id}"/>"#).into_bytes(),
                non_visual_name: None,
                non_visual_shape_properties: if textbox {
                    br#"<p:cNvSpPr txBox="1"/>"#.to_vec()
                } else {
                    b"<p:cNvSpPr/>".to_vec()
                },
                application_properties_attributes: Vec::new(),
                application_properties_raw_children: OrderedRawChildren::default(),
                style: None,
                raw_children: OrderedRawChildren::default(),
            }),
        };
        shape.set_name(name)?;
        Ok(shape)
    }

    /// Creates a minimal placeholder shape whose position inherits from layout.
    pub fn new_placeholder(id: u32, placeholder: CT_Placeholder) -> Result<Self> {
        let text_body = CT_TextBody::from_xml(
            br#"<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>"#,
        )
        .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        Ok(Self {
            placeholder: Some(placeholder),
            shape_properties: Box::new(CT_ShapeProperties::default()),
            text_body: Some(text_body),
            raw: Box::new(ShapeRaw {
                raw_attributes: Vec::new(),
                non_visual_attributes: Vec::new(),
                non_visual_children: OrderedRawChildren::default(),
                non_visual_drawing_properties: format!(
                    r#"<p:cNvPr id="{id}" name="Placeholder {id}"/>"#
                )
                .into_bytes(),
                non_visual_name: Some(format!("Placeholder {id}")),
                non_visual_shape_properties: b"<p:cNvSpPr/>".to_vec(),
                application_properties_attributes: Vec::new(),
                application_properties_raw_children: OrderedRawChildren::default(),
                style: None,
                raw_children: OrderedRawChildren::default(),
            }),
        })
    }

    pub(crate) fn non_visual_id(&self) -> Option<u32> {
        non_visual_drawing_id(&self.raw.non_visual_drawing_properties)
    }

    pub(crate) fn non_visual_name(&self) -> Option<&str> {
        self.raw.non_visual_name.as_deref()
    }

    /// Changes the producer-facing non-visual shape name.
    pub fn set_name(&mut self, name: &str) -> Result<()> {
        set_non_visual_drawing_name(&mut self.raw.non_visual_drawing_properties, name)?;
        self.raw.non_visual_name = Some(name.to_owned());
        Ok(())
    }

    /// Parses a complete `p:sp` with any prefix bound to PresentationML.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_fragment(xml, &[])
    }

    fn from_fragment(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces =
                        NamespaceBindings::from_entries(inherited).with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"sp"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
                    return Self::from_reader(&mut reader, &start, &namespaces, inherited);
                }
                Event::Empty(start) => {
                    return Err(OxmlError::MissingElement(format!(
                        "{} requires p:nvSpPr and p:spPr",
                        String::from_utf8_lossy(start.name().as_ref())
                    )));
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:sp".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_reader(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        namespaces: &NamespaceBindings,
        inherited: &[(String, String)],
    ) -> Result<Self> {
        let mut non_visual = None;
        let mut shape_properties = None;
        let mut style = None;
        let mut text_body = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0usize;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_element(reader, &child)?;
                    capture_shape_child(
                        &name,
                        uri,
                        raw,
                        &child_namespaces,
                        &mut non_visual,
                        &mut shape_properties,
                        &mut style,
                        &mut text_body,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::Empty(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_empty_element(&child)?;
                    capture_shape_child(
                        &name,
                        uri,
                        raw,
                        &child_namespaces,
                        &mut non_visual,
                        &mut shape_properties,
                        &mut style,
                        &mut text_body,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"sp" => break,
                Event::Eof => return Err(OxmlError::MissingElement("closing p:sp".to_owned())),
                _ => {}
            }
            buffer.clear();
        }

        let non_visual = required(non_visual, "p:nvSpPr")?;
        Ok(Self {
            placeholder: non_visual.placeholder,
            shape_properties: Box::new(required(shape_properties, "p:spPr")?),
            text_body,
            raw: Box::new(ShapeRaw {
                raw_attributes: self_contained_attributes(
                    start,
                    FIXED_SHAPE_TREE_PREFIXES,
                    inherited,
                )?,
                non_visual_attributes: non_visual.raw_attributes,
                non_visual_children: non_visual.raw_children,
                non_visual_drawing_properties: non_visual.non_visual_drawing_properties,
                non_visual_name: non_visual.non_visual_name,
                non_visual_shape_properties: non_visual.non_visual_shape_properties,
                application_properties_attributes: non_visual.application_properties_attributes,
                application_properties_raw_children: non_visual.application_properties_raw_children,
                style,
                raw_children,
            }),
        })
    }

    /// Returns the optional typed format-scheme references.
    pub fn style(&self) -> Option<&CT_ShapeStyle> {
        self.raw.style.as_deref()
    }

    /// Returns whether `p:cNvSpPr/@txBox` marks this shape as a text box.
    pub fn is_textbox(&self) -> bool {
        let mut reader = Reader::from_reader(self.raw.non_visual_shape_properties.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start) | Event::Empty(start)) => {
                    return all_attributes(&start).is_ok_and(|attributes| {
                        attributes.iter().any(|(name, value)| {
                            name == "txBox" && matches!(value.as_str(), "1" | "true")
                        })
                    });
                }
                Ok(Event::Eof) | Err(_) => return false,
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Serialises a self-contained shape with fixed modelled prefixes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml_internal(&mut writer, true)?;
        Ok(writer.into_inner())
    }

    fn write_xml_internal<W: Write>(
        &self,
        writer: &mut Writer<W>,
        declare_namespaces: bool,
    ) -> Result<()> {
        let mut start = BytesStart::new("p:sp");
        if declare_namespaces {
            start.push_attribute(("xmlns:p", P_NS));
            start.push_attribute(("xmlns:a", A_NS));
            start.push_attribute(("xmlns:r", R_NS));
            start.push_attribute(("xmlns:mc", MC_NS));
        }
        push_attributes(&mut start, &self.raw.raw_attributes);
        writer.write_event(Event::Start(start))?;

        emit_raw(writer, self.raw.raw_children.at(0))?;
        self.write_non_visual_properties(writer)?;
        emit_raw(writer, self.raw.raw_children.at(1))?;
        self.shape_properties
            .write_xml_as(writer, "p:spPr")
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        emit_raw(writer, self.raw.raw_children.at(2))?;
        if let Some(style) = &self.raw.style {
            style.write_xml(writer)?;
        }
        emit_raw(writer, self.raw.raw_children.at(3))?;
        if let Some(text_body) = &self.text_body {
            text_body
                .write_xml_as(writer, "p:txBody")
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        }
        emit_raw(writer, self.raw.raw_children.at(4))?;
        emit_raw(writer, self.raw.raw_children.at(5))?;
        writer.write_event(Event::End(BytesEnd::new("p:sp")))?;
        Ok(())
    }

    fn write_non_visual_properties<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:nvSpPr");
        push_attributes(&mut start, &self.raw.non_visual_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw.non_visual_children.at(0))?;
        writer
            .get_mut()
            .write_all(&self.raw.non_visual_drawing_properties)?;
        emit_raw(writer, self.raw.non_visual_children.at(1))?;
        writer
            .get_mut()
            .write_all(&self.raw.non_visual_shape_properties)?;
        emit_raw(writer, self.raw.non_visual_children.at(2))?;

        let mut application = BytesStart::new("p:nvPr");
        push_attributes(
            &mut application,
            &self.raw.application_properties_attributes,
        );
        if self.placeholder.is_none() && self.raw.application_properties_raw_children.is_empty() {
            writer.write_event(Event::Empty(application))?;
        } else {
            writer.write_event(Event::Start(application))?;
            emit_raw(writer, self.raw.application_properties_raw_children.at(0))?;
            if let Some(placeholder) = &self.placeholder {
                placeholder.write_xml(writer, false)?;
            }
            emit_raw(writer, self.raw.application_properties_raw_children.at(1))?;
            writer.write_event(Event::End(BytesEnd::new("p:nvPr")))?;
        }

        emit_raw(writer, self.raw.non_visual_children.at(3))?;
        writer.write_event(Event::End(BytesEnd::new("p:nvSpPr")))?;
        Ok(())
    }
}

struct ParsedNonVisualShape {
    placeholder: Option<CT_Placeholder>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
    non_visual_drawing_properties: Vec<u8>,
    non_visual_name: Option<String>,
    non_visual_shape_properties: Vec<u8>,
    application_properties_attributes: RawAttributes,
    application_properties_raw_children: OrderedRawChildren,
}

#[allow(clippy::too_many_arguments)]
fn capture_shape_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    non_visual: &mut Option<ParsedNonVisualShape>,
    shape_properties: &mut Option<CT_ShapeProperties>,
    style: &mut Option<Box<CT_ShapeStyle>>,
    text_body: &mut Option<CT_TextBody>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if uri == Some(P_NS) && matches!(name, b"nvSpPr" | b"spPr" | b"style") {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
    }
    match (uri, name) {
        (Some(P_NS), b"nvSpPr") => {
            if *boundary != 0 || non_visual.is_some() {
                return Err(OxmlError::InvalidValue(
                    "p:nvSpPr must be the first modelled p:sp child".to_owned(),
                ));
            }
            *non_visual = Some(parse_non_visual_shape(&raw, namespaces)?);
            *boundary = 1;
        }
        (Some(P_NS), b"spPr") => {
            if *boundary != 1 || shape_properties.is_some() {
                return Err(OxmlError::InvalidValue(
                    "p:spPr must immediately follow p:nvSpPr".to_owned(),
                ));
            }
            *shape_properties = Some(
                CT_ShapeProperties::from_xml(&raw)
                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
            );
            *boundary = 2;
        }
        (Some(P_NS), b"style") => {
            if *boundary != 2 || style.is_some() {
                return Err(OxmlError::InvalidValue(
                    "p:style must follow p:spPr and precede p:txBody".to_owned(),
                ));
            }
            *style = Some(CT_ShapeStyle::from_fragment(&raw, namespaces)?);
            *boundary = 3;
        }
        (Some(P_NS), b"txBody") => {
            if !matches!(*boundary, 2 | 3) || text_body.is_some() {
                return Err(OxmlError::InvalidValue(
                    "p:txBody must follow p:spPr and optional p:style".to_owned(),
                ));
            }
            namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
            *text_body = Some(
                CT_TextBody::from_xml(&raw)
                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
            );
            *boundary = 4;
        }
        (Some(P_NS), b"extLst") => {
            if !matches!(*boundary, 2..=4) {
                return Err(OxmlError::InvalidValue(
                    "p:extLst must be the final p:sp child".to_owned(),
                ));
            }
            raw_children.push(4, raw);
            *boundary = 5;
        }
        _ => raw_children.push(*boundary, raw),
    }
    Ok(())
}

fn parse_non_visual_shape(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<ParsedNonVisualShape> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                let raw_attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?;
                let mut drawing = None;
                let mut drawing_name = None;
                let mut shape = None;
                let mut application = None;
                let mut raw_children = OrderedRawChildren::default();
                let mut boundary = 0usize;
                loop {
                    buffer.clear();
                    match reader.read_event_into(&mut buffer)? {
                        Event::Start(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            let name = local_name(child.name().as_ref()).to_vec();
                            let uri = child_namespaces.element_uri(child.name().as_ref());
                            let candidate_name = if uri == Some(P_NS) && name == b"cNvPr" {
                                non_visual_drawing_name(&child)?
                            } else {
                                None
                            };
                            let raw = capture_element(&mut reader, &child)?;
                            capture_non_visual_shape_child(
                                &name,
                                uri,
                                raw,
                                candidate_name,
                                &child_namespaces,
                                &mut drawing,
                                &mut drawing_name,
                                &mut shape,
                                &mut application,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::Empty(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            let name = local_name(child.name().as_ref()).to_vec();
                            let uri = child_namespaces.element_uri(child.name().as_ref());
                            let candidate_name = if uri == Some(P_NS) && name == b"cNvPr" {
                                non_visual_drawing_name(&child)?
                            } else {
                                None
                            };
                            let raw = capture_empty_element(&child)?;
                            capture_non_visual_shape_child(
                                &name,
                                uri,
                                raw,
                                candidate_name,
                                &child_namespaces,
                                &mut drawing,
                                &mut drawing_name,
                                &mut shape,
                                &mut application,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::End(end) if local_name(end.name().as_ref()) == b"nvSpPr" => break,
                        Event::Eof => {
                            return Err(OxmlError::MissingElement("closing p:nvSpPr".to_owned()));
                        }
                        _ => {}
                    }
                }
                let application = required(application, "p:nvPr")?;
                let drawing = required(drawing, "p:cNvPr")?;
                return Ok(ParsedNonVisualShape {
                    placeholder: application.placeholder,
                    raw_attributes,
                    raw_children,
                    non_visual_name: drawing_name,
                    non_visual_drawing_properties: drawing,
                    non_visual_shape_properties: required(shape, "p:cNvSpPr")?,
                    application_properties_attributes: application.raw_attributes,
                    application_properties_raw_children: application.raw_children,
                });
            }
            Event::Empty(_) => {
                return Err(OxmlError::MissingElement(
                    "p:nvSpPr requires p:cNvPr, p:cNvSpPr, and p:nvPr".to_owned(),
                ));
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:nvSpPr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

#[allow(clippy::too_many_arguments)]
fn capture_non_visual_shape_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    candidate_name: Option<String>,
    namespaces: &NamespaceBindings,
    drawing: &mut Option<Vec<u8>>,
    drawing_name: &mut Option<String>,
    shape: &mut Option<Vec<u8>>,
    application: &mut Option<ApplicationProperties>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if uri == Some(P_NS) && matches!(name, b"cNvPr" | b"cNvSpPr" | b"nvPr") {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
    }
    match (uri, name) {
        (Some(P_NS), b"cNvPr") if *boundary == 0 && drawing.is_none() => {
            *drawing = Some(raw);
            *drawing_name = candidate_name;
            *boundary = 1;
        }
        (Some(P_NS), b"cNvSpPr") if *boundary == 1 && shape.is_none() => {
            *shape = Some(raw);
            *boundary = 2;
        }
        (Some(P_NS), b"nvPr") if *boundary == 2 && application.is_none() => {
            *application = Some(parse_application_properties(&raw, namespaces)?);
            *boundary = 3;
        }
        (Some(P_NS), b"cNvPr" | b"cNvSpPr" | b"nvPr") => {
            return Err(OxmlError::InvalidValue(
                "p:nvSpPr children must be p:cNvPr, p:cNvSpPr, and p:nvPr in order".to_owned(),
            ));
        }
        _ => raw_children.push(*boundary, raw),
    }
    Ok(())
}

impl CT_ShapeTree {
    /// Creates the required empty slide shape tree with root id 1.
    pub fn new() -> Self {
        Self {
            non_visual_group_properties: NonVisualGroupProperties {
                non_visual_id: Some(1),
                non_visual_name: Some(String::new()),
                drawing_properties_index: Some(0),
                raw_attributes: Vec::new(),
                raw_children: vec![
                    b"<p:cNvPr id=\"1\" name=\"\"/>".to_vec(),
                    b"<p:cNvGrpSpPr/>".to_vec(),
                    b"<p:nvPr/>".to_vec(),
                ],
            },
            group_properties: GroupProperties {
                transform: None,
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            children: Vec::new(),
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        }
    }

    /// Returns the shape tree's own `p:cNvPr/@id`, when present.
    pub fn non_visual_id(&self) -> Option<u32> {
        self.non_visual_group_properties.non_visual_id
    }

    /// Appends one member before preserved schema-final shape-tree content.
    pub fn append_child(&mut self, child: ShapeTreeChild) -> &mut ShapeTreeChild {
        self.raw_children.shift_boundaries_from(self.children.len());
        self.children.push(child);
        self.children
            .last_mut()
            .expect("shape-tree child was appended")
    }

    /// Removes one immediate child identified by `p:cNvPr/@id`.
    pub fn remove_child_by_id(&mut self, id: u32) -> Result<Option<ShapeTreeChild>> {
        let Some(index) = self
            .children
            .iter()
            .position(|child| child.non_visual_id() == Some(id))
        else {
            return Ok(None);
        };
        let removed = self.children[index].clone();
        let xml = self.to_xml()?;
        let range = direct_shape_child_range(&xml, id)?.ok_or_else(|| {
            OxmlError::InvalidValue(format!("shape id {id} disappeared during removal"))
        })?;
        let mut rewritten = Vec::with_capacity(xml.len() - range.len());
        rewritten.extend_from_slice(&xml[..range.start]);
        rewritten.extend_from_slice(&xml[range.end..]);
        *self = Self::from_xml(&rewritten)?;
        Ok(Some(removed))
    }

    /// Parses a complete `p:spTree` with any prefix bound to PresentationML.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_fragment(xml, &[])
    }

    pub(crate) fn from_fragment(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let parsed = parse_group(xml, inherited, GroupKind::ShapeTree)?;
        Ok(Self {
            non_visual_group_properties: parsed.non_visual_group_properties,
            children: parsed.children,
            group_properties: parsed.group_properties,
            raw_attributes: parsed.raw_attributes,
            raw_children: parsed.raw_children,
        })
    }

    /// Returns the typed DrawingML group transform when `p:grpSpPr` has one.
    pub fn group_transform(&self) -> Option<&CT_Transform2D> {
        self.group_properties.transform.as_ref()
    }

    /// Serialises a self-contained shape-tree fragment with fixed prefixes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer)?;
        Ok(writer.into_inner())
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_group(
            writer,
            GroupKind::ShapeTree,
            &self.non_visual_group_properties,
            &self.group_properties,
            &self.children,
            &self.raw_attributes,
            &self.raw_children,
            true,
        )
    }
}

fn direct_shape_child_range(xml: &[u8], id: u32) -> Result<Option<Range<usize>>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(NamespaceBindings::default().with_start(&start)?);
            }
            Event::Start(child) => {
                let namespaces = root.as_ref().expect("shape-tree root").with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = namespaces.element_uri(child.name().as_ref());
                let start = shape_start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                let range = start..reader.buffer_position() as usize;
                if parse_shape_tree_child(&name, uri, &raw, &namespaces)?
                    .is_some_and(|child| child.non_visual_id() == Some(id))
                {
                    return Ok(Some(range));
                }
            }
            Event::Empty(child) => {
                let namespaces = root.as_ref().expect("shape-tree root").with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = namespaces.element_uri(child.name().as_ref());
                let range = shape_start_tag_range(xml, reader.buffer_position() as usize)?;
                let raw = capture_empty_element(&child)?;
                if parse_shape_tree_child(&name, uri, &raw, &namespaces)?
                    .is_some_and(|child| child.non_visual_id() == Some(id))
                {
                    return Ok(Some(range));
                }
            }
            Event::End(_) | Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn shape_start_tag_range(xml: &[u8], end: usize) -> Result<Range<usize>> {
    let start = xml[..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .ok_or_else(|| OxmlError::InvalidValue("shape child start tag is missing".to_owned()))?;
    Ok(start..end)
}

impl Default for CT_ShapeTree {
    fn default() -> Self {
        Self::new()
    }
}

impl CT_GroupShape {
    /// Creates an empty group with the required non-visual and property shells.
    pub fn new_empty(id: u32, name: &str) -> Self {
        let mut group = Self {
            children: Vec::new(),
            non_visual_group_properties: NonVisualGroupProperties {
                non_visual_id: Some(id),
                non_visual_name: None,
                drawing_properties_index: Some(0),
                raw_attributes: Vec::new(),
                raw_children: vec![
                    format!(r#"<p:cNvPr id="{id}"/>"#).into_bytes(),
                    b"<p:cNvGrpSpPr/>".to_vec(),
                    b"<p:nvPr/>".to_vec(),
                ],
            },
            group_properties: GroupProperties {
                transform: None,
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        };
        group
            .set_name(name)
            .expect("new group always contains p:cNvPr");
        group
    }

    pub(crate) fn non_visual_id(&self) -> Option<u32> {
        self.non_visual_group_properties.non_visual_id
    }

    pub(crate) fn non_visual_name(&self) -> Option<&str> {
        self.non_visual_group_properties.non_visual_name.as_deref()
    }

    /// Parses a complete recursive `p:grpSp` element.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let parsed = parse_group(xml, &[], GroupKind::GroupShape)?;
        Ok(Self::from_parsed(parsed))
    }

    fn from_fragment(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        Ok(Self::from_parsed(parse_group(
            xml,
            inherited,
            GroupKind::GroupShape,
        )?))
    }

    fn from_parsed(parsed: ParsedGroup) -> Self {
        Self {
            non_visual_group_properties: parsed.non_visual_group_properties,
            children: parsed.children,
            group_properties: parsed.group_properties,
            raw_attributes: parsed.raw_attributes,
            raw_children: parsed.raw_children,
        }
    }

    /// Returns the typed DrawingML group transform when `p:grpSpPr` has one.
    pub fn group_transform(&self) -> Option<&CT_Transform2D> {
        self.group_properties.transform.as_ref()
    }

    /// Returns the group transform, creating an empty one when absent.
    pub fn group_transform_mut(&mut self) -> &mut CT_Transform2D {
        if self.group_properties.transform.is_none() {
            self.group_properties.raw_children.shift_boundaries_from(0);
        }
        self.group_properties
            .transform
            .get_or_insert_with(CT_Transform2D::default)
    }

    /// Changes the producer-facing non-visual group name.
    pub fn set_name(&mut self, name: &str) -> Result<()> {
        let drawing_properties = self
            .non_visual_group_properties
            .raw_children
            .get_mut(
                self.non_visual_group_properties
                    .drawing_properties_index
                    .ok_or_else(|| OxmlError::MissingElement("p:cNvPr".to_owned()))?,
            )
            .ok_or_else(|| OxmlError::MissingElement("p:cNvPr".to_owned()))?;
        set_non_visual_drawing_name(drawing_properties, name)?;
        self.non_visual_group_properties.non_visual_name = Some(name.to_owned());
        Ok(())
    }

    /// Serialises a self-contained group-shape fragment with fixed prefixes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml_internal(&mut writer, true)?;
        Ok(writer.into_inner())
    }

    fn write_xml_internal<W: Write>(
        &self,
        writer: &mut Writer<W>,
        declare_namespaces: bool,
    ) -> Result<()> {
        write_group(
            writer,
            GroupKind::GroupShape,
            &self.non_visual_group_properties,
            &self.group_properties,
            &self.children,
            &self.raw_attributes,
            &self.raw_children,
            declare_namespaces,
        )
    }
}

fn parse_group(xml: &[u8], inherited: &[(String, String)], kind: GroupKind) -> Result<ParsedGroup> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = NamespaceBindings::from_entries(inherited).with_start(&start)?;
                if local_name(start.name().as_ref()) != kind.local_name()
                    || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                {
                    return Err(unexpected(&start));
                }
                namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
                return parse_group_children(&mut reader, &start, &namespaces, kind);
            }
            Event::Empty(start) => {
                let namespaces = NamespaceBindings::from_entries(inherited).with_start(&start)?;
                if local_name(start.name().as_ref()) != kind.local_name()
                    || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                {
                    return Err(unexpected(&start));
                }
                return Err(OxmlError::MissingElement(format!(
                    "{} requires p:nvGrpSpPr and p:grpSpPr",
                    kind.tag()
                )));
            }
            Event::Eof => return Err(OxmlError::MissingElement(kind.tag().to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_group_children(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
    kind: GroupKind,
) -> Result<ParsedGroup> {
    let raw_attributes = root_attributes(start, FIXED_SHAPE_TREE_PREFIXES)?;
    let mut non_visual = None;
    let mut group_properties = None;
    let mut children = Vec::new();
    let mut raw_children = OrderedRawChildren::default();
    let mut state = 0usize;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_namespaces = namespaces.with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = child_namespaces.element_uri(child.name().as_ref());
                let raw = capture_element(reader, &child)?;
                capture_group_child(
                    &name,
                    uri,
                    raw,
                    &child_namespaces,
                    &mut non_visual,
                    &mut group_properties,
                    &mut children,
                    &mut raw_children,
                    &mut state,
                )?;
            }
            Event::Empty(child) => {
                let child_namespaces = namespaces.with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = child_namespaces.element_uri(child.name().as_ref());
                let raw = capture_empty_element(&child)?;
                capture_group_child(
                    &name,
                    uri,
                    raw,
                    &child_namespaces,
                    &mut non_visual,
                    &mut group_properties,
                    &mut children,
                    &mut raw_children,
                    &mut state,
                )?;
            }
            Event::End(end) if local_name(end.name().as_ref()) == kind.local_name() => break,
            Event::Eof => {
                return Err(OxmlError::MissingElement(format!("closing {}", kind.tag())));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(ParsedGroup {
        non_visual_group_properties: required(non_visual, "p:nvGrpSpPr")?,
        group_properties: required(group_properties, "p:grpSpPr")?,
        children,
        raw_attributes,
        raw_children,
    })
}

#[allow(clippy::too_many_arguments)]
fn capture_group_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    non_visual: &mut Option<NonVisualGroupProperties>,
    group_properties: &mut Option<GroupProperties>,
    children: &mut Vec<ShapeTreeChild>,
    raw_children: &mut OrderedRawChildren,
    state: &mut usize,
) -> Result<()> {
    if uri == Some(P_NS) && matches!(name, b"nvGrpSpPr" | b"grpSpPr" | b"grpSp") {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
    }
    if *state == 0 {
        if uri == Some(P_NS) && name == b"nvGrpSpPr" {
            *non_visual = Some(NonVisualGroupProperties::from_fragment(&raw, namespaces)?);
            *state = 1;
            return Ok(());
        }
        return Err(OxmlError::InvalidValue(
            "p:nvGrpSpPr must be the first shape-tree element".to_owned(),
        ));
    }
    if *state == 1 {
        if uri == Some(P_NS) && name == b"grpSpPr" {
            *group_properties = Some(GroupProperties::from_fragment(&raw, namespaces)?);
            *state = 2;
            return Ok(());
        }
        return Err(OxmlError::InvalidValue(
            "p:grpSpPr must immediately follow p:nvGrpSpPr".to_owned(),
        ));
    }

    if uri == Some(P_NS) && matches!(name, b"nvGrpSpPr" | b"grpSpPr") {
        return Err(OxmlError::InvalidValue(format!(
            "duplicate p:{} shape-tree element",
            String::from_utf8_lossy(name)
        )));
    }

    if let Some(child) = parse_shape_tree_child(name, uri, &raw, namespaces)? {
        children.push(child);
    } else {
        raw_children.push(children.len(), raw);
    }
    Ok(())
}

fn parse_compatibility_members(
    xml: &[u8],
    inherited: &[(String, String)],
    expected_name: &[u8],
) -> Result<Vec<ShapeTreeChild>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = NamespaceBindings::from_entries(inherited).with_start(&start)?;
                if local_name(start.name().as_ref()) != expected_name
                    || namespaces.element_uri(start.name().as_ref()) != Some(MC_NS)
                {
                    return Err(unexpected(&start));
                }
                return parse_compatibility_reader(&mut reader, &namespaces, expected_name);
            }
            Event::Empty(start) => {
                let namespaces = NamespaceBindings::from_entries(inherited).with_start(&start)?;
                if local_name(start.name().as_ref()) != expected_name
                    || namespaces.element_uri(start.name().as_ref()) != Some(MC_NS)
                {
                    return Err(unexpected(&start));
                }
                return Ok(Vec::new());
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(format!(
                    "mc:{}",
                    String::from_utf8_lossy(expected_name)
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_compatibility_reader(
    reader: &mut Reader<&[u8]>,
    namespaces: &NamespaceBindings,
    expected_name: &[u8],
) -> Result<Vec<ShapeTreeChild>> {
    let mut children = Vec::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_namespaces = namespaces.with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = child_namespaces.element_uri(child.name().as_ref());
                let raw = capture_element(reader, &child)?;
                if let Some(child) = parse_shape_tree_child(&name, uri, &raw, &child_namespaces)? {
                    children.push(child);
                }
            }
            Event::Empty(child) => {
                let child_namespaces = namespaces.with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = child_namespaces.element_uri(child.name().as_ref());
                let raw = capture_empty_element(&child)?;
                if let Some(child) = parse_shape_tree_child(&name, uri, &raw, &child_namespaces)? {
                    children.push(child);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == expected_name => {
                return Ok(children);
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(format!(
                    "closing mc:{}",
                    String::from_utf8_lossy(expected_name)
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_shape_tree_child(
    name: &[u8],
    uri: Option<&str>,
    raw: &[u8],
    namespaces: &NamespaceBindings,
) -> Result<Option<ShapeTreeChild>> {
    let inherited = namespaces.entries();
    let child = match (uri, name) {
        (Some(P_NS), b"sp") => ShapeTreeChild::Shape(CT_Shape::from_fragment(raw, &inherited)?),
        (Some(P_NS), b"pic") => {
            ShapeTreeChild::Picture(CT_Picture::from_fragment(raw, &inherited)?)
        }
        (Some(P_NS), b"graphicFrame") => {
            ShapeTreeChild::GraphicFrame(Box::new(CT_GraphicFrame::from_fragment(raw, &inherited)?))
        }
        (Some(P_NS), b"grpSp") => {
            ShapeTreeChild::GroupShape(Box::new(CT_GroupShape::from_fragment(raw, &inherited)?))
        }
        (Some(P_NS), b"cxnSp") => {
            ShapeTreeChild::Connector(CT_ConnectionShape::from_fragment(raw, &inherited)?)
        }
        (Some(MC_NS), b"AlternateContent") => ShapeTreeChild::AlternateContent(Box::new(
            CT_AlternateContent::from_fragment(raw, &inherited)?,
        )),
        _ => return Ok(None),
    };
    Ok(Some(child))
}

#[allow(clippy::too_many_arguments)]
fn write_group<W: Write>(
    writer: &mut Writer<W>,
    kind: GroupKind,
    non_visual: &NonVisualGroupProperties,
    group_properties: &GroupProperties,
    children: &[ShapeTreeChild],
    attributes: &RawAttributes,
    raw_children: &OrderedRawChildren,
    declare_namespaces: bool,
) -> Result<()> {
    let mut start = BytesStart::new(kind.tag());
    if declare_namespaces {
        start.push_attribute(("xmlns:p", P_NS));
        start.push_attribute(("xmlns:a", A_NS));
        start.push_attribute(("xmlns:r", R_NS));
        start.push_attribute(("xmlns:mc", MC_NS));
    }
    push_attributes(&mut start, attributes);
    writer.write_event(Event::Start(start))?;
    non_visual.write_xml(writer)?;
    group_properties.write_xml(writer)?;
    for (index, child) in children.iter().enumerate() {
        emit_raw(writer, raw_children.at(index))?;
        child.write_xml(writer)?;
    }
    emit_raw(writer, raw_children.at(children.len()))?;
    writer.write_event(Event::End(BytesEnd::new(kind.tag())))?;
    Ok(())
}

impl NonVisualGroupProperties {
    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    let raw_attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?;
                    let mut raw_children = Vec::new();
                    let mut non_visual_id = None;
                    let mut non_visual_name = None;
                    let mut drawing_properties_index = None;
                    loop {
                        buffer.clear();
                        match reader.read_event_into(&mut buffer)? {
                            Event::Start(child) => {
                                let child_namespaces = namespaces.with_start(&child)?;
                                let is_drawing_properties = child_namespaces
                                    .element_uri(child.name().as_ref())
                                    == Some(P_NS)
                                    && local_name(child.name().as_ref()) == b"cNvPr";
                                if is_drawing_properties {
                                    non_visual_name = non_visual_drawing_name(&child)?;
                                }
                                let raw = capture_element(&mut reader, &child)?;
                                if is_drawing_properties {
                                    non_visual_id = non_visual_drawing_id(&raw);
                                    drawing_properties_index = Some(raw_children.len());
                                }
                                raw_children.push(raw);
                            }
                            Event::Empty(child) => {
                                let child_namespaces = namespaces.with_start(&child)?;
                                let is_drawing_properties = child_namespaces
                                    .element_uri(child.name().as_ref())
                                    == Some(P_NS)
                                    && local_name(child.name().as_ref()) == b"cNvPr";
                                if is_drawing_properties {
                                    non_visual_name = non_visual_drawing_name(&child)?;
                                }
                                let raw = capture_empty_element(&child)?;
                                if is_drawing_properties {
                                    non_visual_id = non_visual_drawing_id(&raw);
                                    drawing_properties_index = Some(raw_children.len());
                                }
                                raw_children.push(raw);
                            }
                            Event::End(end) if local_name(end.name().as_ref()) == b"nvGrpSpPr" => {
                                return Ok(Self {
                                    non_visual_id,
                                    non_visual_name,
                                    drawing_properties_index,
                                    raw_attributes,
                                    raw_children,
                                });
                            }
                            Event::Eof => {
                                return Err(OxmlError::MissingElement(
                                    "closing p:nvGrpSpPr".to_owned(),
                                ));
                            }
                            _ => {}
                        }
                    }
                }
                Event::Empty(start) => {
                    return Ok(Self {
                        non_visual_id: None,
                        non_visual_name: None,
                        drawing_properties_index: None,
                        raw_attributes: root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?,
                        raw_children: Vec::new(),
                    });
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:nvGrpSpPr".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:nvGrpSpPr");
        push_attributes(&mut start, &self.raw_attributes);
        if self.raw_children.is_empty() {
            writer.write_event(Event::Empty(start))?;
            return Ok(());
        }
        writer.write_event(Event::Start(start))?;
        for child in &self.raw_children {
            writer.get_mut().write_all(child)?;
        }
        writer.write_event(Event::End(BytesEnd::new("p:nvGrpSpPr")))?;
        Ok(())
    }
}

impl GroupProperties {
    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    return Self::from_reader(&mut reader, &start, &namespaces);
                }
                Event::Empty(start) => {
                    return Ok(Self {
                        transform: None,
                        raw_attributes: root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?,
                        raw_children: OrderedRawChildren::default(),
                    });
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:grpSpPr".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_reader(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        namespaces: &NamespaceBindings,
    ) -> Result<Self> {
        let mut transform = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0usize;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let is_transform = child_namespaces.element_uri(child.name().as_ref())
                        == Some(A_NS)
                        && local_name(child.name().as_ref()) == b"xfrm";
                    if is_transform {
                        child_namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
                    }
                    let raw = capture_element(reader, &child)?;
                    capture_group_property_child(
                        is_transform,
                        raw,
                        &mut transform,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::Empty(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let is_transform = child_namespaces.element_uri(child.name().as_ref())
                        == Some(A_NS)
                        && local_name(child.name().as_ref()) == b"xfrm";
                    if is_transform {
                        child_namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
                    }
                    let raw = capture_empty_element(&child)?;
                    capture_group_property_child(
                        is_transform,
                        raw,
                        &mut transform,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"grpSpPr" => break,
                Event::Eof => {
                    return Err(OxmlError::MissingElement("closing p:grpSpPr".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
        Ok(Self {
            transform,
            raw_attributes: root_attributes(start, FIXED_SHAPE_TREE_PREFIXES)?,
            raw_children,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:grpSpPr");
        push_attributes(&mut start, &self.raw_attributes);
        if self.transform.is_none() && self.raw_children.is_empty() {
            writer.write_event(Event::Empty(start))?;
            return Ok(());
        }
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(transform) = &self.transform {
            transform
                .write_xml(writer)
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        writer.write_event(Event::End(BytesEnd::new("p:grpSpPr")))?;
        Ok(())
    }
}

fn capture_group_property_child(
    is_transform: bool,
    raw: Vec<u8>,
    transform: &mut Option<CT_Transform2D>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if is_transform {
        if transform.is_some() {
            return Err(OxmlError::InvalidValue(
                "duplicate DrawingML group transform".to_owned(),
            ));
        }
        if !raw_children.is_empty() {
            return Err(OxmlError::InvalidValue(
                "a:xfrm must precede other p:grpSpPr children".to_owned(),
            ));
        }
        *transform = Some(
            CT_Transform2D::from_xml(&raw)
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
        );
        *boundary = 1;
    } else {
        raw_children.push(*boundary, raw);
    }
    Ok(())
}

fn push_attributes(start: &mut BytesStart<'_>, attributes: &RawAttributes) {
    for (name, value) in attributes {
        start.push_attribute((name.as_str(), value.as_str()));
    }
}

fn emit_raw<'a, W: Write>(
    writer: &mut Writer<W>,
    children: impl Iterator<Item = &'a [u8]>,
) -> Result<()> {
    for child in children {
        writer.get_mut().write_all(child)?;
    }
    Ok(())
}

fn required<T>(value: Option<T>, name: &str) -> Result<T> {
    value.ok_or_else(|| OxmlError::MissingElement(name.to_owned()))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn unexpected(element: &BytesStart<'_>) -> OxmlError {
    OxmlError::UnexpectedElement(String::from_utf8_lossy(element.name().as_ref()).into_owned())
}

#[cfg(test)]
mod style_tests {
    use oxml_core::units::Emu;
    use oxml_drawing::style_ref::FontCollectionIndex;
    use oxml_drawing::xfrm::{CT_Point2D, CT_PositiveSize2D, CT_Transform2D};

    use super::{CT_GroupShape, CT_Shape, CT_ShapeTree, ShapeIdAllocator, ShapeTreeChild};

    const ALLOCATOR_TREE: &[u8] = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Root shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="6" name="Nested picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:grpSp><mc:AlternateContent><mc:Fallback><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="8" name="Fallback connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></mc:Fallback></mc:AlternateContent></p:spTree>"#;

    #[test]
    fn ordinary_shape_and_textbox_constructors_emit_canonical_shells() {
        let mut transform = CT_Transform2D::default();
        transform.offset = Some(CT_Point2D {
            x: Emu(10),
            y: Emu(20),
        });
        transform.extent = Some(CT_PositiveSize2D {
            cx: Emu(30),
            cy: Emu(40),
        });
        let shape = CT_Shape::new_preset(2, "Shape & 2", "triangle", transform.clone()).unwrap();
        let textbox = CT_Shape::new_textbox(3, "TextBox 3", transform).unwrap();

        let shape_xml = String::from_utf8(shape.to_xml().unwrap()).unwrap();
        assert!(shape_xml.contains("<p:cNvPr id=\"2\" name=\"Shape &amp; 2\"/>"));
        assert!(shape_xml.contains("<p:cNvSpPr/>"));
        assert!(shape_xml.contains("<a:prstGeom prst=\"triangle\"><a:avLst/></a:prstGeom>"));
        assert!(shape_xml.contains("<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"));
        assert!(shape_xml.find("<p:nvSpPr").unwrap() < shape_xml.find("<p:spPr").unwrap());
        assert!(shape_xml.find("<p:spPr").unwrap() < shape_xml.find("<p:txBody").unwrap());

        let textbox_xml = String::from_utf8(textbox.to_xml().unwrap()).unwrap();
        assert!(textbox_xml.contains("<p:cNvSpPr txBox=\"1\"/>"));
        assert!(textbox_xml.contains("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"));
        assert!(textbox_xml.contains("<a:noFill/>"));
        assert!(textbox_xml.contains("<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"));
        assert_eq!(CT_Shape::from_xml(textbox_xml.as_bytes()).unwrap(), textbox);
    }

    #[test]
    fn empty_group_constructor_has_required_children() {
        let group = CT_GroupShape::new_empty(9, "Group 9");
        let xml = String::from_utf8(group.to_xml().unwrap()).unwrap();
        assert!(xml.contains("<p:nvGrpSpPr><p:cNvPr id=\"9\" name=\"Group 9\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"));
        assert!(xml.contains("<p:grpSpPr/>"));
        assert!(xml.find("<p:nvGrpSpPr").unwrap() < xml.find("<p:grpSpPr").unwrap());
        assert!(group.children.is_empty());
        assert!(group.group_transform().is_none());
        assert_eq!(CT_GroupShape::from_xml(xml.as_bytes()).unwrap(), group);
    }

    #[test]
    fn shape_id_allocator_scans_nested_groups_and_alternate_content() {
        let tree = CT_ShapeTree::from_xml(ALLOCATOR_TREE).unwrap();
        let ids = tree
            .children
            .iter()
            .filter_map(|child| child.non_visual_id())
            .collect::<Vec<_>>();
        assert_eq!(ids, [2, 4]);

        let mut allocator = ShapeIdAllocator::scan(&tree);
        assert_eq!(allocator.allocate(), 3);
        assert_eq!(allocator.allocate(), 5);
        assert_eq!(allocator.allocate(), 7);
        assert_eq!(allocator.allocate(), 9);
    }

    #[test]
    fn shape_tree_child_names_are_decoded_once_and_follow_mutation() {
        let decoded = CT_Shape::from_xml(
            br#"<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main"><q:nvSpPr><q:cNvPr id="9" name="!!Hero &amp; Partner"/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr/></q:sp>"#,
        )
        .unwrap();
        assert_eq!(
            ShapeTreeChild::Shape(decoded).non_visual_name().as_deref(),
            Some("!!Hero & Partner")
        );

        let mut tree = CT_ShapeTree::from_xml(ALLOCATOR_TREE).unwrap();
        assert_eq!(
            tree.children[0].non_visual_name().as_deref(),
            Some("Root shape")
        );
        assert_eq!(tree.children[1].non_visual_name().as_deref(), Some("Group"));
        let ShapeTreeChild::GroupShape(group) = &tree.children[1] else {
            panic!("expected group")
        };
        assert_eq!(
            group.children[0].non_visual_name().as_deref(),
            Some("Nested picture")
        );

        let ShapeTreeChild::Shape(shape) = &mut tree.children[0] else {
            panic!("expected shape")
        };
        shape.set_name("!!Hero & Partner").unwrap();
        assert_eq!(
            tree.children[0].non_visual_name().as_deref(),
            Some("!!Hero & Partner")
        );
    }

    #[test]
    fn shape_id_allocator_starts_at_two_and_skips_sparse_ids() {
        let xml = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="3" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Duplicate"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:spTree>"#;
        let tree = CT_ShapeTree::from_xml(xml).unwrap();
        let mut allocator = ShapeIdAllocator::scan(&tree);

        assert_eq!(allocator.allocate(), 2);
        assert_eq!(allocator.allocate(), 4);
        assert_eq!(allocator.allocate(), 5);
    }

    #[test]
    fn shape_id_allocator_scans_raw_members_and_non_selected_choices() {
        let xml = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><x:producer xmlns:x="urn:producer" xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main"><q:cNvPr id="2"/></x:producer><mc:AlternateContent><mc:Choice Requires="p14"><r:sp xmlns:r="http://schemas.openxmlformats.org/presentationml/2006/main"><r:nvSpPr><r:cNvPr id="3"/><r:cNvSpPr/><r:nvPr/></r:nvSpPr><r:spPr/></r:sp></mc:Choice><mc:Fallback/></mc:AlternateContent></p:spTree>"#;
        let tree = CT_ShapeTree::from_xml(xml).unwrap();
        let mut allocator = ShapeIdAllocator::scan(&tree);

        assert_eq!(allocator.allocate(), 4);
        assert_eq!(allocator.allocate(), 5);
    }

    #[test]
    fn typed_non_visual_ids_preserve_original_shape_xml() {
        let xml = br#"<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:test"><q:nvSpPr><q:cNvPr id="7" name="Shape"><x:producer-data x:value="kept"/></q:cNvPr><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr/></q:sp>"#;
        let shape = CT_Shape::from_xml(xml).unwrap();

        assert_eq!(shape.non_visual_id(), Some(7));
        let written = shape.to_xml().unwrap();
        let text = String::from_utf8(written.clone()).unwrap();
        assert!(text.starts_with("<p:sp "));
        assert!(text.contains("<q:cNvPr id=\"7\" name=\"Shape\">"));
        assert!(text.contains("<x:producer-data x:value=\"kept\"/>"));
        assert!(text.find("<q:cNvPr").unwrap() < text.find("<q:cNvSpPr").unwrap());
        assert!(text.find("<q:cNvSpPr").unwrap() < text.find("<p:nvPr").unwrap());
        assert_eq!(CT_Shape::from_xml(&written).unwrap(), shape);
    }

    #[test]
    fn group_name_mutation_uses_the_namespace_resolved_drawing_properties() {
        let xml = br#"<q:grpSp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:test"><q:nvGrpSpPr><x:cNvPr id="99" name="extension"/><q:cNvPr id="4" name="Group"><x:raw/></q:cNvPr><q:cNvGrpSpPr/><q:nvPr/></q:nvGrpSpPr><q:grpSpPr/></q:grpSp>"#;
        let mut group = CT_GroupShape::from_xml(xml).unwrap();

        group.set_name("Changed & safe").unwrap();
        let written = String::from_utf8(group.to_xml().unwrap()).unwrap();

        assert!(written.contains("<x:cNvPr id=\"99\" name=\"extension\"/>"));
        assert!(
            written.contains("<q:cNvPr id=\"4\" name=\"Changed &amp; safe\"><x:raw/></q:cNvPr>")
        );
    }

    #[test]
    fn ordinary_shape_style_round_trips_in_schema_order() {
        let xml = br#"<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:test"><q:nvSpPr><q:cNvPr/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr/><x:before-style x:value="kept"/><q:style x:custom="yes"><x:before/><d:lnRef idx="2"><d:schemeClr val="accent1"/></d:lnRef><x:between-line-fill/><d:fillRef idx="1"><d:schemeClr val="accent2"/></d:fillRef><d:effectRef idx="3"><d:schemeClr val="accent3"/></d:effectRef><d:fontRef idx="major"><d:schemeClr val="tx1"/></d:fontRef><x:after/></q:style><x:after-style/></q:sp>"#;
        let shape = CT_Shape::from_xml(xml).unwrap();
        let style = shape.style().unwrap();

        assert_eq!(style.line_reference.index, 2);
        assert_eq!(style.fill_reference.index, 1);
        assert_eq!(style.effect_reference.index, 3);
        assert_eq!(style.font_reference.index, FontCollectionIndex::Major);

        let written = shape.to_xml().unwrap();
        let text = String::from_utf8(written.clone()).unwrap();
        assert!(text.contains("<p:style x:custom=\"yes\">"));
        assert!(text.contains("<x:before/>"));
        assert!(text.contains("<x:between-line-fill/>"));
        assert!(text.contains("<x:after/>"));
        let line = text.find("<a:lnRef").unwrap();
        let fill = text.find("<a:fillRef").unwrap();
        let effect = text.find("<a:effectRef").unwrap();
        let font = text.find("<a:fontRef").unwrap();
        assert!(line < fill && fill < effect && effect < font);
        assert!(text.find("<p:spPr").unwrap() < text.find("<p:style").unwrap());
        assert!(text.find("<p:style").unwrap() < text.find("<x:after-style").unwrap());
        assert_eq!(CT_Shape::from_xml(&written).unwrap(), shape);
    }
}
