use std::collections::HashSet;
use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_drawing::namespace::A_NS;
use oxml_drawing::order::OrderedRawChildren;
use oxml_drawing::table::CT_Table;
use oxml_drawing::xfrm::CT_Transform2D;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::diagram::DiagramRelationshipIds;
use crate::namespace::{
    FIXED_SHAPE_TREE_PREFIXES, NamespaceBindings, P_NS, all_attributes, non_visual_drawing_id,
    non_visual_drawing_name, root_attributes, self_contained_attributes,
    set_non_visual_drawing_name,
};
use crate::picture::CT_Picture;

const TABLE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
const CHART_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DIAGRAM_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const OLE_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";
const GRAPHIC_FRAME_PREFIXES: &[&str] = &["p", "a"];

pub type Result<T> = std::result::Result<T, OxmlError>;
type RawAttributes = Vec<(String, String)>;

/// The payload selected by `a:graphicData@uri`.
#[derive(Clone, Debug, PartialEq)]
pub enum GraphicDataPayload {
    Table(Box<CT_Table>),
    Chart(Vec<u8>),
    SmartArt(Box<DiagramRelationshipIds>),
    Ole {
        raw: Vec<u8>,
        preview: Option<Box<CT_Picture>>,
    },
    Other(Vec<u8>),
}

impl GraphicDataPayload {
    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Table(table) => writer.get_mut().write_all(&table.to_xml()?)?,
            Self::Chart(xml) | Self::Other(xml) => writer.get_mut().write_all(xml)?,
            Self::SmartArt(relationships) => writer.get_mut().write_all(&relationships.to_xml())?,
            Self::Ole { raw, .. } => writer.get_mut().write_all(raw)?,
        }
        Ok(())
    }
}

/// One URI-dispatched `a:graphicData` payload.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_GraphicData {
    pub uri: String,
    payload: GraphicDataPayload,
    chart_relationship_id: Option<String>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

impl CT_GraphicData {
    /// Returns the URI-dispatched payload as a read-only serialization source.
    pub fn payload(&self) -> &GraphicDataPayload {
        &self.payload
    }

    /// Returns the mutable typed table payload, when this graphic is a table.
    pub fn table_mut(&mut self) -> Option<&mut CT_Table> {
        match &mut self.payload {
            GraphicDataPayload::Table(table) => Some(table),
            _ => None,
        }
    }

    /// Returns whether an OLE payload embeds its object rather than linking it.
    ///
    /// Like python-pptx, the last `p:oleObj` decides, because alternate
    /// content lists its lowest common denominator last. `None` means the
    /// payload is not OLE.
    pub fn is_embedded_ole_object(&self) -> Option<bool> {
        let GraphicDataPayload::Ole { raw, .. } = &self.payload else {
            return None;
        };
        let mut reader = Reader::from_reader(raw.as_slice());
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut object_depth = None;
        let mut embedded = false;
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start)) => {
                    depth += 1;
                    match local_name(start.name().as_ref()) {
                        b"oleObj" => {
                            object_depth = Some(depth);
                            embedded = false;
                        }
                        b"embed" if object_depth == Some(depth - 1) => embedded = true,
                        _ => {}
                    }
                }
                Ok(Event::Empty(start)) => match local_name(start.name().as_ref()) {
                    b"oleObj" => {
                        object_depth = None;
                        embedded = false;
                    }
                    b"embed" if object_depth == Some(depth) => embedded = true,
                    _ => {}
                },
                Ok(Event::End(_)) => {
                    if object_depth == Some(depth) {
                        object_depth = None;
                    }
                    depth = depth.saturating_sub(1);
                }
                Ok(Event::Eof) | Err(_) => return Some(embedded),
                _ => {}
            }
            buffer.clear();
        }
    }
}

/// One typed PresentationML `p:graphicFrame`.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_GraphicFrame {
    pub transform: CT_Transform2D,
    pub graphic_data: CT_GraphicData,
    non_visual_properties: RawElementShell,
    graphic_attributes: RawAttributes,
    graphic_raw_children: OrderedRawChildren,
    extension_xml: Option<Vec<u8>>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct RawElementShell {
    non_visual_id: Option<u32>,
    non_visual_name: Option<String>,
    drawing_properties_index: Option<usize>,
    attributes: RawAttributes,
    children: Vec<Vec<u8>>,
}

struct ParsedGraphic {
    graphic_data: CT_GraphicData,
    attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

impl CT_GraphicFrame {
    /// Creates a table-bearing graphic frame with canonical non-visual shells.
    pub fn new_table(
        id: u32,
        name: &str,
        transform: CT_Transform2D,
        table: CT_Table,
    ) -> Result<Self> {
        table.to_xml()?;
        let non_visual_name = name.to_owned();
        let escaped_name = quick_xml::escape::escape(name);
        Ok(Self {
            transform,
            graphic_data: CT_GraphicData {
                uri: TABLE_URI.to_owned(),
                payload: GraphicDataPayload::Table(Box::new(table)),
                chart_relationship_id: None,
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            non_visual_properties: RawElementShell {
                non_visual_id: Some(id),
                non_visual_name: Some(non_visual_name),
                drawing_properties_index: Some(0),
                attributes: Vec::new(),
                children: vec![
                    format!(r#"<p:cNvPr id="{id}" name="{escaped_name}"/>"#).into_bytes(),
                    b"<p:cNvGraphicFramePr/>".to_vec(),
                    b"<p:nvPr/>".to_vec(),
                ],
            },
            graphic_attributes: Vec::new(),
            graphic_raw_children: OrderedRawChildren::default(),
            extension_xml: None,
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        })
    }

    /// Creates a chart-bearing graphic frame with a relationship payload.
    pub fn new_chart(
        id: u32,
        name: &str,
        transform: CT_Transform2D,
        relationship_id: &str,
    ) -> Result<Self> {
        if relationship_id.is_empty() {
            return Err(OxmlError::InvalidValue(
                "a chart graphic frame requires a relationship id".to_owned(),
            ));
        }
        let mut writer = Writer::new(Vec::new());
        let mut chart = BytesStart::new("c:chart");
        chart.push_attribute(("xmlns:c", CHART_URI));
        chart.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        chart.push_attribute(("r:id", relationship_id));
        writer.write_event(Event::Empty(chart))?;
        let non_visual_name = name.to_owned();
        let escaped_name = quick_xml::escape::escape(name);
        Ok(Self {
            transform,
            graphic_data: CT_GraphicData {
                uri: CHART_URI.to_owned(),
                payload: GraphicDataPayload::Chart(writer.into_inner()),
                chart_relationship_id: Some(relationship_id.to_owned()),
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            non_visual_properties: RawElementShell {
                non_visual_id: Some(id),
                non_visual_name: Some(non_visual_name),
                drawing_properties_index: Some(0),
                attributes: Vec::new(),
                children: vec![
                    format!(r#"<p:cNvPr id="{id}" name="{escaped_name}"/>"#).into_bytes(),
                    b"<p:cNvGraphicFramePr/>".to_vec(),
                    b"<p:nvPr/>".to_vec(),
                ],
            },
            graphic_attributes: Vec::new(),
            graphic_raw_children: OrderedRawChildren::default(),
            extension_xml: None,
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        })
    }

    pub(crate) fn non_visual_id(&self) -> Option<u32> {
        self.non_visual_properties.non_visual_id
    }

    pub(crate) fn non_visual_name(&self) -> Option<&str> {
        self.non_visual_properties.non_visual_name.as_deref()
    }

    /// Returns the namespace-resolved `c:chart@r:id` rendering projection.
    pub fn chart_relationship_id(&self) -> Option<&str> {
        self.graphic_data.chart_relationship_id.as_deref()
    }

    /// Changes the producer-facing non-visual frame name.
    pub fn set_name(&mut self, name: &str) -> Result<()> {
        let drawing_properties = self
            .non_visual_properties
            .children
            .get_mut(
                self.non_visual_properties
                    .drawing_properties_index
                    .ok_or_else(|| OxmlError::MissingElement("p:cNvPr".to_owned()))?,
            )
            .ok_or_else(|| OxmlError::MissingElement("p:cNvPr".to_owned()))?;
        set_non_visual_drawing_name(drawing_properties, name)?;
        self.non_visual_properties.non_visual_name = Some(name.to_owned());
        Ok(())
    }

    /// Parses a complete `p:graphicFrame` with any PresentationML prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_fragment(xml, &[])
    }

    pub(crate) fn from_fragment(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let inherited = referenced_namespaces(xml, inherited)?;
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces =
                        NamespaceBindings::from_entries(&inherited).with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"graphicFrame"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
                    return Self::from_reader(&mut reader, &start, &namespaces, &inherited);
                }
                Event::Empty(start) => {
                    return Err(OxmlError::MissingElement(format!(
                        "{} requires p:nvGraphicFramePr, p:xfrm, and a:graphic",
                        String::from_utf8_lossy(start.name().as_ref())
                    )));
                }
                Event::Eof => {
                    return Err(OxmlError::MissingElement("p:graphicFrame".to_owned()));
                }
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
        let mut transform = None;
        let mut graphic = None;
        let mut extension_xml = None;
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
                    capture_frame_child(
                        &name,
                        uri,
                        raw,
                        &child_namespaces,
                        &mut non_visual,
                        &mut transform,
                        &mut graphic,
                        &mut extension_xml,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::Empty(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_empty_element(&child)?;
                    capture_frame_child(
                        &name,
                        uri,
                        raw,
                        &child_namespaces,
                        &mut non_visual,
                        &mut transform,
                        &mut graphic,
                        &mut extension_xml,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"graphicFrame" => break,
                Event::Eof => {
                    return Err(OxmlError::MissingElement(
                        "closing p:graphicFrame".to_owned(),
                    ));
                }
                event => {
                    if let Some(raw) = capture_non_element(event)? {
                        raw_children.push(boundary, raw);
                    }
                }
            }
            buffer.clear();
        }

        let graphic = required(graphic, "a:graphic")?;
        Ok(Self {
            transform: required(transform, "p:xfrm")?,
            graphic_data: graphic.graphic_data,
            non_visual_properties: required(non_visual, "p:nvGraphicFramePr")?,
            graphic_attributes: graphic.attributes,
            graphic_raw_children: graphic.raw_children,
            extension_xml,
            raw_attributes: self_contained_attributes(start, GRAPHIC_FRAME_PREFIXES, inherited)?,
            raw_children,
        })
    }

    /// Serialises a self-contained graphic frame with fixed modelled prefixes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml_internal(&mut writer, true)?;
        Ok(writer.into_inner())
    }

    pub(crate) fn write_xml_internal<W: Write>(
        &self,
        writer: &mut Writer<W>,
        declare_namespaces: bool,
    ) -> Result<()> {
        let mut start = BytesStart::new("p:graphicFrame");
        if declare_namespaces {
            start.push_attribute(("xmlns:p", P_NS));
            start.push_attribute(("xmlns:a", A_NS));
        }
        push_attributes(&mut start, &self.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        self.non_visual_properties
            .write_xml(writer, "p:nvGraphicFramePr")?;
        emit_raw(writer, self.raw_children.at(1))?;
        self.transform
            .write_xml_with_root(writer, "p:xfrm")
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        emit_raw(writer, self.raw_children.at(2))?;
        self.write_graphic(writer)?;
        emit_raw(writer, self.raw_children.at(3))?;
        if let Some(extension) = &self.extension_xml {
            writer.get_mut().write_all(extension)?;
        }
        emit_raw(writer, self.raw_children.at(4))?;
        writer.write_event(Event::End(BytesEnd::new("p:graphicFrame")))?;
        Ok(())
    }

    fn write_graphic<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:graphic");
        push_attributes(&mut start, &self.graphic_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.graphic_raw_children.at(0))?;
        self.graphic_data.write_xml(writer)?;
        emit_raw(writer, self.graphic_raw_children.at(1))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn capture_frame_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    non_visual: &mut Option<RawElementShell>,
    transform: &mut Option<CT_Transform2D>,
    graphic: &mut Option<ParsedGraphic>,
    extension_xml: &mut Option<Vec<u8>>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if matches!(
        (uri, name),
        (Some(P_NS), b"nvGraphicFramePr" | b"xfrm" | b"extLst") | (Some(A_NS), b"graphic")
    ) {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
    }
    match (uri, name) {
        (Some(P_NS), b"nvGraphicFramePr") => {
            if *boundary != 0 || non_visual.is_some() {
                return Err(order_error(
                    "p:nvGraphicFramePr must be the first modelled graphic-frame child",
                ));
            }
            *non_visual = Some(RawElementShell::from_fragment(&raw, namespaces)?);
            *boundary = 1;
        }
        (Some(P_NS), b"xfrm") => {
            if *boundary != 1 || transform.is_some() {
                return Err(order_error(
                    "p:xfrm must immediately follow p:nvGraphicFramePr",
                ));
            }
            *transform = Some(
                CT_Transform2D::from_xml(&raw)
                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
            );
            *boundary = 2;
        }
        (Some(A_NS), b"graphic") => {
            if *boundary != 2 || graphic.is_some() {
                return Err(order_error("a:graphic must immediately follow p:xfrm"));
            }
            *graphic = Some(parse_graphic(&raw, namespaces)?);
            *boundary = 3;
        }
        (Some(P_NS), b"extLst") => {
            if *boundary != 3 || extension_xml.replace(raw).is_some() {
                return Err(order_error(
                    "p:extLst must follow a:graphic and occur at most once",
                ));
            }
            *boundary = 4;
        }
        _ => raw_children.push(*boundary, raw),
    }
    Ok(())
}

fn parse_graphic(xml: &[u8], inherited: &NamespaceBindings) -> Result<ParsedGraphic> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                let attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?;
                let mut graphic_data = None;
                let mut raw_children = OrderedRawChildren::default();
                let mut boundary = 0usize;
                loop {
                    buffer.clear();
                    match reader.read_event_into(&mut buffer)? {
                        Event::Start(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            let uri = child_namespaces.element_uri(child.name().as_ref());
                            let name = local_name(child.name().as_ref()).to_vec();
                            let raw = capture_element(&mut reader, &child)?;
                            capture_graphic_child(
                                &name,
                                uri,
                                raw,
                                &child_namespaces,
                                &mut graphic_data,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::Empty(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            let uri = child_namespaces.element_uri(child.name().as_ref());
                            let name = local_name(child.name().as_ref()).to_vec();
                            let raw = capture_empty_element(&child)?;
                            capture_graphic_child(
                                &name,
                                uri,
                                raw,
                                &child_namespaces,
                                &mut graphic_data,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::End(end) if local_name(end.name().as_ref()) == b"graphic" => break,
                        Event::Eof => {
                            return Err(OxmlError::MissingElement("closing a:graphic".to_owned()));
                        }
                        event => {
                            if let Some(raw) = capture_non_element(event)? {
                                raw_children.push(boundary, raw);
                            }
                        }
                    }
                }
                return Ok(ParsedGraphic {
                    graphic_data: required(graphic_data, "a:graphicData")?,
                    attributes,
                    raw_children,
                });
            }
            Event::Empty(_) => {
                return Err(OxmlError::MissingElement("a:graphicData".to_owned()));
            }
            Event::Eof => return Err(OxmlError::MissingElement("a:graphic".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn capture_graphic_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    graphic_data: &mut Option<CT_GraphicData>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if uri == Some(A_NS) && name == b"graphicData" {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
        if graphic_data.is_some() {
            return Err(order_error("duplicate a:graphicData"));
        }
        *graphic_data = Some(CT_GraphicData::from_fragment(&raw, namespaces)?);
        *boundary = 1;
    } else {
        raw_children.push(*boundary, raw);
    }
    Ok(())
}

impl CT_GraphicData {
    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    let attributes = all_attributes(&start)?;
                    let uri = attributes
                        .iter()
                        .find(|(name, _)| name == "uri")
                        .map(|(_, value)| value.clone())
                        .ok_or_else(|| OxmlError::MissingElement("a:graphicData@uri".to_owned()))?;
                    let raw_attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?
                        .into_iter()
                        .filter(|(name, _)| name != "uri")
                        .collect();
                    let mut payload = None;
                    let mut chart_relationship_id = None;
                    let mut raw_children = OrderedRawChildren::default();
                    let mut buffer = Vec::new();
                    loop {
                        match reader.read_event_into(&mut buffer)? {
                            Event::Start(child) => {
                                let child_namespaces = namespaces.with_start(&child)?;
                                let raw = capture_element(&mut reader, &child)?;
                                capture_payload(
                                    &uri,
                                    raw,
                                    &child_namespaces,
                                    &mut payload,
                                    &mut chart_relationship_id,
                                    &mut raw_children,
                                )?;
                            }
                            Event::Empty(child) => {
                                let child_namespaces = namespaces.with_start(&child)?;
                                let raw = capture_empty_element(&child)?;
                                capture_payload(
                                    &uri,
                                    raw,
                                    &child_namespaces,
                                    &mut payload,
                                    &mut chart_relationship_id,
                                    &mut raw_children,
                                )?;
                            }
                            Event::End(end)
                                if local_name(end.name().as_ref()) == b"graphicData" =>
                            {
                                break;
                            }
                            Event::Eof => {
                                return Err(OxmlError::MissingElement(
                                    "closing a:graphicData".to_owned(),
                                ));
                            }
                            event => {
                                if let Some(raw) = capture_non_element(event)? {
                                    raw_children.push(usize::from(payload.is_some()), raw);
                                }
                            }
                        }
                        buffer.clear();
                    }
                    return Ok(Self {
                        uri,
                        payload: required(payload, "a:graphicData payload")?,
                        chart_relationship_id,
                        raw_attributes,
                        raw_children,
                    });
                }
                Event::Empty(_) => {
                    return Err(OxmlError::MissingElement(
                        "a:graphicData payload".to_owned(),
                    ));
                }
                Event::Eof => return Err(OxmlError::MissingElement("a:graphicData".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:graphicData");
        start.push_attribute(("uri", self.uri.as_str()));
        push_attributes(&mut start, &self.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        self.payload.write_xml(writer)?;
        emit_raw(writer, self.raw_children.at(1))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
        Ok(())
    }
}

fn capture_payload(
    uri: &str,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    payload: &mut Option<GraphicDataPayload>,
    chart_relationship_id: &mut Option<String>,
    raw_children: &mut OrderedRawChildren,
) -> Result<()> {
    if payload.is_some() {
        raw_children.push(1, raw);
        return Ok(());
    }
    *payload = Some(match uri {
        TABLE_URI => {
            let inherited = referenced_namespaces(&raw, &namespaces.entries())?;
            GraphicDataPayload::Table(Box::new(CT_Table::from_xml_with_inherited_namespaces(
                &raw, &inherited,
            )?))
        }
        CHART_URI => {
            *chart_relationship_id = parse_chart_relationship_id(&raw, namespaces)?;
            GraphicDataPayload::Chart(raw)
        }
        DIAGRAM_URI => GraphicDataPayload::SmartArt(Box::new(
            DiagramRelationshipIds::from_xml_with_namespaces(&raw, &namespaces.entries())?,
        )),
        OLE_URI => GraphicDataPayload::Ole {
            preview: parse_ole_preview(&raw, &namespaces.entries())
                .ok()
                .flatten()
                .map(Box::new),
            raw,
        },
        _ => GraphicDataPayload::Other(raw),
    });
    Ok(())
}

fn parse_chart_relationship_id(
    raw: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) | Event::Empty(start) => {
                let namespaces = inherited.with_start(&start)?;
                if local_name(start.name().as_ref()) != b"chart"
                    || namespaces.element_uri(start.name().as_ref()) != Some(CHART_URI)
                {
                    return Ok(None);
                }
                for attribute in start.attributes() {
                    let attribute = attribute?;
                    if local_name(attribute.key.as_ref()) == b"id"
                        && namespaces.attribute_uri(attribute.key.as_ref())
                            == Some(crate::namespace::R_NS)
                    {
                        let value = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                start.decoder(),
                            )?
                            .into_owned();
                        return Ok((!value.is_empty()).then_some(value));
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

fn parse_ole_preview(raw: &[u8], inherited: &[(String, String)]) -> Result<Option<CT_Picture>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if local_name(start.name().as_ref()) == b"pic" => {
                let picture = capture_element(&mut reader, &start)?;
                return CT_Picture::from_fragment(&picture, inherited).map(Some);
            }
            Event::Empty(start) if local_name(start.name().as_ref()) == b"pic" => {
                return Ok(None);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

impl RawElementShell {
    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    let attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?;
                    let mut children = Vec::new();
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
                                    drawing_properties_index = Some(children.len());
                                }
                                children.push(raw);
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
                                    drawing_properties_index = Some(children.len());
                                }
                                children.push(raw);
                            }
                            Event::End(end)
                                if local_name(end.name().as_ref()) == b"nvGraphicFramePr" =>
                            {
                                return Ok(Self {
                                    non_visual_id,
                                    non_visual_name,
                                    drawing_properties_index,
                                    attributes,
                                    children,
                                });
                            }
                            Event::Eof => {
                                return Err(OxmlError::MissingElement(
                                    "closing p:nvGraphicFramePr".to_owned(),
                                ));
                            }
                            event => {
                                if let Some(raw) = capture_non_element(event)? {
                                    children.push(raw);
                                }
                            }
                        }
                    }
                }
                Event::Empty(start) => {
                    return Ok(Self {
                        non_visual_id: None,
                        non_visual_name: None,
                        drawing_properties_index: None,
                        attributes: root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?,
                        children: Vec::new(),
                    });
                }
                Event::Eof => {
                    return Err(OxmlError::MissingElement("p:nvGraphicFramePr".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>, name: &str) -> Result<()> {
        let mut start = BytesStart::new(name);
        push_attributes(&mut start, &self.attributes);
        if self.children.is_empty() {
            writer.write_event(Event::Empty(start))?;
            return Ok(());
        }
        writer.write_event(Event::Start(start))?;
        for child in &self.children {
            writer.get_mut().write_all(child)?;
        }
        writer.write_event(Event::End(BytesEnd::new(name)))?;
        Ok(())
    }
}

fn push_attributes(start: &mut BytesStart<'_>, attributes: &RawAttributes) {
    for (name, value) in attributes {
        start.push_attribute((name.as_str(), value.as_str()));
    }
}

fn referenced_namespaces(
    xml: &[u8],
    inherited: &[(String, String)],
) -> Result<Vec<(String, String)>> {
    let inherited_prefixes = inherited
        .iter()
        .map(|(prefix, _)| prefix.clone())
        .collect::<HashSet<_>>();
    let mut referenced = HashSet::new();
    let mut local_scopes = Vec::new();
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                record_element_namespaces(
                    &start,
                    &inherited_prefixes,
                    &mut local_scopes,
                    &mut referenced,
                )?;
            }
            Event::Empty(start) => {
                record_element_namespaces(
                    &start,
                    &inherited_prefixes,
                    &mut local_scopes,
                    &mut referenced,
                )?;
                local_scopes.pop();
            }
            Event::End(_) => {
                local_scopes.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(inherited
        .iter()
        .filter(|(prefix, _)| referenced.contains(prefix))
        .cloned()
        .collect())
}

fn record_element_namespaces(
    start: &BytesStart<'_>,
    inherited_prefixes: &HashSet<String>,
    local_scopes: &mut Vec<HashSet<String>>,
    referenced: &mut HashSet<String>,
) -> Result<()> {
    let mut declarations = HashSet::new();
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref().to_vec();
        if name == b"xmlns" {
            declarations.insert(String::new());
            continue;
        }
        if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            declarations.insert(std::str::from_utf8(prefix)?.to_owned());
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        attributes.push((name, value));
    }
    local_scopes.push(declarations);
    record_name_prefix(
        start.name().as_ref(),
        true,
        inherited_prefixes,
        local_scopes,
        referenced,
    )?;
    for (name, value) in attributes {
        record_name_prefix(&name, false, inherited_prefixes, local_scopes, referenced)?;
        record_value_prefixes(&name, &value, inherited_prefixes, local_scopes, referenced);
    }
    Ok(())
}

fn record_name_prefix(
    name: &[u8],
    use_default: bool,
    inherited: &HashSet<String>,
    local_scopes: &[HashSet<String>],
    referenced: &mut HashSet<String>,
) -> Result<()> {
    let prefix = if let Some(position) = name.iter().position(|byte| *byte == b':') {
        std::str::from_utf8(&name[..position])?
    } else if use_default {
        ""
    } else {
        return Ok(());
    };
    if inherited.contains(prefix) && !prefix_is_locally_bound(prefix, local_scopes) {
        referenced.insert(prefix.to_owned());
    }
    Ok(())
}

fn prefix_is_locally_bound(prefix: &str, local_scopes: &[HashSet<String>]) -> bool {
    local_scopes
        .iter()
        .rev()
        .any(|scope| scope.contains(prefix))
}

fn record_value_prefixes(
    attribute_name: &[u8],
    value: &str,
    inherited: &HashSet<String>,
    local_scopes: &[HashSet<String>],
    referenced: &mut HashSet<String>,
) {
    let local_attribute_name = local_name(attribute_name);
    let is_prefix_list = matches!(
        local_attribute_name,
        b"Requires"
            | b"Ignorable"
            | b"MustUnderstand"
            | b"ProcessContent"
            | b"PreserveElements"
            | b"PreserveAttributes"
    );
    for token in value.split_ascii_whitespace() {
        let prefix = if is_prefix_list {
            token.split_once(':').map_or(token, |(prefix, _)| prefix)
        } else if let Some((prefix, _)) = token.split_once(':') {
            prefix
        } else {
            continue;
        };
        if inherited.contains(prefix) && !prefix_is_locally_bound(prefix, local_scopes) {
            referenced.insert(prefix.to_owned());
        }
    }
}

fn capture_non_element(event: Event<'_>) -> Result<Option<Vec<u8>>> {
    let event = match event {
        Event::Text(value) => Event::Text(value.into_owned()),
        Event::CData(value) => Event::CData(value.into_owned()),
        Event::Comment(value) => Event::Comment(value.into_owned()),
        Event::PI(value) => Event::PI(value.into_owned()),
        Event::Decl(value) => Event::Decl(value.into_owned()),
        Event::DocType(value) => Event::DocType(value.into_owned()),
        Event::GeneralRef(value) => Event::GeneralRef(value.into_owned()),
        _ => return Ok(None),
    };
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event)?;
    Ok(Some(writer.into_inner()))
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

fn order_error(message: &str) -> OxmlError {
    OxmlError::InvalidValue(message.to_owned())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn unexpected(element: &BytesStart<'_>) -> OxmlError {
    OxmlError::UnexpectedElement(String::from_utf8_lossy(element.name().as_ref()).into_owned())
}
