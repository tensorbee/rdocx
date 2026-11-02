use std::collections::{HashMap, HashSet};
use std::io::Write;
use std::ops::Range;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_drawing::effect::{CT_EffectList, raw_is_effect_dag};
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
    FIXED_SHAPE_TREE_PREFIXES, MC_NS, NamespaceBindings, P_NS, R_NS, accessibility_parent_bindings,
    all_attributes, non_visual_drawing_decorative, non_visual_drawing_id, non_visual_drawing_name,
    non_visual_drawing_text, root_attributes, self_contained_attributes,
    set_non_visual_drawing_decorative, set_non_visual_drawing_name, set_non_visual_drawing_text,
};
use crate::picture::CT_Picture;
use crate::placeholder::{ApplicationProperties, CT_Placeholder, parse_application_properties};

pub type Result<T> = std::result::Result<T, OxmlError>;
type RawAttributes = Vec<(String, String)>;

/// The theme references python-pptx writes for a new autoshape.
const DEFAULT_STYLE: &str = r#"<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>"#;

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
    /// Reports the selected physical drawing owner's hidden paint flag.
    /// Semantic text consumers retain the shape and its content.
    pub fn hidden(&self) -> bool {
        match self {
            Self::Shape(value) => value.hidden(),
            Self::Picture(value) => value.hidden(),
            Self::GraphicFrame(value) => value.hidden(),
            Self::GroupShape(value) => value.hidden(),
            Self::Connector(value) => value.hidden(),
            Self::AlternateContent(value) => {
                value.chart_choice().is_some_and(CT_GraphicFrame::hidden)
            }
        }
    }

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
    ///
    /// Typed ids and preserved raw blobs are enough for a tree built from
    /// shapes. Serialising the whole tree is reserved for a child whose
    /// preserved bytes are not already stored on that child.
    pub fn scan(tree: &CT_ShapeTree) -> Self {
        let mut occupied = HashSet::new();
        collect_non_visual_ids(&tree.children, &mut occupied);
        if let Some(id) = tree.non_visual_id() {
            occupied.insert(id);
        }
        let mut ancestor_scope = false;
        occupy_group_preserved_ids(
            &tree.non_visual_group_properties,
            &tree.group_properties,
            &tree.raw_children,
            &mut occupied,
            &mut ancestor_scope,
        );
        for child in &tree.children {
            occupy_child_preserved_ids(child, &mut occupied, &mut ancestor_scope);
        }
        if ancestor_scope && let Ok(xml) = tree.to_xml() {
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
    rewrite_shape_ids_mapped(raw).map(|(rewritten, _)| rewritten)
}

/// Returns the qualified shape identity map for imported collaboration anchors.
pub fn rewrite_shape_ids_mapped(raw: &[u8]) -> Result<(Vec<u8>, HashMap<u32, u32>)> {
    let occurrences = shape_id_occurrences(raw, NamespaceBindings::default())?;
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
        return Ok((raw.to_vec(), map));
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
    Ok((rewritten, map))
}

fn shape_id_occurrences(
    raw: &[u8],
    inherited: NamespaceBindings,
) -> Result<Vec<ShapeIdOccurrence>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut scopes = vec![inherited];
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

fn cnvpr_count(bytes: &[u8]) -> usize {
    bytes
        .windows(5)
        .filter(|window| *window == b"cNvPr")
        .count()
}

/// A non-`p` `cNvPr` whose prefix is declared on an ancestor is only visible
/// in the serialised tree. `p:cNvPr` stays local because an unbound `p` is
/// PresentationML.
fn preserved_blob_needs_ancestor_scope(blob: &[u8]) -> bool {
    let mut search = blob;
    while let Some(position) = search.windows(5).position(|window| window == b"cNvPr") {
        let absolute = blob.len() - search.len() + position;
        let prefixed = absolute > 0 && blob[absolute - 1] == b':';
        let prefix_is_p = absolute >= 2 && &blob[absolute - 2..absolute] == b"p:";
        if prefixed && !prefix_is_p {
            let prefix_end = absolute - 1;
            let prefix_start = blob[..prefix_end]
                .iter()
                .rposition(|byte| !byte.is_ascii_alphanumeric())
                .map(|index| index + 1)
                .unwrap_or(0);
            let prefix = &blob[prefix_start..prefix_end];
            let mut declaration = b"xmlns:".to_vec();
            declaration.extend_from_slice(prefix);
            if !blob
                .windows(declaration.len())
                .any(|window| window == declaration.as_slice())
            {
                return true;
            }
        }
        if position + 5 >= search.len() {
            break;
        }
        search = &search[position + 5..];
    }
    false
}

fn occupy_raw_ids(
    raw: &OrderedRawChildren,
    occupied: &mut HashSet<u32>,
    ancestor_scope: &mut bool,
) {
    for blob in raw.iter() {
        occupy_blob_ids(blob, occupied, ancestor_scope);
    }
}

fn occupy_blob_ids(blob: &[u8], occupied: &mut HashSet<u32>, ancestor_scope: &mut bool) {
    if cnvpr_count(blob) == 0 {
        return;
    }
    if preserved_blob_needs_ancestor_scope(blob) {
        *ancestor_scope = true;
        return;
    }
    collect_preserved_non_visual_ids(blob, occupied);
}

fn occupy_group_preserved_ids(
    non_visual: &NonVisualGroupProperties,
    properties: &GroupProperties,
    raw_children: &OrderedRawChildren,
    occupied: &mut HashSet<u32>,
    ancestor_scope: &mut bool,
) {
    for blob in &non_visual.raw_children {
        occupy_blob_ids(blob, occupied, ancestor_scope);
    }
    occupy_raw_ids(&properties.raw_children, occupied, ancestor_scope);
    occupy_raw_ids(raw_children, occupied, ancestor_scope);
}

fn occupy_child_preserved_ids(
    child: &ShapeTreeChild,
    occupied: &mut HashSet<u32>,
    ancestor_scope: &mut bool,
) {
    match child {
        ShapeTreeChild::Shape(shape) => {
            occupy_shape_preserved_ids(shape, occupied, ancestor_scope);
        }
        ShapeTreeChild::GroupShape(group) => {
            occupy_group_preserved_ids(
                &group.non_visual_group_properties,
                &group.group_properties,
                &group.raw_children,
                occupied,
                ancestor_scope,
            );
            for child in &group.children {
                occupy_child_preserved_ids(child, occupied, ancestor_scope);
            }
        }
        ShapeTreeChild::AlternateContent(alternate) => {
            occupy_blob_ids(alternate.raw_xml(), occupied, ancestor_scope);
        }
        ShapeTreeChild::Picture(_)
        | ShapeTreeChild::GraphicFrame(_)
        | ShapeTreeChild::Connector(_) => {
            if let Ok(xml) = child_xml(child) {
                collect_preserved_non_visual_ids(&xml, occupied);
            }
        }
    }
}

fn child_xml(child: &ShapeTreeChild) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    child.write_xml(&mut writer)?;
    Ok(writer.into_inner())
}

fn occupy_shape_preserved_ids(
    shape: &CT_Shape,
    occupied: &mut HashSet<u32>,
    ancestor_scope: &mut bool,
) {
    let raw = &shape.raw;
    occupy_raw_ids(&raw.raw_children, occupied, ancestor_scope);
    occupy_raw_ids(&raw.non_visual_children, occupied, ancestor_scope);
    occupy_raw_ids(
        &raw.application_properties_raw_children,
        occupied,
        ancestor_scope,
    );
    occupy_blob_ids(&raw.non_visual_drawing_properties, occupied, ancestor_scope);
    occupy_blob_ids(&raw.non_visual_shape_properties, occupied, ancestor_scope);
    if let Some(style) = &raw.style {
        occupy_raw_ids(&style.raw_children, occupied, ancestor_scope);
        if let Some(source) = &style.source_xml {
            occupy_blob_ids(source, occupied, ancestor_scope);
        }
    }
    occupy_raw_ids(
        shape.shape_properties.raw_children(),
        occupied,
        ancestor_scope,
    );
    if let Some(text_body) = &shape.text_body {
        occupy_raw_ids(text_body.raw_children(), occupied, ancestor_scope);
    }
    if shape.shape_properties.custom_geometry.is_some()
        && let Ok(xml) = shape.to_xml()
    {
        collect_preserved_non_visual_ids(&xml, occupied);
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
#[derive(Clone, Debug)]
pub struct CT_AlternateContent {
    raw_source: std::sync::Arc<[u8]>,
    raw_range: Range<usize>,
    chart_choice: Option<Box<CT_GraphicFrame>>,
    selected_fallback: Option<Vec<ShapeTreeChild>>,
}

impl PartialEq for CT_AlternateContent {
    fn eq(&self, other: &Self) -> bool {
        self.raw_xml() == other.raw_xml()
            && self.chart_choice == other.chart_choice
            && self.selected_fallback == other.selected_fallback
    }
}

impl CT_AlternateContent {
    /// Returns the original subtree used as the sole serialisation source.
    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_source[self.raw_range.clone()]
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
        let (owner, _, _) = parse_shape_frames(xml, inherited, ShapeParseRoot::Alternate)?;
        let ParsedShapeOwner::Alternate(alternate) = owner else {
            unreachable!("the selected root is alternate content")
        };
        Ok(alternate)
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.get_mut().write_all(self.raw_xml())?;
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
#[derive(Clone, Debug)]
pub struct CT_ShapeStyle {
    pub line_reference: StyleMatrixReference,
    pub fill_reference: StyleMatrixReference,
    pub effect_reference: StyleMatrixReference,
    pub font_reference: FontReference,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
    source_xml: Option<Vec<u8>>,
    source_namespaces: NamespaceBindings,
}

impl PartialEq for CT_ShapeStyle {
    fn eq(&self, other: &Self) -> bool {
        self.line_reference == other.line_reference
            && self.fill_reference == other.fill_reference
            && self.effect_reference == other.effect_reference
            && self.font_reference == other.font_reference
            && self.raw_attributes == other.raw_attributes
            && self.raw_children == other.raw_children
    }
}
impl Eq for CT_ShapeStyle {}

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
    pub(crate) fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Box<Self>> {
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
                    let mut style = Self::from_reader(&mut reader, &start, &namespaces)?;
                    style.source_xml = Some(xml.to_vec());
                    style.source_namespaces = namespaces;
                    return Ok(style);
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
            source_xml: None,
            source_namespaces: namespaces.clone(),
        }))
    }

    /// Serializes the style on its own, for owners that keep it as bytes.
    pub(crate) fn to_fragment(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer)?;
        Ok(writer.into_inner())
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

fn prepare_shape_style_for_materialization(
    style: &CT_ShapeStyle,
    neutralize: bool,
) -> Result<Box<CT_ShapeStyle>> {
    let generated = style.to_fragment()?;
    let source = style.source_xml.as_deref().unwrap_or(&generated);
    let mut reader = Reader::from_reader(source);
    let mut scopes = vec![style.source_namespaces.clone()];
    let mut roles: Vec<u8> = Vec::new();
    let mut seen = HashSet::new();
    let mut writer = Writer::new(Vec::new());
    loop {
        match reader.read_event()? {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let empty = matches!(&event, Event::Empty(_));
                let start = match &event {
                    Event::Start(start) | Event::Empty(start) => start,
                    _ => unreachable!(),
                };
                let scope = scopes.last().unwrap().with_start(start)?;
                let name = start.name();
                let local = local_name(name.as_ref());
                let uri = scope.element_uri(name.as_ref());
                let reference = roles.len() == 1
                    && uri == Some(A_NS)
                    && matches!(local, b"lnRef" | b"fillRef" | b"effectRef" | b"fontRef");
                let color = roles.last().copied() == Some(1)
                    && matches!(local, b"srgbClr" | b"schemeClr" | b"sysClr" | b"prstClr");
                let transform = roles.last().copied() == Some(2)
                    && matches!(
                        local,
                        b"tint"
                            | b"shade"
                            | b"comp"
                            | b"inv"
                            | b"gray"
                            | b"alpha"
                            | b"alphaOff"
                            | b"alphaMod"
                            | b"hue"
                            | b"hueOff"
                            | b"hueMod"
                            | b"sat"
                            | b"satOff"
                            | b"satMod"
                            | b"lum"
                            | b"lumOff"
                            | b"lumMod"
                            | b"red"
                            | b"redOff"
                            | b"redMod"
                            | b"green"
                            | b"greenOff"
                            | b"greenMod"
                            | b"blue"
                            | b"blueOff"
                            | b"blueMod"
                            | b"gamma"
                            | b"invGamma"
                    );
                if transform && uri != Some(A_NS) {
                    return Err(OxmlError::InvalidValue(
                        "unrepresentable foreign style-color transform lookalike".into(),
                    ));
                }
                if color && uri != Some(A_NS) {
                    return Err(OxmlError::InvalidValue(
                        "unrepresentable foreign style color lookalike".into(),
                    ));
                }
                if reference && !seen.insert(local.to_vec()) {
                    return Err(OxmlError::InvalidValue(
                        "duplicate owned notes style reference".into(),
                    ));
                }
                let mut idx = None;
                for attribute in start.attributes() {
                    let attribute = attribute?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    if reference {
                        if key != b"idx" {
                            return Err(OxmlError::InvalidValue(
                                "unrepresentable opaque style-reference attribute".into(),
                            ));
                        }
                        let value = attribute.decoded_and_normalized_value(
                            XmlVersion::Implicit1_0,
                            start.decoder(),
                        )?;
                        if local == b"fontRef" {
                            if !matches!(value.as_ref(), "major" | "minor" | "none") {
                                return Err(OxmlError::InvalidValue(
                                    "unsupported source font collection reference".into(),
                                ));
                            }
                            idx = Some(if neutralize {
                                "none".to_owned()
                            } else {
                                match style.font_reference.index {
                                    oxml_drawing::style_ref::FontCollectionIndex::Major => "major",
                                    oxml_drawing::style_ref::FontCollectionIndex::Minor => "minor",
                                    oxml_drawing::style_ref::FontCollectionIndex::None => "none",
                                }
                                .to_owned()
                            });
                        } else {
                            value.parse::<u32>().map_err(|_| {
                                OxmlError::InvalidValue(
                                    "invalid source style matrix reference".into(),
                                )
                            })?;
                            idx = Some(if neutralize {
                                "0".to_owned()
                            } else {
                                match local {
                                    b"lnRef" => style.line_reference.index,
                                    b"fillRef" => style.fill_reference.index,
                                    _ => style.effect_reference.index,
                                }
                                .to_string()
                            });
                        }
                    } else if transform && key != b"val" {
                        return Err(OxmlError::InvalidValue(
                            "unrepresentable opaque style-color transform attribute".into(),
                        ));
                    } else if color && !matches!(key, b"val" | b"lastClr") {
                        return Err(OxmlError::InvalidValue(
                            "unrepresentable opaque source style-color attribute".into(),
                        ));
                    }
                }
                if reference && idx.is_none() {
                    return Err(OxmlError::MissingElement("owned source style idx".into()));
                }
                let mut prepared = start.to_owned();
                if let Some(idx) = idx {
                    let attributes = all_attributes(start)?;
                    prepared.clear_attributes();
                    for (key, value) in attributes {
                        let value = if key == "idx" { idx.as_str() } else { &value };
                        let escaped = quick_xml::escape::escape(value);
                        prepared.push_attribute((key.as_str(), escaped.as_ref()));
                    }
                }
                // Every transferred opaque descendant keeps its physical
                // source scope, including declarations on discarded roots.
                let attributes = all_attributes(&prepared)?;
                let mut entries = scope.entries();
                if !entries.iter().any(|(key, _)| key.is_empty()) {
                    entries.push((String::new(), String::new()));
                }
                for (prefix, value) in entries {
                    let key = if prefix.is_empty() {
                        "xmlns".into()
                    } else {
                        format!("xmlns:{prefix}")
                    };
                    if !attributes.iter().any(|(old, _)| old == &key) {
                        let escaped = quick_xml::escape::escape(&value);
                        prepared.push_attribute((key.as_str(), escaped.as_ref()));
                    }
                }
                writer.write_event(if empty {
                    Event::Empty(prepared)
                } else {
                    Event::Start(prepared)
                })?;
                if !empty {
                    scopes.push(scope);
                    roles.push(if reference {
                        1
                    } else if color {
                        2
                    } else {
                        0
                    });
                }
            }
            Event::End(end) => {
                writer.write_event(Event::End(end))?;
                scopes.pop();
                roles.pop();
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
    }
    if seen.len() != 4 {
        return Err(OxmlError::MissingElement(
            "four owned source style references".into(),
        ));
    }
    let xml = writer.into_inner();
    let staged = CT_ShapeStyle::from_fragment(&xml, &style.source_namespaces)?;
    if neutralize
        && (staged.line_reference.index != 0
            || staged.fill_reference.index != 0
            || staged.effect_reference.index != 0
            || staged.font_reference.index != oxml_drawing::style_ref::FontCollectionIndex::None)
    {
        return Err(OxmlError::InvalidValue(
            "source style neutralization failed owning reparse".into(),
        ));
    }
    if style_root_opaque_attribute_bytes(source)?
        != style_root_opaque_attribute_bytes(&staged.to_fragment()?)?
    {
        return Err(OxmlError::InvalidValue(
            "unrepresentable lexical source style-root metadata".into(),
        ));
    }
    Ok(staged)
}

fn style_root_opaque_attribute_bytes(xml: &[u8]) -> Result<Vec<Vec<u8>>> {
    let mut reader = Reader::from_reader(xml);
    let start = match reader.read_event()? {
        Event::Start(start) | Event::Empty(start) => start,
        _ => return Err(OxmlError::MissingElement("source style root".into())),
    };
    let body: &[u8] = start.as_ref();
    let mut result = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        let begin = name.as_ptr() as usize - body.as_ptr() as usize;
        let end =
            attribute.value.as_ptr() as usize - body.as_ptr() as usize + attribute.value.len() + 1;
        result.push(body[begin..end].to_vec());
    }
    Ok(result)
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

/// Validates recognized group and compatibility owners, returning their deepest level.
/// The shape-tree bookkeeping root is not a counted owner. Selected fallback
/// owners share their alternate-content parent's counter.
pub fn validate_shape_nesting(children: &[ShapeTreeChild]) -> Result<usize> {
    let mut pending = children
        .iter()
        .rev()
        .map(|child| (child, 0usize))
        .collect::<Vec<_>>();
    let mut deepest = 0;
    while let Some((child, parent_depth)) = pending.pop() {
        let descendants = match child {
            ShapeTreeChild::GroupShape(group) => Some(group.children.as_slice()),
            ShapeTreeChild::AlternateContent(alternate) => {
                Some(alternate.selected_fallback().unwrap_or(&[]))
            }
            _ => None,
        };
        if let Some(descendants) = descendants {
            let depth = parent_depth + 1;
            crate::check_presentation_nesting("group and alternate content", depth)?;
            deepest = deepest.max(depth);
            pending.extend(descendants.iter().rev().map(|child| (child, depth)));
        }
    }
    Ok(deepest)
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
    header_comments: [Vec<Vec<u8>>; 2],
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
    header_comments: [Vec<Vec<u8>>; 2],
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct GroupProperties {
    transform: Option<CT_Transform2D>,
    effects: Option<CT_EffectList>,
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
    header_comments: [Vec<Vec<u8>>; 2],
}

impl CT_Shape {
    fn hidden(&self) -> bool {
        let bindings = accessibility_parent_bindings(
            &self.raw.raw_attributes,
            &self.raw.non_visual_attributes,
        );
        crate::namespace::non_visual_drawing_hidden(
            &self.raw.non_visual_drawing_properties,
            &bindings,
        )
    }

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

    /// Returns the direct accessibility description, including an explicit empty value.
    pub fn alt_text(&self) -> Option<String> {
        non_visual_drawing_text(&self.raw.non_visual_drawing_properties, "descr")
    }

    /// Returns the direct accessibility title.
    pub fn title(&self) -> Option<String> {
        non_visual_drawing_text(&self.raw.non_visual_drawing_properties, "title")
    }

    /// Replaces or clears the direct accessibility description.
    pub fn set_alt_text(&mut self, text: Option<&str>) -> Result<()> {
        set_non_visual_drawing_text(&mut self.raw.non_visual_drawing_properties, "descr", text)
    }

    /// Replaces or clears the direct accessibility title.
    pub fn set_title(&mut self, text: Option<&str>) -> Result<()> {
        set_non_visual_drawing_text(&mut self.raw.non_visual_drawing_properties, "title", text)
    }

    /// Returns the qualified decorative flag without promoting defaults.
    pub fn decorative(&self) -> bool {
        let bindings = accessibility_parent_bindings(
            &self.raw.raw_attributes,
            &self.raw.non_visual_attributes,
        );
        non_visual_drawing_decorative(&self.raw.non_visual_drawing_properties, &bindings)
    }

    /// Changes only the qualified decorative flag, retaining unrelated extensions.
    pub fn set_decorative(&mut self, decorative: bool) -> Result<()> {
        let bindings = accessibility_parent_bindings(
            &self.raw.raw_attributes,
            &self.raw.non_visual_attributes,
        );
        set_non_visual_drawing_decorative(
            &mut self.raw.non_visual_drawing_properties,
            decorative,
            &bindings,
        )
    }

    /// Copies a placeholder with a fresh slide identity while retaining its
    /// producer properties and qualified opaque payload.
    pub fn clone_with_id(&self, id: u32) -> Result<Self> {
        let mut cloned = self.clone();
        let bindings = accessibility_parent_bindings(
            &cloned.raw.raw_attributes,
            &cloned.raw.non_visual_attributes,
        );
        let raw = &mut cloned.raw.non_visual_drawing_properties;
        let occurrence = shape_id_occurrences(raw, bindings)?
            .into_iter()
            .find(|value| value.defines_shape)
            .ok_or_else(|| OxmlError::MissingElement("placeholder cNvPr id".into()))?;
        raw.splice(occurrence.range, id.to_string().bytes());
        Ok(cloned)
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
                    namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
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

    /// Prepares actual source scopes without changing reference selections or owners.
    pub fn style_for_materialization(&self) -> Result<Option<CT_ShapeStyle>> {
        self.raw
            .style
            .as_ref()
            .map(|style| prepare_shape_style_for_materialization(style, false).map(|style| *style))
            .transpose()
    }

    /// Stages neutral references after their source appearance is materialized.
    pub fn materialize_style_references(&mut self) -> Result<()> {
        if let Some(style) = &self.raw.style {
            let staged = prepare_shape_style_for_materialization(style, true)?;
            self.raw.style = Some(staged);
        }
        Ok(())
    }

    /// Sets the `idx` of the style's `a:effectRef`, the theme effect style
    /// the shape draws. Index 0 references no theme effect. Returns `false`
    /// and changes nothing when the shape has no style.
    pub fn set_effect_reference_index(&mut self, index: u32) -> bool {
        let Some(style) = self.raw.style.as_mut() else {
            return false;
        };
        style.effect_reference.index = index;
        true
    }

    /// Sets `p:style` to the theme references python-pptx writes for a new
    /// autoshape: the first theme line, the third theme fill, and the second
    /// theme effect in `accent1`, and the minor font in `lt1`.
    ///
    /// A shape with neither a style nor a direct fill or line draws nothing.
    pub fn set_default_style(&mut self) -> Result<()> {
        let namespaces = NamespaceBindings::from_entries(&[
            ("p".to_owned(), P_NS.to_owned()),
            ("a".to_owned(), A_NS.to_owned()),
        ]);
        self.raw.style = Some(CT_ShapeStyle::from_fragment(
            DEFAULT_STYLE.as_bytes(),
            &namespaces,
        )?);
        Ok(())
    }

    /// Removes `p:style` when it is present.
    ///
    /// Returns false when the element is already absent. The remaining
    /// children stay in schema order.
    pub fn clear_style(&mut self) -> bool {
        if self.raw.style.is_none() {
            return false;
        }
        self.raw.style = None;
        true
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

    /// Converts the selected picture placeholder without replacing its identity.
    /// The native owner stages image allocation and supplies inherited geometry.
    pub fn fill_picture(
        &self,
        relationship_id: &str,
        transform: CT_Transform2D,
    ) -> Result<CT_Picture> {
        if self.placeholder.as_ref().map(|ph| ph.effective_type())
            != Some(crate::placeholder::PhType::Picture)
        {
            return Err(OxmlError::InvalidValue(
                "fill picture requires a picture placeholder".to_owned(),
            ));
        }
        // A new blip relationship uses the picture writer's fixed r prefix.
        // Its root cannot also retain an imported foreign r owner. Refuse before
        // conversion rather than discard that owner or emit duplicate bindings.
        if self
            .raw
            .raw_attributes
            .iter()
            .any(|(name, uri)| name == "xmlns:r" && uri != R_NS)
        {
            return Err(OxmlError::InvalidValue(
                "fill picture cannot preserve a foreign root relationship binding".to_owned(),
            ));
        }
        let unsupported_text_body = if let Some(body) = &self.text_body {
            // Conversion removes the text owner. Admit only its qualified empty
            // shell, independent of harmless producer namespace aliases.
            let xml = body
                .to_xml()
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
            let mut reader = NsReader::from_reader(xml.as_slice());
            let mut buffer = Vec::new();
            let mut depth = 0;
            let mut child_order = 0;
            let mut root_seen = false;
            let mut unsupported = false;
            loop {
                let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
                match event {
                    Event::Start(ref start) | Event::Empty(ref start) => {
                        let local = start.local_name();
                        let qualified = match namespace {
                            ResolveResult::Bound(Namespace(uri)) if depth == 0 => {
                                (uri == P_NS.as_bytes() || uri == A_NS.as_bytes())
                                    && local.as_ref() == b"txBody"
                                    && !root_seen
                            }
                            ResolveResult::Bound(Namespace(uri)) if depth == 1 => {
                                let next = match local.as_ref() {
                                    b"bodyPr" if child_order == 0 => 1,
                                    b"lstStyle" if child_order == 1 => 2,
                                    b"p" if child_order == 1 || child_order == 2 => 3,
                                    _ => 0,
                                };
                                let qualified = uri == A_NS.as_bytes() && next != 0;
                                if qualified {
                                    child_order = next;
                                }
                                qualified
                            }
                            _ => false,
                        };
                        if !qualified
                            || all_attributes(start)?
                                .iter()
                                .any(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
                        {
                            unsupported = true;
                            break;
                        }
                        root_seen = true;
                        if matches!(event, Event::Start(_)) {
                            depth += 1;
                        }
                    }
                    Event::End(_) => {
                        depth -= 1;
                    }
                    Event::Text(text) if text.iter().all(u8::is_ascii_whitespace) => {}
                    Event::Eof => break,
                    _ => {
                        unsupported = true;
                        break;
                    }
                }
                buffer.clear();
            }
            unsupported || !root_seen || depth != 0 || child_order != 3
        } else {
            false
        };
        if unsupported_text_body {
            return Err(OxmlError::InvalidValue(
                "picture placeholder text body contains source that a picture cannot retain"
                    .to_owned(),
            ));
        }
        // cNvPr and nvPr remain lexical source. Only the owned type-specific
        // container and the shared lock element change their schema names.
        let raw = &self.raw.non_visual_shape_properties;
        let fixed = NamespaceBindings::from_entries(&[
            ("p".to_owned(), P_NS.to_owned()),
            ("a".to_owned(), A_NS.to_owned()),
            ("r".to_owned(), R_NS.to_owned()),
        ]);
        let mut source_root = BytesStart::new("p:sp");
        push_attributes(&mut source_root, &self.raw.raw_attributes);
        let mut source_non_visual = BytesStart::new("p:nvSpPr");
        push_attributes(&mut source_non_visual, &self.raw.non_visual_attributes);
        let mut scopes = vec![
            fixed
                .with_start(&source_root)?
                .with_start(&source_non_visual)?,
        ];
        let mut reader = Reader::from_reader(raw.as_slice());
        let mut edits = Vec::new();
        let mut mapped = Vec::new();
        loop {
            let position = reader.buffer_position() as usize;
            let event = reader.read_event()?;
            match &event {
                Event::Start(start) | Event::Empty(start) => {
                    let scope = scopes.last().unwrap().with_start(start)?;
                    let owned_root = mapped.is_empty()
                        && scope.element_uri(start.name().as_ref()) == Some(P_NS)
                        && local_name(start.name().as_ref()) == b"cNvSpPr";
                    let owned_lock = mapped.len() == 1
                        && scope.element_uri(start.name().as_ref()) == Some(A_NS)
                        && local_name(start.name().as_ref()) == b"spLocks";
                    if owned_root || owned_lock {
                        for attribute in start.attributes() {
                            let attribute = attribute.map_err(OxmlError::from)?;
                            let key = attribute.key.as_ref();
                            if key == b"xmlns" || key.starts_with(b"xmlns:") || key.contains(&b':')
                            {
                                continue;
                            }
                            let shared = owned_lock
                                && matches!(
                                    key,
                                    b"noGrp"
                                        | b"noSelect"
                                        | b"noRot"
                                        | b"noChangeAspect"
                                        | b"noMove"
                                        | b"noResize"
                                );
                            if !shared {
                                return Err(OxmlError::InvalidValue(format!(
                                    "picture placeholder has an unrepresentable owned attribute {}",
                                    String::from_utf8_lossy(key)
                                )));
                            }
                        }
                        let name = start.name();
                        let old = std::str::from_utf8(name.as_ref())?;
                        let prefix = old.rsplit_once(':').map(|(prefix, _)| prefix);
                        let local = if owned_root { "cNvPicPr" } else { "picLocks" };
                        let new = prefix
                            .map_or_else(|| local.to_owned(), |prefix| format!("{prefix}:{local}"));
                        edits.push((position + 1..position + 1 + old.len(), new.clone()));
                        if matches!(event, Event::Start(_)) {
                            mapped.push(Some(new));
                            scopes.push(scope);
                        }
                    } else if matches!(event, Event::Start(_)) {
                        mapped.push(None);
                        scopes.push(scope);
                    }
                }
                Event::End(end) => {
                    if let Some(Some(new)) = mapped.pop() {
                        edits.push((position + 2..position + 2 + end.name().as_ref().len(), new));
                    }
                    scopes.pop();
                }
                Event::Eof => break,
                _ => {}
            }
        }
        let mut picture_properties = raw.to_vec();
        for (range, replacement) in edits.into_iter().rev() {
            picture_properties.splice(range, replacement.into_bytes());
        }
        let id = self
            .non_visual_id()
            .ok_or_else(|| OxmlError::InvalidValue("placeholder has no shape id".to_owned()))?;
        let base = CT_Picture::new(
            id,
            self.non_visual_name().unwrap_or(""),
            relationship_id,
            transform.clone(),
        )?;
        let mut writer = Writer::new(Vec::new());
        let mut start = BytesStart::new("p:pic");
        for (prefix, uri) in [
            ("xmlns:p", P_NS),
            ("xmlns:a", A_NS),
            ("xmlns:r", R_NS),
            ("xmlns:mc", MC_NS),
        ] {
            start.push_attribute((prefix, uri));
        }
        push_attributes(&mut start, &self.raw.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(&mut writer, self.raw.raw_children.at(0))?;
        self.write_non_visual_properties(&mut writer, Some(&picture_properties))?;
        emit_raw(&mut writer, self.raw.raw_children.at(1))?;
        base.blip_fill
            .as_ref()
            .ok_or_else(|| OxmlError::MissingElement("picture blip fill".to_owned()))?
            .write_xml_as(&mut writer, "p:blipFill")
            .map_err(|e| OxmlError::InvalidValue(e.to_string()))?;
        let mut properties = (*self.shape_properties).clone();
        properties.transform = Some(transform);
        // A newly filled placeholder needs real picture geometry in the saved file.
        // Retain imported geometry and its opaque payload rather than replacing it.
        if properties.custom_geometry.is_none() && properties.preset_geometry.is_none() {
            properties.preset_geometry = base.shape_properties.preset_geometry.clone();
        }
        properties
            .write_xml_as(&mut writer, "p:spPr")
            .map_err(|e| OxmlError::InvalidValue(e.to_string()))?;
        for boundary in 2..=5 {
            emit_raw(&mut writer, self.raw.raw_children.at(boundary))?;
            if boundary == 2
                && let Some(style) = &self.raw.style
            {
                style.write_xml(&mut writer)?;
            }
        }
        writer.write_event(Event::End(BytesEnd::new("p:pic")))?;
        CT_Picture::from_xml(&writer.into_inner())
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
            if !self
                .raw
                .raw_attributes
                .iter()
                .any(|(name, _)| name == "xmlns:r")
            {
                start.push_attribute(("xmlns:r", R_NS));
            }
            start.push_attribute(("xmlns:mc", MC_NS));
        }
        push_attributes(&mut start, &self.raw.raw_attributes);
        writer.write_event(Event::Start(start))?;

        emit_raw(writer, self.raw.raw_children.at(0))?;
        self.write_non_visual_properties(writer, None)?;
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

    fn write_non_visual_properties<W: Write>(
        &self,
        writer: &mut Writer<W>,
        picture_properties: Option<&[u8]>,
    ) -> Result<()> {
        let tag = if picture_properties.is_some() {
            "p:nvPicPr"
        } else {
            "p:nvSpPr"
        };
        let mut start = BytesStart::new(tag);
        push_attributes(&mut start, &self.raw.non_visual_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw.non_visual_children.at(0))?;
        writer
            .get_mut()
            .write_all(&self.raw.non_visual_drawing_properties)?;
        emit_raw(writer, self.raw.non_visual_children.at(1))?;
        writer
            .get_mut()
            .write_all(picture_properties.unwrap_or(&self.raw.non_visual_shape_properties))?;
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
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
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
        namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
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
                CT_ShapeProperties::from_xml_with_inherited_namespaces(&raw, &namespaces.entries())
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
            namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
            *text_body = Some(
                CT_TextBody::from_xml_with_inherited_namespaces(&raw, &namespaces.entries())
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
        namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
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
                effects: None,
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            children: Vec::new(),
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
            header_comments: Default::default(),
        }
    }

    /// Returns the shape tree's own `p:cNvPr/@id`, when present.
    pub fn non_visual_id(&self) -> Option<u32> {
        self.non_visual_group_properties.non_visual_id
    }

    /// Appends one member right after the last typed member.
    ///
    /// Preserved content after that member, an unmodelled member such as a
    /// trailing `p:contentPart` as well as schema-final `p:extLst`, stays
    /// after the new member.
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
        let range = direct_shape_child_ranges(&xml)?
            .get(index)
            .cloned()
            .ok_or_else(|| {
                OxmlError::InvalidValue(format!("shape id {id} disappeared during removal"))
            })?;
        let mut rewritten = Vec::with_capacity(xml.len() - range.len());
        rewritten.extend_from_slice(&xml[..range.start]);
        rewritten.extend_from_slice(&xml[range.end..]);
        *self = Self::from_xml(&rewritten)?;
        Ok(Some(removed))
    }

    /// Moves one immediate child so that it ends up at index `to`, where
    /// later children draw on top.
    ///
    /// The moved child passes only the children between its old and new
    /// index. Unmodelled members, such as `p:contentPart`, and schema-final
    /// content keep their bytes and their place among the other children.
    pub fn move_child(&mut self, from: usize, to: usize) -> Result<()> {
        let count = self.children.len();
        if let Some(index) = [from, to].into_iter().find(|index| *index >= count) {
            return Err(OxmlError::InvalidValue(format!(
                "shape-tree child index {index} is out of range for {count} children"
            )));
        }
        if from == to {
            return Ok(());
        }
        let xml = self.to_xml()?;
        let ranges = direct_shape_child_ranges(&xml)?;
        if ranges.len() != count {
            return Err(OxmlError::InvalidValue(
                "shape-tree children changed during a move".to_owned(),
            ));
        }
        let moved = ranges[from].clone();
        let mut rewritten = Vec::with_capacity(xml.len());
        if from < to {
            let after = ranges[to].end;
            rewritten.extend_from_slice(&xml[..moved.start]);
            rewritten.extend_from_slice(&xml[moved.end..after]);
            rewritten.extend_from_slice(&xml[moved]);
            rewritten.extend_from_slice(&xml[after..]);
        } else {
            let before = ranges[to].start;
            rewritten.extend_from_slice(&xml[..before]);
            rewritten.extend_from_slice(&xml[moved.clone()]);
            rewritten.extend_from_slice(&xml[before..moved.start]);
            rewritten.extend_from_slice(&xml[moved.end..]);
        }
        *self = Self::from_xml(&rewritten)?;
        Ok(())
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
            header_comments: parsed.header_comments,
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
        validate_shape_nesting(&self.children)?;
        write_group(
            writer,
            GroupKind::ShapeTree,
            &self.non_visual_group_properties,
            &self.group_properties,
            &self.children,
            &self.raw_attributes,
            &self.raw_children,
            &self.header_comments,
            true,
        )
    }
}

/// Returns the byte range of every typed shape-tree child, in child order.
fn direct_shape_child_ranges(xml: &[u8]) -> Result<Vec<Range<usize>>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut ranges = Vec::new();
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
                if parse_shape_tree_child(&name, uri, &raw, &namespaces)?.is_some() {
                    ranges.push(range);
                }
            }
            Event::Empty(child) => {
                let namespaces = root.as_ref().expect("shape-tree root").with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let uri = namespaces.element_uri(child.name().as_ref());
                let range = shape_start_tag_range(xml, reader.buffer_position() as usize)?;
                let raw = capture_empty_element(&child)?;
                if parse_shape_tree_child(&name, uri, &raw, &namespaces)?.is_some() {
                    ranges.push(range);
                }
            }
            Event::End(_) | Event::Eof => return Ok(ranges),
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
    fn hidden(&self) -> bool {
        let bindings = accessibility_parent_bindings(
            &self.raw_attributes,
            &self.non_visual_group_properties.raw_attributes,
        );
        self.accessibility_xml()
            .is_some_and(|xml| crate::namespace::non_visual_drawing_hidden(xml, &bindings))
    }

    /// Copies an ancestor group with a fresh identity, retaining exact geometry
    /// and producer metadata for a selected placeholder descendant.
    pub fn clone_with_id(&self, id: u32) -> Result<Self> {
        let mut cloned = self.clone();
        set_non_visual_drawing_text(cloned.accessibility_xml_mut()?, "id", Some(&id.to_string()))?;
        cloned.non_visual_group_properties.non_visual_id = Some(id);
        Ok(cloned)
    }

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
                effects: None,
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
            header_comments: Default::default(),
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

    /// Appends one member right after the last typed member, as
    /// [`CT_ShapeTree::append_child`] does.
    ///
    /// Preserved content after that member, an unmodelled member such as a
    /// trailing `p:contentPart` as well as schema-final `p:extLst`, stays
    /// after the new member.
    pub fn append_child(&mut self, child: ShapeTreeChild) -> &mut ShapeTreeChild {
        self.raw_children.shift_boundaries_from(self.children.len());
        self.children.push(child);
        self.children.last_mut().expect("group child was appended")
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
            header_comments: parsed.header_comments,
        }
    }

    /// Returns the typed DrawingML group transform when `p:grpSpPr` has one.
    pub fn group_transform(&self) -> Option<&CT_Transform2D> {
        self.group_properties.transform.as_ref()
    }

    /// Returns the group transform, creating an empty one when absent.
    pub fn group_transform_mut(&mut self) -> &mut CT_Transform2D {
        self.group_properties
            .transform
            .get_or_insert_with(CT_Transform2D::default)
    }

    /// Returns the direct `a:effectLst` of `p:grpSpPr`, when present.
    pub fn effects(&self) -> Option<&CT_EffectList> {
        self.group_properties.effects.as_ref()
    }

    /// Replaces or removes the direct `a:effectLst` of `p:grpSpPr`.
    ///
    /// An `a:effectDag` excludes an effect list, so a group that carries one
    /// refuses a new list rather than writing both.
    pub fn set_effects(&mut self, effects: Option<CT_EffectList>) -> Result<()> {
        if effects.is_some() && self.group_properties.has_effect_dag() {
            return Err(OxmlError::InvalidValue(
                "p:grpSpPr carries an a:effectDag, which excludes an a:effectLst".to_owned(),
            ));
        }
        self.group_properties.effects = effects;
        Ok(())
    }

    fn accessibility_xml(&self) -> Option<&[u8]> {
        self.non_visual_group_properties
            .raw_children
            .get(self.non_visual_group_properties.drawing_properties_index?)
            .map(Vec::as_slice)
    }

    fn accessibility_xml_mut(&mut self) -> Result<&mut Vec<u8>> {
        let index = self
            .non_visual_group_properties
            .drawing_properties_index
            .ok_or_else(|| OxmlError::MissingElement("p:cNvPr".to_owned()))?;
        self.non_visual_group_properties
            .raw_children
            .get_mut(index)
            .ok_or_else(|| OxmlError::MissingElement("p:cNvPr".to_owned()))
    }

    /// Returns the direct accessibility description.
    pub fn alt_text(&self) -> Option<String> {
        non_visual_drawing_text(self.accessibility_xml()?, "descr")
    }

    /// Returns the direct accessibility title.
    pub fn title(&self) -> Option<String> {
        non_visual_drawing_text(self.accessibility_xml()?, "title")
    }

    /// Replaces or clears the direct accessibility description.
    pub fn set_alt_text(&mut self, text: Option<&str>) -> Result<()> {
        set_non_visual_drawing_text(self.accessibility_xml_mut()?, "descr", text)
    }

    /// Replaces or clears the direct accessibility title.
    pub fn set_title(&mut self, text: Option<&str>) -> Result<()> {
        set_non_visual_drawing_text(self.accessibility_xml_mut()?, "title", text)
    }

    /// Returns the qualified decorative flag without promoting defaults.
    pub fn decorative(&self) -> bool {
        let bindings = accessibility_parent_bindings(
            &self.raw_attributes,
            &self.non_visual_group_properties.raw_attributes,
        );
        self.accessibility_xml()
            .is_some_and(|xml| non_visual_drawing_decorative(xml, &bindings))
    }

    /// Changes only the qualified decorative flag, retaining unrelated extensions.
    pub fn set_decorative(&mut self, decorative: bool) -> Result<()> {
        let bindings = accessibility_parent_bindings(
            &self.raw_attributes,
            &self.non_visual_group_properties.raw_attributes,
        );
        set_non_visual_drawing_decorative(self.accessibility_xml_mut()?, decorative, &bindings)
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
        let depth = 1 + validate_shape_nesting(&self.children)?;
        crate::check_presentation_nesting("group and alternate content", depth)?;
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
            &self.header_comments,
            declare_namespaces,
        )
    }
}

struct GroupParseFrame {
    non_visual: Option<NonVisualGroupProperties>,
    properties: Option<GroupProperties>,
    children: Vec<ShapeTreeChild>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
    header_comments: [Vec<Vec<u8>>; 2],
    state: usize,
}

enum ShapeParseOwner {
    Group(Box<GroupParseFrame>),
    Alternate {
        start: usize,
        chart_choice: Option<Box<CT_GraphicFrame>>,
        fallback: Option<Vec<ShapeTreeChild>>,
    },
    Fallback(Vec<ShapeTreeChild>),
}

struct ShapeParseFrame {
    namespaces: NamespaceBindings,
    depth: usize,
    owner: ShapeParseOwner,
}

enum ParsedShapeOwner {
    Group(Box<ParsedGroup>),
    Alternate(CT_AlternateContent),
    Fallback(Vec<ShapeTreeChild>),
}

#[derive(Clone, Copy)]
enum ShapeParseRoot {
    Group(GroupKind),
    Alternate,
}

impl ShapeParseFrame {
    fn group(start: &BytesStart<'_>, namespaces: NamespaceBindings, depth: usize) -> Result<Self> {
        crate::check_presentation_nesting("group and alternate content", depth)?;
        namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
        let raw_attributes =
            self_contained_attributes(start, FIXED_SHAPE_TREE_PREFIXES, &namespaces.entries())?;
        Ok(Self {
            namespaces,
            depth,
            owner: ShapeParseOwner::Group(Box::new(GroupParseFrame {
                non_visual: None,
                properties: None,
                children: Vec::new(),
                raw_attributes,
                raw_children: OrderedRawChildren::default(),
                header_comments: Default::default(),
                state: 0,
            })),
        })
    }

    fn finish(self, source: &std::sync::Arc<[u8]>, end: usize) -> Result<ParsedShapeOwner> {
        match self.owner {
            ShapeParseOwner::Group(group) => Ok(ParsedShapeOwner::Group(Box::new(ParsedGroup {
                non_visual_group_properties: required(group.non_visual, "p:nvGrpSpPr")?,
                group_properties: required(group.properties, "p:grpSpPr")?,
                children: group.children,
                raw_attributes: group.raw_attributes,
                raw_children: group.raw_children,
                header_comments: group.header_comments,
            }))),
            ShapeParseOwner::Alternate {
                start,
                chart_choice,
                fallback,
            } => Ok(ParsedShapeOwner::Alternate(CT_AlternateContent {
                raw_source: std::sync::Arc::clone(source),
                raw_range: start..end,
                chart_choice,
                selected_fallback: fallback,
            })),
            ShapeParseOwner::Fallback(children) => Ok(ParsedShapeOwner::Fallback(children)),
        }
    }

    fn attach(&mut self, owner: ParsedShapeOwner) -> Result<()> {
        let child = match owner {
            ParsedShapeOwner::Group(group) => {
                ShapeTreeChild::GroupShape(Box::new(CT_GroupShape::from_parsed(*group)))
            }
            ParsedShapeOwner::Alternate(alternate) => {
                ShapeTreeChild::AlternateContent(Box::new(alternate))
            }
            ParsedShapeOwner::Fallback(children) => {
                let ShapeParseOwner::Alternate { fallback, .. } = &mut self.owner else {
                    return Err(OxmlError::InvalidValue(
                        "mc:Fallback requires an immediate mc:AlternateContent owner".to_owned(),
                    ));
                };
                if fallback.is_some() {
                    return Err(duplicate_fallback());
                }
                *fallback = Some(children);
                return Ok(());
            }
        };
        match &mut self.owner {
            ShapeParseOwner::Group(group) => group.children.push(child),
            ShapeParseOwner::Fallback(children) => children.push(child),
            ShapeParseOwner::Alternate { .. } => {
                return Err(OxmlError::InvalidValue(
                    "mc:AlternateContent requires an immediate branch owner".to_owned(),
                ));
            }
        }
        Ok(())
    }
}

/// Recursive owners use one cursor. Only leaf or opaque boundaries are captured.
fn parse_shape_frames(
    xml: &[u8],
    inherited: &[(String, String)],
    root: ShapeParseRoot,
) -> Result<(ParsedShapeOwner, usize, usize)> {
    let mut reader = Reader::from_reader(xml);
    let source: std::sync::Arc<[u8]> = std::sync::Arc::from(xml);
    let mut frames: Vec<ShapeParseFrame> = Vec::new();
    let mut buffer = Vec::new();
    let mut visits = 0;
    let mut captured_bytes = 0;
    loop {
        let start_position = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        visits += 1;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let namespaces = match frames.last() {
                    Some(frame) => frame.namespaces.with_start(&start)?,
                    None => NamespaceBindings::from_entries(inherited).with_start(&start)?,
                };
                let name = local_name(start.name().as_ref()).to_vec();
                let uri = namespaces.element_uri(start.name().as_ref());
                if frames.is_empty() {
                    let frame = match root {
                        ShapeParseRoot::Group(kind)
                            if name == kind.local_name() && uri == Some(P_NS) =>
                        {
                            let depth = usize::from(matches!(kind, GroupKind::GroupShape));
                            ShapeParseFrame::group(&start, namespaces, depth)?
                        }
                        ShapeParseRoot::Alternate
                            if name == b"AlternateContent" && uri == Some(MC_NS) =>
                        {
                            crate::check_presentation_nesting("group and alternate content", 1)?;
                            ShapeParseFrame {
                                namespaces,
                                depth: 1,
                                owner: ShapeParseOwner::Alternate {
                                    start: start_position,
                                    chart_choice: None,
                                    fallback: None,
                                },
                            }
                        }
                        _ => return Err(unexpected(&start)),
                    };
                    if empty {
                        return Ok((
                            frame.finish(&source, reader.buffer_position() as usize)?,
                            visits,
                            captured_bytes,
                        ));
                    }
                    frames.push(frame);
                } else {
                    let parent = frames.last_mut().expect("the current owner is retained");
                    let can_descend = match &parent.owner {
                        ShapeParseOwner::Group(group) => group.state == 2,
                        ShapeParseOwner::Fallback(_) => true,
                        ShapeParseOwner::Alternate { .. } => false,
                    };
                    let group = can_descend && uri == Some(P_NS) && name == b"grpSp";
                    let alternate =
                        can_descend && uri == Some(MC_NS) && name == b"AlternateContent";
                    if group || alternate {
                        let depth = parent.depth + 1;
                        crate::check_presentation_nesting("group and alternate content", depth)?;
                        let frame = if group {
                            ShapeParseFrame::group(&start, namespaces, depth)?
                        } else {
                            ShapeParseFrame {
                                namespaces,
                                depth,
                                owner: ShapeParseOwner::Alternate {
                                    start: start_position,
                                    chart_choice: None,
                                    fallback: None,
                                },
                            }
                        };
                        if empty {
                            parent.attach(
                                frame.finish(&source, reader.buffer_position() as usize)?,
                            )?;
                        } else {
                            frames.push(frame);
                        }
                    } else if matches!(parent.owner, ShapeParseOwner::Alternate { .. })
                        && uri == Some(MC_NS)
                        && name == b"Fallback"
                    {
                        let ShapeParseOwner::Alternate { fallback, .. } = &parent.owner else {
                            unreachable!()
                        };
                        if fallback.is_some() {
                            return Err(duplicate_fallback());
                        }
                        let frame = ShapeParseFrame {
                            namespaces,
                            depth: parent.depth,
                            owner: ShapeParseOwner::Fallback(Vec::new()),
                        };
                        if empty {
                            parent.attach(
                                frame.finish(&source, reader.buffer_position() as usize)?,
                            )?;
                        } else {
                            frames.push(frame);
                        }
                    } else {
                        let raw = if empty {
                            capture_empty_element(&start)?
                        } else {
                            capture_element(&mut reader, &start)?
                        };
                        captured_bytes += raw.len();
                        match &mut parent.owner {
                            ShapeParseOwner::Group(group) => capture_group_child(
                                &name,
                                uri,
                                raw,
                                &namespaces,
                                &mut group.non_visual,
                                &mut group.properties,
                                &mut group.children,
                                &mut group.raw_children,
                                &mut group.state,
                            )?,
                            ShapeParseOwner::Fallback(children) => {
                                if let Some(child) =
                                    parse_shape_tree_child(&name, uri, &raw, &namespaces)?
                                {
                                    children.push(child);
                                }
                            }
                            ShapeParseOwner::Alternate { chart_choice, .. } => {
                                if uri == Some(MC_NS) && name == b"Choice" && chart_choice.is_none()
                                {
                                    *chart_choice =
                                        parse_chart_choice(&raw, &namespaces.entries())?;
                                }
                            }
                        }
                    }
                }
            }
            Event::End(_) => {
                let frame = frames.pop().ok_or_else(|| {
                    OxmlError::InvalidValue(
                        "shape owner closing element has no open owner".to_owned(),
                    )
                })?;
                let owner = frame.finish(&source, reader.buffer_position() as usize)?;
                if let Some(parent) = frames.last_mut() {
                    parent.attach(owner)?;
                } else {
                    return Ok((owner, visits, captured_bytes));
                }
            }
            Event::Comment(_) | Event::PI(_) => {
                if let Some(frame) = frames.last_mut()
                    && let ShapeParseOwner::Group(group) = &mut frame.owner
                {
                    let raw = xml[start_position..reader.buffer_position() as usize].to_vec();
                    captured_bytes += raw.len();
                    if group.state < 2 {
                        group.header_comments[group.state].push(raw);
                    } else {
                        group.raw_children.push(group.children.len(), raw);
                    }
                }
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "closing shape-tree owner".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_group(xml: &[u8], inherited: &[(String, String)], kind: GroupKind) -> Result<ParsedGroup> {
    let (owner, _, _) = parse_shape_frames(xml, inherited, ShapeParseRoot::Group(kind))?;
    let ParsedShapeOwner::Group(group) = owner else {
        unreachable!("the selected root is a group")
    };
    Ok(*group)
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
        namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
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
    header_comments: &[Vec<Vec<u8>>; 2],
    declare_namespaces: bool,
) -> Result<()> {
    let mut start = BytesStart::new(kind.tag());
    if declare_namespaces {
        start.push_attribute(("xmlns:p", P_NS));
        start.push_attribute(("xmlns:a", A_NS));
        if !attributes.iter().any(|(name, _)| name == "xmlns:r") {
            start.push_attribute(("xmlns:r", R_NS));
        }
        start.push_attribute(("xmlns:mc", MC_NS));
    }
    push_attributes(&mut start, attributes);
    writer.write_event(Event::Start(start))?;
    emit_raw(writer, header_comments[0].iter().map(Vec::as_slice))?;
    non_visual.write_xml(writer)?;
    emit_raw(writer, header_comments[1].iter().map(Vec::as_slice))?;
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
                        effects: None,
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
        let mut properties = Self {
            transform: None,
            effects: None,
            raw_attributes: root_attributes(start, FIXED_SHAPE_TREE_PREFIXES)?,
            raw_children: OrderedRawChildren::default(),
        };
        let mut boundary = 0usize;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let name = local_name(child.name().as_ref()).to_vec();
                    let modelled = is_modelled_group_property(namespaces, &child)?;
                    let raw = capture_element(reader, &child)?;
                    properties.capture_child(&name, modelled, raw, &mut boundary)?;
                }
                Event::Empty(child) => {
                    let name = local_name(child.name().as_ref()).to_vec();
                    let modelled = is_modelled_group_property(namespaces, &child)?;
                    let raw = capture_empty_element(&child)?;
                    properties.capture_child(&name, modelled, raw, &mut boundary)?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"grpSpPr" => break,
                Event::Eof => {
                    return Err(OxmlError::MissingElement("closing p:grpSpPr".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
        Ok(properties)
    }

    /// Records one `p:grpSpPr` child. Boundary 0 lies between `a:xfrm` and
    /// the effect slot, boundary 1 after it, where `a:scene3d` and
    /// `a:extLst` always go so that an effect list added later precedes them.
    fn capture_child(
        &mut self,
        name: &[u8],
        modelled: bool,
        raw: Vec<u8>,
        boundary: &mut usize,
    ) -> Result<()> {
        match (modelled, name) {
            (true, b"xfrm") => {
                if self.transform.is_some() {
                    return Err(OxmlError::InvalidValue(
                        "duplicate DrawingML group transform".to_owned(),
                    ));
                }
                if !self.raw_children.is_empty() || self.effects.is_some() {
                    return Err(OxmlError::InvalidValue(
                        "a:xfrm must precede other p:grpSpPr children".to_owned(),
                    ));
                }
                self.transform = Some(
                    CT_Transform2D::from_xml(&raw)
                        .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
                );
            }
            (true, _) if self.effects.is_none() => {
                self.effects = Some(
                    CT_EffectList::from_xml(&raw)
                        .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
                );
                *boundary = 1;
            }
            _ => {
                if matches!(name, b"scene3d" | b"extLst") {
                    *boundary = 1;
                }
                self.raw_children.push(*boundary, raw);
            }
        }
        Ok(())
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:grpSpPr");
        push_attributes(&mut start, &self.raw_attributes);
        if self.transform.is_none() && self.effects.is_none() && self.raw_children.is_empty() {
            writer.write_event(Event::Empty(start))?;
            return Ok(());
        }
        writer.write_event(Event::Start(start))?;
        if let Some(transform) = &self.transform {
            transform
                .write_xml(writer)
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        }
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(effects) = &self.effects {
            effects
                .write_xml(writer)
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        writer.write_event(Event::End(BytesEnd::new("p:grpSpPr")))?;
        Ok(())
    }

    fn has_effect_dag(&self) -> bool {
        (0..=1).any(|boundary| self.raw_children.at(boundary).any(raw_is_effect_dag))
    }
}

/// Reports whether a `p:grpSpPr` child is the DrawingML `a:xfrm` or
/// `a:effectLst` it models, after checking that its fixed `a:` prefix can be
/// written back.
fn is_modelled_group_property(
    namespaces: &NamespaceBindings,
    child: &BytesStart<'_>,
) -> Result<bool> {
    let child_namespaces = namespaces.with_start(child)?;
    let modelled = child_namespaces.element_uri(child.name().as_ref()) == Some(A_NS)
        && matches!(local_name(child.name().as_ref()), b"xfrm" | b"effectLst");
    if modelled {
        child_namespaces.reject_writer_conflicts(&["p", "a", "mc"])?;
    }
    Ok(modelled)
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

impl CT_ShapeTree {
    /// Adds source notes-master members while retaining this note's raw tree authority.
    pub fn materialize_notes_master_surface(&mut self, master: &Self) -> Result<()> {
        // Master and note ids belong to distinct original physical owners.
        // Allocate before merging so a decoration cannot alias a local shape.
        let local_xml = self.to_xml()?;
        let master_xml = master.to_xml()?;
        let mut occupied: HashSet<_> =
            shape_id_occurrences(&local_xml, NamespaceBindings::default())?
                .into_iter()
                .filter(|entry| entry.defines_shape)
                .map(|entry| entry.id)
                .collect();
        let occurrences = shape_id_occurrences(&master_xml, NamespaceBindings::default())?;
        let mut mapping = HashMap::new();
        let mut defined = HashSet::new();
        for entry in occurrences.iter().filter(|entry| entry.defines_shape) {
            if !defined.insert(entry.id) {
                return Err(OxmlError::InvalidValue(
                    "duplicate source notes-master shape identity".into(),
                ));
            }
        }
        let mut next = 2u32;
        for occurrence in occurrences
            .iter()
            .filter(|entry| entry.defines_shape)
            .skip(1)
        {
            if mapping.contains_key(&occurrence.id) {
                return Err(OxmlError::InvalidValue(
                    "duplicate source notes-master shape identity".into(),
                ));
            }
            while occupied.contains(&next) {
                next = next.checked_add(1).ok_or_else(shape_ids_exhausted)?;
            }
            mapping.insert(occurrence.id, next);
            occupied.insert(next);
            next = next.checked_add(1).ok_or_else(shape_ids_exhausted)?;
        }
        for source in &master.children {
            let key = match source {
                ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref().map(CT_Placeholder::key),
                ShapeTreeChild::Picture(picture) => {
                    picture.placeholder.as_ref().map(CT_Placeholder::key)
                }
                _ => None,
            };
            if let (Some(key), Some(source_id)) = (key, source.non_visual_id()) {
                for local in &self.children {
                    let local_key = match local {
                        ShapeTreeChild::Shape(shape) => {
                            shape.placeholder.as_ref().map(CT_Placeholder::key)
                        }
                        ShapeTreeChild::Picture(picture) => {
                            picture.placeholder.as_ref().map(CT_Placeholder::key)
                        }
                        _ => None,
                    };
                    if local_key.is_some_and(|local| key.matches(&local))
                        && let Some(local_id) = local.non_visual_id()
                    {
                        mapping.insert(source_id, local_id);
                    }
                }
            }
        }
        let mut rewritten = Vec::new();
        let mut copied = 0;
        for occurrence in occurrences {
            if let Some(id) = mapping.get(&occurrence.id) {
                rewritten.extend_from_slice(&master_xml[copied..occurrence.range.start]);
                rewritten.extend_from_slice(id.to_string().as_bytes());
                copied = occurrence.range.end;
            }
        }
        rewritten.extend_from_slice(&master_xml[copied..]);
        let prepared_master = Self::from_xml(&rewritten)?;
        let master = &prepared_master;
        let mut candidate = self.clone();
        let normalize_group = |group: &GroupProperties| {
            let mut group = group.clone();
            if let Some(transform) = &mut group.transform {
                if transform
                    .offset
                    .is_some_and(|point| point.x.0 == 0 && point.y.0 == 0)
                {
                    transform.offset = None;
                }
                if transform
                    .extent
                    .is_some_and(|size| size.cx.0 == 0 && size.cy.0 == 0)
                {
                    transform.extent = None;
                }
                if transform
                    .child_offset
                    .is_some_and(|point| point.x.0 == 0 && point.y.0 == 0)
                {
                    transform.child_offset = None;
                }
                if transform
                    .child_extent
                    .is_some_and(|size| size.cx.0 == 0 && size.cy.0 == 0)
                {
                    transform.child_extent = None;
                }
                if *transform == CT_Transform2D::default() {
                    group.transform = None;
                }
            }
            group
        };
        if normalize_group(&master.group_properties) != normalize_group(&self.group_properties)
            || master.non_visual_group_properties != self.non_visual_group_properties
        {
            return Err(OxmlError::InvalidValue(
                "unrepresentable source notes-master tree group authority".into(),
            ));
        }
        let bindings = |tree: &Self| -> Result<NamespaceBindings> {
            let mut start = BytesStart::new("p:spTree");
            push_attributes(&mut start, &tree.raw_attributes);
            NamespaceBindings::from_entries(&[
                ("p".into(), P_NS.into()),
                ("a".into(), A_NS.into()),
                ("r".into(), R_NS.into()),
                ("mc".into(), MC_NS.into()),
                ("xml".into(), "http://www.w3.org/XML/1998/namespace".into()),
            ])
            .with_start(&start)
        };
        let local_scope = bindings(self)?;
        let source_scope = bindings(master)?;
        let mut local_context = local_scope
            .entries()
            .into_iter()
            .collect::<std::collections::BTreeMap<_, _>>();
        let mut source_context = source_scope
            .entries()
            .into_iter()
            .collect::<std::collections::BTreeMap<_, _>>();
        local_context.entry(String::new()).or_default();
        source_context.entry(String::new()).or_default();
        let same_context = local_context == source_context;
        let identity = |name: &str, scope: &NamespaceBindings| -> Result<(String, String)> {
            let uri = if name.contains(':') {
                scope.attribute_uri(name.as_bytes()).ok_or_else(|| {
                    OxmlError::InvalidValue("unbound notes tree metadata attribute".into())
                })?
            } else {
                ""
            };
            Ok((uri.into(), name.rsplit(':').next().unwrap().into()))
        };
        for (key, value) in &master.raw_attributes {
            if key == "xmlns" || key.starts_with("xmlns:") {
                continue;
            }
            let source_identity = identity(key, &source_scope)?;
            if let Some((_, local)) = candidate
                .raw_attributes
                .iter()
                .filter(|(key, _)| key != "xmlns" && !key.starts_with("xmlns:"))
                .find(|(key, _)| {
                    identity(key, &local_scope).is_ok_and(|key| key == source_identity)
                })
            {
                if local != value || !same_context {
                    return Err(OxmlError::InvalidValue(
                        "conflicting expanded source notes-master tree metadata authority".into(),
                    ));
                }
                continue;
            }
            if !same_context && self != &Self::new() {
                return Err(OxmlError::InvalidValue(
                    "source notes-master root context would reinterpret local tree authority"
                        .into(),
                ));
            }
            candidate.raw_attributes.push((key.clone(), value.clone()));
        }
        if !same_context && self != &Self::new() {
            return Err(OxmlError::InvalidValue(
                "source notes-master root context would reinterpret local tree authority".into(),
            ));
        }
        for (key, value) in master
            .raw_attributes
            .iter()
            .filter(|(key, _)| key == "xmlns" || key.starts_with("xmlns:"))
        {
            if let Some((_, local)) = candidate
                .raw_attributes
                .iter()
                .find(|(name, _)| name == key)
            {
                if local != value {
                    return Err(OxmlError::InvalidValue(
                        "conflicting source notes-master tree metadata".into(),
                    ));
                }
            } else {
                candidate.raw_attributes.push((key.clone(), value.clone()));
            }
        }
        let mut matches = HashSet::new();
        let mut source_positions = Vec::with_capacity(master.children.len());
        let mut skipped_placeholder_ids = HashSet::new();
        for source in &master.children {
            let key = match source {
                ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref().map(CT_Placeholder::key),
                ShapeTreeChild::Picture(picture) => {
                    picture.placeholder.as_ref().map(CT_Placeholder::key)
                }
                _ => None,
            };
            let matching: Vec<_> = candidate
                .children
                .iter()
                .enumerate()
                .filter_map(|(index, child)| {
                    let local = match child {
                        ShapeTreeChild::Shape(shape) => {
                            shape.placeholder.as_ref().map(CT_Placeholder::key)
                        }
                        ShapeTreeChild::Picture(picture) => {
                            picture.placeholder.as_ref().map(CT_Placeholder::key)
                        }
                        _ => None,
                    };
                    match (&key, local) {
                        (Some(key), Some(local)) if key.matches(&local) => Some(index),
                        _ => None,
                    }
                })
                .collect();
            if matching.len() > 1 {
                return Err(OxmlError::InvalidValue(
                    "ambiguous imported notes placeholder".into(),
                ));
            }
            if let Some(&index) = matching.first() {
                source_positions.push(Some(index));
                if !matches.insert(index) {
                    return Err(OxmlError::InvalidValue(
                        "duplicate source notes-master placeholder".into(),
                    ));
                }
                match (&mut candidate.children[index], source) {
                    (ShapeTreeChild::Shape(local), ShapeTreeChild::Shape(parent)) => {
                        *local.shape_properties = local
                            .shape_properties
                            .with_inherited_defaults(&parent.shape_properties)
                            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
                        if local.raw.style.is_none() {
                            local.raw.style = parent.raw.style.clone();
                        }
                    }
                    (ShapeTreeChild::Picture(local), ShapeTreeChild::Picture(parent)) => {
                        local.shape_properties = local
                            .shape_properties
                            .with_inherited_defaults(&parent.shape_properties)
                            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
                    }
                    _ => {
                        return Err(OxmlError::InvalidValue(
                            "notes placeholder changes source member kind".into(),
                        ));
                    }
                }
            } else {
                if key.is_some() {
                    // A master hint is not an instantiated note placeholder.
                    source_positions.push(None);
                    if let Some(id) = source.non_visual_id() {
                        skipped_placeholder_ids.insert(id);
                    }
                } else {
                    source_positions.push(Some(candidate.children.len()));
                    candidate.append_child(source.clone());
                }
            }
        }
        let end = candidate.children.len();
        let retained_raw_children = candidate.raw_children.clone();
        let mut raw_match_cursors = HashMap::new();
        let mut appended_raw_boundaries = HashSet::new();
        for boundary in 0..=master.children.len() {
            for raw in master.raw_children.at(boundary) {
                let target = if boundary == master.children.len() {
                    end
                } else {
                    let next = source_positions[boundary..]
                        .iter()
                        .flatten()
                        .next()
                        .copied()
                        .unwrap_or(end);
                    let previous = source_positions[..boundary]
                        .iter()
                        .rev()
                        .flatten()
                        .next()
                        .copied();
                    if previous.is_some_and(|previous| previous >= next) {
                        return Err(OxmlError::InvalidValue(
                            "unrepresentable inherited notes raw boundary order".into(),
                        ));
                    }
                    next
                };
                let cursor = raw_match_cursors.get(&target).copied().unwrap_or(0);
                if same_context
                    && !appended_raw_boundaries.contains(&target)
                    && let Some((index, _)) = retained_raw_children
                        .at(target)
                        .enumerate()
                        .find(|(index, existing)| *index >= cursor && *existing == raw)
                {
                    raw_match_cursors.insert(target, index + 1);
                    continue;
                }
                let mut reader = Reader::from_reader(raw);
                let extension = match reader.read_event()? {
                    Event::Start(start) | Event::Empty(start) => {
                        let scope = source_scope.with_start(&start)?;
                        scope.element_uri(start.name().as_ref()) == Some(P_NS)
                            && local_name(start.name().as_ref()) == b"extLst"
                    }
                    _ => false,
                };
                if extension {
                    for at in 0..=end {
                        for existing in candidate.raw_children.at(at) {
                            let mut reader = Reader::from_reader(existing);
                            if let Event::Start(start) | Event::Empty(start) =
                                reader.read_event()?
                            {
                                let scope = local_scope.with_start(&start)?;
                                if scope.element_uri(start.name().as_ref()) == Some(P_NS)
                                    && local_name(start.name().as_ref()) == b"extLst"
                                {
                                    return Err(OxmlError::InvalidValue(
                                        "conflicting schema-final notes tree extension owners"
                                            .into(),
                                    ));
                                }
                            }
                        }
                    }
                }
                appended_raw_boundaries.insert(target);
                candidate.raw_children.push(target, raw.to_vec());
            }
        }
        for (local, inherited) in candidate
            .header_comments
            .iter_mut()
            .zip(&master.header_comments)
        {
            let original = local.clone();
            let mut cursor = 0;
            let mut appended = false;
            for comment in inherited {
                if !appended
                    && let Some(offset) = original[cursor..].iter().position(|raw| raw == comment)
                {
                    cursor += offset + 1;
                } else {
                    local.push(comment.clone());
                    appended = true;
                }
            }
        }
        if !skipped_placeholder_ids.is_empty()
            && shape_id_occurrences(&candidate.to_xml()?, NamespaceBindings::default())?
                .iter()
                .any(|reference| {
                    !reference.defines_shape && skipped_placeholder_ids.contains(&reference.id)
                })
        {
            return Err(OxmlError::InvalidValue(
                "retained member references an uninstantiated source notes-master placeholder"
                    .into(),
            ));
        }
        *self = candidate;
        Ok(())
    }

    pub fn take_children(&mut self, indices: &[usize]) -> Result<Vec<ShapeTreeChild>> {
        if indices.iter().any(|i| *i >= self.children.len())
            || indices.windows(2).any(|p| p[0] >= p[1])
        {
            return Err(OxmlError::InvalidValue(
                "selection indices are not ascending distinct members".to_owned(),
            ));
        }
        Ok(take_members(
            &mut self.children,
            &mut self.raw_children,
            indices,
        ))
    }
    pub fn insert_children(&mut self, index: usize, members: Vec<ShapeTreeChild>) -> Result<()> {
        if index > self.children.len() {
            return Err(OxmlError::InvalidValue(
                "member insertion is outside parent".to_owned(),
            ));
        }
        insert_members(&mut self.children, &mut self.raw_children, index, members);
        Ok(())
    }
}
impl CT_GroupShape {
    pub fn take_children(&mut self, indices: &[usize]) -> Result<Vec<ShapeTreeChild>> {
        if indices.iter().any(|i| *i >= self.children.len())
            || indices.windows(2).any(|p| p[0] >= p[1])
        {
            return Err(OxmlError::InvalidValue(
                "selection indices are not ascending distinct members".to_owned(),
            ));
        }
        Ok(take_members(
            &mut self.children,
            &mut self.raw_children,
            indices,
        ))
    }
    pub fn insert_children(&mut self, index: usize, members: Vec<ShapeTreeChild>) -> Result<()> {
        if index > self.children.len() {
            return Err(OxmlError::InvalidValue(
                "member insertion is outside parent".to_owned(),
            ));
        }
        insert_members(&mut self.children, &mut self.raw_children, index, members);
        Ok(())
    }
}
/// Removes typed members by ascending index. A raw child is recorded at the
/// index of the typed member it precedes, so each removal folds the boundary
/// after the member into the one before it.
fn take_members(
    children: &mut Vec<ShapeTreeChild>,
    raw_children: &mut OrderedRawChildren,
    indices: &[usize],
) -> Vec<ShapeTreeChild> {
    let mut taken = Vec::with_capacity(indices.len());
    for index in indices.iter().rev() {
        raw_children.close_boundary(index + 1);
        taken.push(children.remove(*index));
    }
    taken.reverse();
    taken
}

fn insert_members(
    children: &mut Vec<ShapeTreeChild>,
    raw_children: &mut OrderedRawChildren,
    index: usize,
    members: Vec<ShapeTreeChild>,
) {
    for (offset, member) in members.into_iter().enumerate() {
        raw_children.shift_boundaries_from(index + offset);
        children.insert(index + offset, member);
    }
}

// Selection replacement retains the raw source boundaries, not just typed indices.
impl CT_ShapeTree {
    pub fn group_member_span(
        &mut self,
        start: usize,
        count: usize,
        id: u32,
        name: &str,
    ) -> Result<()> {
        group_member_span(
            &mut self.children,
            &mut self.raw_children,
            &self.raw_attributes,
            start,
            count,
            id,
            name,
        )
    }
    pub fn replace_group_member(
        &mut self,
        index: usize,
        placed: Vec<ShapeTreeChild>,
        identity: bool,
    ) -> Result<()> {
        replace_group_member(
            &mut self.children,
            &mut self.raw_children,
            &self.raw_attributes,
            index,
            placed,
            identity,
        )
    }
}
impl CT_GroupShape {
    pub fn group_member_span(
        &mut self,
        start: usize,
        count: usize,
        id: u32,
        name: &str,
    ) -> Result<()> {
        group_member_span(
            &mut self.children,
            &mut self.raw_children,
            &self.raw_attributes,
            start,
            count,
            id,
            name,
        )
    }
    pub fn replace_group_member(
        &mut self,
        index: usize,
        placed: Vec<ShapeTreeChild>,
        identity: bool,
    ) -> Result<()> {
        replace_group_member(
            &mut self.children,
            &mut self.raw_children,
            &self.raw_attributes,
            index,
            placed,
            identity,
        )
    }
}

/// Add only inherited namespace declarations to a relocated raw root. All
/// original root spelling, attributes, entity spellings and inner bytes remain.
fn relocated_raw(
    xml: &[u8],
    attributes: &RawAttributes,
    destination: &RawAttributes,
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let (declared, end) = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) | Event::Empty(start) => {
                break (all_attributes(&start)?, reader.buffer_position() as usize);
            }
            Event::Eof => return Err(OxmlError::MissingElement("relocated raw root".to_owned())),
            _ => {}
        }
        buffer.clear();
    };
    let mut declarations = String::new();
    for (name, value) in attributes {
        if (name == "xmlns" || name.starts_with("xmlns:"))
            && !declared.iter().any(|(existing, _)| existing == name)
            && !destination
                .iter()
                .any(|(existing, uri)| existing == name && uri == value)
        {
            declarations.push(' ');
            declarations.push_str(name);
            declarations.push_str("=\"");
            declarations.push_str(&quick_xml::escape::escape(value));
            declarations.push('"');
        }
    }
    if declarations.is_empty() {
        return Ok(xml.to_vec());
    }
    let close = end
        .checked_sub(1)
        .filter(|at| xml.get(*at) == Some(&b'>'))
        .ok_or_else(|| OxmlError::InvalidValue("invalid relocated root span".to_owned()))?;
    let insert = if close > 0 && xml[close - 1] == b'/' {
        close - 1
    } else {
        close
    };
    let mut output = Vec::with_capacity(xml.len() + declarations.len());
    output.extend_from_slice(&xml[..insert]);
    output.extend_from_slice(declarations.as_bytes());
    output.extend_from_slice(&xml[insert..]);
    Ok(output)
}

fn replace_source_span(
    children: &mut Vec<ShapeTreeChild>,
    raw: &mut OrderedRawChildren,
    start: usize,
    count: usize,
    inserted: Vec<ShapeTreeChild>,
    inserted_raw: OrderedRawChildren,
) -> Result<()> {
    let end = start
        .checked_add(count)
        .filter(|end| *end <= children.len())
        .ok_or_else(|| OxmlError::InvalidValue("replacement span outside parent".to_owned()))?;
    let next_end = start
        .checked_add(inserted.len())
        .ok_or_else(|| OxmlError::InvalidValue("replacement boundary overflow".to_owned()))?;
    let mut next_raw = OrderedRawChildren::default();
    // The raw material preceding the replaced owner stays before its replacement.
    for boundary in 0..=start {
        for xml in raw.at(boundary) {
            next_raw.push(boundary, xml.to_vec());
        }
    }
    for boundary in 0..=inserted.len() {
        for xml in inserted_raw.at(boundary) {
            next_raw.push(start + boundary, xml.to_vec());
        }
    }
    for boundary in end..=children.len() {
        let destination = next_end
            .checked_add(boundary - end)
            .ok_or_else(|| OxmlError::InvalidValue("replacement boundary overflow".to_owned()))?;
        for xml in raw.at(boundary) {
            next_raw.push(destination, xml.to_vec());
        }
    }
    children.splice(start..end, inserted);
    *raw = next_raw;
    Ok(())
}

fn group_member_span(
    children: &mut Vec<ShapeTreeChild>,
    raw: &mut OrderedRawChildren,
    attributes: &RawAttributes,
    start: usize,
    count: usize,
    id: u32,
    name: &str,
) -> Result<()> {
    let end = start
        .checked_add(count)
        .filter(|end| count > 0 && *end <= children.len())
        .ok_or_else(|| OxmlError::InvalidValue("group span outside parent".to_owned()))?;
    let mut group = CT_GroupShape::new_empty(id, name);
    group.children = children[start..end].to_vec();
    group.raw_attributes = attributes
        .iter()
        .filter(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"))
        .cloned()
        .collect();
    // Only internal boundaries are enclosed. Outside neighbors stay outside.
    for boundary in start + 1..end {
        for xml in raw.at(boundary) {
            group.raw_children.push(boundary - start, xml.to_vec());
        }
    }
    replace_source_span(
        children,
        raw,
        start,
        count,
        vec![ShapeTreeChild::GroupShape(Box::new(group))],
        OrderedRawChildren::default(),
    )
}

fn has_namespace_only(attributes: &RawAttributes) -> bool {
    attributes
        .iter()
        .all(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"))
}

fn retained_owner_shell_is_default(group: &CT_GroupShape) -> Result<bool> {
    if !has_namespace_only(&group.raw_attributes)
        || !has_namespace_only(&group.non_visual_group_properties.raw_attributes)
        || !has_namespace_only(&group.group_properties.raw_attributes)
        || group.group_properties.effects.is_some()
        || !group.group_properties.raw_children.is_empty()
    {
        return Ok(false);
    }
    let mut inherited = vec![
        ("p".to_owned(), P_NS.to_owned()),
        ("a".to_owned(), A_NS.to_owned()),
        ("r".to_owned(), R_NS.to_owned()),
        ("mc".to_owned(), MC_NS.to_owned()),
    ];
    for (name, value) in group
        .raw_attributes
        .iter()
        .chain(&group.non_visual_group_properties.raw_attributes)
    {
        if name == "xmlns" {
            inherited.push((String::new(), value.clone()));
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            inherited.push((prefix.to_owned(), value.clone()));
        }
    }
    let bindings = NamespaceBindings::from_entries(&inherited);
    for xml in &group.non_visual_group_properties.raw_children {
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        let mut root = None;
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) | Event::Empty(start) if root.is_none() => {
                    if bindings
                        .with_start(&start)?
                        .element_uri(start.name().as_ref())
                        != Some(P_NS)
                    {
                        return Ok(false);
                    }
                    let allowed: &[&str] = match local_name(start.name().as_ref()) {
                        b"cNvPr" => &["id", "name", "descr", "title"],
                        b"cNvGrpSpPr" | b"nvPr" => &[],
                        _ => return Ok(false),
                    };
                    if !all_attributes(&start)?.iter().all(|(name, _)| {
                        name == "xmlns"
                            || name.starts_with("xmlns:")
                            || allowed.contains(&name.as_str())
                    }) {
                        return Ok(false);
                    }
                    root = Some(start.name().as_ref().to_vec());
                }
                Event::End(end) if root.as_deref() == Some(end.name().as_ref()) => {}
                Event::Text(text) if text.iter().all(u8::is_ascii_whitespace) => {}
                Event::Eof => break,
                // Unknown shell bodies belong to the removed group. They cannot
                // be transplanted to its parent or silently discarded.
                _ => return Ok(false),
            }
            buffer.clear();
        }
    }
    Ok(true)
}

fn replace_group_member(
    children: &mut Vec<ShapeTreeChild>,
    raw: &mut OrderedRawChildren,
    parent_attributes: &RawAttributes,
    index: usize,
    placed: Vec<ShapeTreeChild>,
    identity: bool,
) -> Result<()> {
    let Some(ShapeTreeChild::GroupShape(group)) = children.get(index) else {
        return Err(OxmlError::InvalidValue(
            "ungroup owner is not a group".to_owned(),
        ));
    };
    if placed.len() != group.children.len() {
        return Err(OxmlError::InvalidValue(
            "ungroup child count changed".to_owned(),
        ));
    }
    if !retained_owner_shell_is_default(group)? {
        return Err(OxmlError::InvalidValue("ungroup cannot preserve unsupported group-owned metadata or effects in a different schema owner".to_owned()));
    }
    let mut moved = OrderedRawChildren::default();
    for boundary in 0..=group.children.len() {
        for xml in group.raw_children.at(boundary) {
            let mut reader = Reader::from_reader(xml);
            let mut buffer = Vec::new();
            let owner_only = loop {
                match reader.read_event_into(&mut buffer)? {
                    Event::Start(start) | Event::Empty(start) => {
                        let mut inherited = vec![
                            ("p".to_owned(), P_NS.to_owned()),
                            ("a".to_owned(), A_NS.to_owned()),
                            ("r".to_owned(), R_NS.to_owned()),
                            ("mc".to_owned(), MC_NS.to_owned()),
                        ];
                        inherited.extend(group.raw_attributes.iter().filter_map(
                            |(name, value)| {
                                if name == "xmlns" {
                                    Some((String::new(), value.clone()))
                                } else {
                                    name.strip_prefix("xmlns:")
                                        .map(|prefix| (prefix.to_owned(), value.clone()))
                                }
                            },
                        ));
                        let bindings =
                            NamespaceBindings::from_entries(&inherited).with_start(&start)?;
                        break local_name(start.name().as_ref()) == b"extLst"
                            && bindings.element_uri(start.name().as_ref()) == Some(P_NS);
                    }
                    Event::Eof => {
                        return Err(OxmlError::MissingElement("raw group member".to_owned()));
                    }
                    _ => {}
                }
                buffer.clear();
            };
            if owner_only {
                return Err(OxmlError::InvalidValue(
                    "ungroup cannot relocate group-owned p:extLst to a different schema owner"
                        .to_owned(),
                ));
            }
            if !identity {
                return Err(OxmlError::InvalidValue(
                    "ungroup cannot transform an unmodeled member under a nonidentity group map"
                        .to_owned(),
                ));
            }
            moved.push(
                boundary,
                relocated_raw(xml, &group.raw_attributes, parent_attributes)?,
            );
        }
    }
    replace_source_span(children, raw, index, 1, placed, moved)
}

#[cfg(test)]
mod style_tests {
    use oxml_core::units::Emu;
    use oxml_drawing::style_ref::FontCollectionIndex;
    use oxml_drawing::xfrm::{CT_Point2D, CT_PositiveSize2D, CT_Transform2D};

    use super::{CT_GroupShape, CT_Shape, CT_ShapeTree, ShapeIdAllocator, ShapeTreeChild};

    const ALLOCATOR_TREE: &[u8] = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Root shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="6" name="Nested picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:grpSp><mc:AlternateContent><mc:Fallback><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="8" name="Fallback connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></mc:Fallback></mc:AlternateContent></p:spTree>"#;

    #[test]
    fn fx206_shape_save_reread_keeps_drawing_text_under_presentation_owner() {
        let mut shape = CT_Shape::new_textbox(2, "Body", CT_Transform2D::default()).unwrap();
        shape.text_body = Some(oxml_drawing::text::CT_TextBody::from_xml(br#"<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>shape text survives</a:t></a:r></a:p></a:txBody>"#).unwrap());
        let saved = shape.to_xml().unwrap();
        let reopened = CT_Shape::from_xml(&saved).unwrap();
        assert_eq!(
            reopened.text_body.as_ref().unwrap().plain_text(),
            "shape text survives"
        );
        assert_eq!(reopened.to_xml().unwrap(), saved);
        let xml = std::str::from_utf8(&saved).unwrap();
        assert!(xml.contains("<p:txBody"));
        assert!(!xml.contains("<a:txBody"));
    }

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
        assert!(shape_xml.contains("<p:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"));
        assert!(shape_xml.find("<p:nvSpPr").unwrap() < shape_xml.find("<p:spPr").unwrap());
        assert!(shape_xml.find("<p:spPr").unwrap() < shape_xml.find("<p:txBody").unwrap());

        let textbox_xml = String::from_utf8(textbox.to_xml().unwrap()).unwrap();
        assert!(textbox_xml.contains("<p:cNvSpPr txBox=\"1\"/>"));
        assert!(textbox_xml.contains("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"));
        assert!(textbox_xml.contains("<a:noFill/>"));
        assert!(textbox_xml.contains("<p:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"));
        assert_eq!(CT_Shape::from_xml(textbox_xml.as_bytes()).unwrap(), textbox);
        assert!(shape.style().is_none());
        assert!(textbox.style().is_none());
    }

    #[test]
    fn default_style_follows_shape_properties_and_has_a_typed_view() {
        let mut shape =
            CT_Shape::new_preset(2, "Shape 2", "rect", CT_Transform2D::default()).unwrap();
        shape.set_default_style().unwrap();

        let style = shape.style().expect("typed default style");
        assert_eq!(style.line_reference.index, 1);
        assert_eq!(style.fill_reference.index, 3);
        assert_eq!(style.effect_reference.index, 2);
        assert_eq!(style.font_reference.index, FontCollectionIndex::Minor);
        let xml = shape.to_xml().unwrap();
        let text = String::from_utf8(xml.clone()).unwrap();
        assert!(text.contains(
            r#"</p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#
        ));
        assert_eq!(CT_Shape::from_xml(&xml).unwrap(), shape);
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
    fn group_members_append_right_after_the_last_typed_member() {
        let xml = br#"<p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="5" name="First"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:contentPart r:id="rId9"/><p:extLst><p:ext uri="{kept}"/></p:extLst></p:grpSp>"#;
        let mut group = CT_GroupShape::from_xml(xml).unwrap();
        let appended = group.append_child(ShapeTreeChild::GroupShape(Box::new(
            CT_GroupShape::new_empty(6, "Nested"),
        )));
        assert_eq!(appended.non_visual_id(), Some(6));

        let written = String::from_utf8(group.to_xml().unwrap()).unwrap();
        let first = written.find(r#"id="5""#).unwrap();
        let nested = written.find(r#"id="6""#).unwrap();
        let ink = written.find("<p:contentPart").unwrap();
        let extension = written.find("<p:extLst>").unwrap();
        assert!(
            first < nested && nested < ink && ink < extension,
            "{written}"
        );
        assert_eq!(CT_GroupShape::from_xml(written.as_bytes()).unwrap(), group);
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
    fn shape_id_allocator_occupies_a_typed_tree_id() {
        let xml = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvGrpSpPr><p:cNvPr id="4" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>"#;
        let tree = CT_ShapeTree::from_xml(xml).unwrap();
        let mut allocator = ShapeIdAllocator::scan(&tree);

        assert_eq!(allocator.allocate(), 2);
        assert_eq!(allocator.allocate(), 3);
        assert_eq!(allocator.allocate(), 5);
    }

    #[test]
    fn shape_id_allocator_uses_typed_ids_for_textboxes() {
        let mut tree = CT_ShapeTree::new();
        let first = CT_Shape::new_textbox(2, "TextBox 2", CT_Transform2D::default()).unwrap();
        tree.append_child(ShapeTreeChild::Shape(first));
        let second = CT_Shape::new_textbox(4, "TextBox 4", CT_Transform2D::default()).unwrap();
        tree.append_child(ShapeTreeChild::Shape(second));
        let mut allocator = ShapeIdAllocator::scan(&tree);

        assert_eq!(allocator.allocate(), 3);
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
    fn fx206_foreign_relationship_shape_and_group_writers_keep_one_binding() {
        fn foreign_payload_count(xml: &[u8]) -> usize {
            let mut reader = quick_xml::reader::NsReader::from_reader(xml);
            let mut buffer = Vec::new();
            let mut count = 0;
            loop {
                let (namespace, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
                match event {
                    quick_xml::events::Event::Start(start)
                    | quick_xml::events::Event::Empty(start) => {
                        // Duplicate namespace attributes are an invalid document,
                        // even when the parser could otherwise choose one binding.
                        for attribute in start.attributes() {
                            attribute.unwrap();
                        }
                        if start.local_name().as_ref() == b"payload" {
                            assert!(matches!(namespace,
                                quick_xml::name::ResolveResult::Bound(uri)
                                if uri.as_ref() == b"urn:foreign"));
                            count += 1;
                        }
                    }
                    quick_xml::events::Event::Eof => break,
                    _ => {}
                }
                buffer.clear();
            }
            count
        }

        let raw = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="urn:foreign"><p:nvSpPr><p:cNvPr id="19" name="Imported"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><r:payload value='&#65;'><r:opaque/></r:payload></p:sp>"#;
        let shape = CT_Shape::from_xml(raw).unwrap();
        let saved = shape.to_xml().unwrap();
        assert_eq!(foreign_payload_count(&saved), 1);
        assert!(String::from_utf8_lossy(&saved).contains("value='&#65;'"));
        assert_eq!(CT_Shape::from_xml(&saved).unwrap().to_xml().unwrap(), saved);

        // The child has no physical declaration. Its foreign r ownership comes
        // from the group and must survive both embedded and standalone writes.
        let raw = br#"<p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="urn:foreign"><p:nvGrpSpPr><p:cNvPr id="20" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="21" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><r:payload value='&#66;'/></p:sp><r:payload value='&#67;'/></p:grpSp>"#;
        let group = CT_GroupShape::from_xml(raw).unwrap();
        let saved = group.to_xml().unwrap();
        assert_eq!(foreign_payload_count(&saved), 2);
        assert_eq!(
            CT_GroupShape::from_xml(&saved).unwrap().to_xml().unwrap(),
            saved
        );
        let ShapeTreeChild::Shape(child) = &group.children[0] else {
            panic!("expected original child shape");
        };
        let child_saved = child.to_xml().unwrap();
        assert_eq!(foreign_payload_count(&child_saved), 1);
        assert_eq!(
            CT_Shape::from_xml(&child_saved).unwrap().to_xml().unwrap(),
            child_saved
        );

        // Only r is a preserved opaque owner. Genuine conflicting p/a/mc
        // declarations are still rejected, rather than broadly admitted.
        for prefix in ["p", "a", "mc"] {
            let source = format!(
                "<q:sp xmlns:q=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:{prefix}=\"urn:foreign\"><q:nvSpPr><q:cNvPr id=\"22\" name=\"Conflict\"/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr/></q:sp>"
            );
            assert!(CT_Shape::from_xml(source.as_bytes()).is_err(), "{prefix}");
        }
    }

    #[test]
    fn fx206_foreign_relationship_picture_placeholder_refuses_without_source_loss() {
        let raw = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="urn:foreign"><p:nvSpPr><p:cNvPr id="23" name="Picture"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="4"/></p:nvPr></p:nvSpPr><p:spPr/><r:payload value='&#68;'/></p:sp>"#;
        let shape = CT_Shape::from_xml(raw).unwrap();
        let before = shape.to_xml().unwrap();
        let error = shape
            .fill_picture("rIdImage", CT_Transform2D::default())
            .unwrap_err();
        assert!(matches!(error, oxml_core::OxmlError::InvalidValue(message)
            if message == "fill picture cannot preserve a foreign root relationship binding"));
        assert_eq!(shape.to_xml().unwrap(), before);
        assert_eq!(
            CT_Shape::from_xml(&before).unwrap().to_xml().unwrap(),
            before
        );

        let canonical = String::from_utf8(raw.to_vec()).unwrap().replace(
            "xmlns:r=\"urn:foreign\"",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        );
        let shape = CT_Shape::from_xml(canonical.as_bytes()).unwrap();
        let picture = shape
            .fill_picture("rIdImage", CT_Transform2D::default())
            .unwrap();
        let saved = picture.to_xml().unwrap();
        assert_eq!(
            crate::picture::CT_Picture::from_xml(&saved)
                .unwrap()
                .to_xml()
                .unwrap(),
            saved
        );
    }

    #[test]
    fn fx206_picture_placeholder_opaque_empty_text_body_is_atomic() {
        let raw = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="19" name="Imported"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="4"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody xmlns:x="urn:producer" x:keep='&#65;'><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#;
        let shape = CT_Shape::from_xml(raw).unwrap();
        let before = shape.to_xml().unwrap();
        assert!(
            std::str::from_utf8(&before)
                .unwrap()
                .contains("x:keep='&#65;'")
        );
        assert_eq!(
            CT_Shape::from_xml(&before).unwrap().to_xml().unwrap(),
            before
        );
        assert!(
            matches!(shape.fill_picture("rIdImage", CT_Transform2D::default()),
            Err(oxml_core::OxmlError::InvalidValue(message)) if message == "picture placeholder text body contains source that a picture cannot retain"),
            "FX206_OPAQUE_EMPTY_TEXT_BODY_MUST_REFUSE"
        );
        assert_eq!(shape.to_xml().unwrap(), before);
        assert_eq!(shape.non_visual_id(), Some(19));
    }

    #[test]
    fn fx208_placeholder_identity_clone_respects_actual_namespace_ownership() {
        let source = "<q:sp xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:f='urn:foreign'><q:nvSpPr><q:cNvPr id='19' name='Imported' f:id='opaque'><f:payload id='44'/></q:cNvPr><q:cNvSpPr/><q:nvPr><q:ph type='ftr' idx='4'/></q:nvPr></q:nvSpPr><q:spPr/></q:sp>";
        let shape = CT_Shape::from_xml(source.as_bytes()).unwrap();
        let before = shape.to_xml().unwrap();
        let cloned = shape.clone_with_id(77).unwrap();
        let saved = cloned.to_xml().unwrap();
        assert_eq!(cloned.non_visual_id(), Some(77));
        assert_eq!(cloned.placeholder, shape.placeholder);
        assert!(
            std::str::from_utf8(&saved)
                .unwrap()
                .contains("f:id='opaque'")
        );
        assert!(
            std::str::from_utf8(&saved)
                .unwrap()
                .contains("<f:payload id='44'/>")
        );
        assert_eq!(shape.to_xml().unwrap(), before);
        assert_eq!(
            CT_Shape::from_xml(&saved).unwrap().non_visual_id(),
            Some(77)
        );
        for invalid in [
            source
                .replace("<q:cNvPr", "<f:cNvPr")
                .replace("</q:cNvPr>", "</f:cNvPr>"),
            source
                .replace("<q:cNvPr", "<u:cNvPr")
                .replace("</q:cNvPr>", "</u:cNvPr>"),
            source.replace("<q:cNvPr", "<q:cNvPr xmlns:q='urn:foreign'"),
        ] {
            assert!(CT_Shape::from_xml(invalid.as_bytes()).is_err());
        }
    }

    #[test]
    fn fx206_picture_placeholder_empty_namespace_aliases_remain_supported() {
        for body in [
            "<p:txBody xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>",
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main'><d:bodyPr/><d:lstStyle/><d:p/></q:txBody>",
            "<txBody xmlns='http://schemas.openxmlformats.org/presentationml/2006/main'><a:bodyPr/><a:lstStyle/><a:p/></txBody>",
        ] {
            let raw = format!(
                "<p:sp xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><p:nvSpPr><p:cNvPr id='19' name='Imported'/><p:cNvSpPr/><p:nvPr><p:ph type='pic' idx='4'/></p:nvPr></p:nvSpPr><p:spPr/>{body}</p:sp>"
            );
            let shape = CT_Shape::from_xml(raw.as_bytes()).unwrap();
            let before = shape.to_xml().unwrap();
            let converted = shape
                .fill_picture("rIdImage", CT_Transform2D::default())
                .unwrap_or_else(|error| panic!("{body}: {error}"));
            assert_eq!(converted.non_visual_id(), Some(19));
            assert_eq!(converted.placeholder, shape.placeholder);
            assert_eq!(shape.to_xml().unwrap(), before);
            let saved = converted.to_xml().unwrap();
            assert_eq!(
                super::CT_Picture::from_xml(&saved)
                    .unwrap()
                    .to_xml()
                    .unwrap(),
                saved
            );
        }
    }

    #[test]
    fn fx209_picture_placeholder_empty_shell_refuses_owned_or_opaque_source() {
        let (p_ns, a_ns) = (super::P_NS, super::A_NS);
        for body in [
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main' keep='&#65;'><d:bodyPr/><d:lstStyle/><d:p/></q:txBody>",
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main'><d:bodyPr lIns='100'/><d:lstStyle/><d:p/></q:txBody>",
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main'><d:bodyPr/><d:lstStyle xmlns:d='urn:foreign'/><d:p/></q:txBody>",
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:e='urn:producer'><d:bodyPr/><d:lstStyle/><d:p e:keep='&#65;'/></q:txBody>",
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:e='urn:producer'><d:bodyPr/><d:lstStyle/><e:opaque/><d:p/></q:txBody>",
            "<q:txBody xmlns:q='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:d='http://schemas.openxmlformats.org/drawingml/2006/main'><d:bodyPr/><d:lstStyle/><d:p><d:r><d:t>text</d:t></d:r></d:p></q:txBody>",
        ] {
            let raw = format!(
                "<p:sp xmlns:p='{p_ns}' xmlns:a='{a_ns}'><p:nvSpPr><p:cNvPr id='19' name='Imported'/><p:cNvSpPr/><p:nvPr><p:ph type='pic' idx='4'/></p:nvPr></p:nvSpPr><p:spPr/>{body}</p:sp>"
            );
            let shape = CT_Shape::from_xml(raw.as_bytes()).unwrap();
            let before = shape.to_xml().unwrap();
            assert!(
                matches!(shape.fill_picture("rIdImage", CT_Transform2D::default()),
                Err(oxml_core::OxmlError::InvalidValue(message)) if message == "picture placeholder text body contains source that a picture cannot retain"),
                "F209_EMPTY_SHELL_SOURCE_REFUSED {body}"
            );
            assert_eq!(shape.to_xml().unwrap(), before);
        }
    }

    #[test]
    fn fx206_picture_placeholder_conversion_preserves_shared_locks_and_foreign_metadata() {
        let raw = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:f="urn:foreign"><p:nvSpPr f:nv='kept'><p:cNvPr id='19' name='Imported' descr='&#65;'><f:meta data='&#66;'/></p:cNvPr><p:cNvSpPr f:keep='&#67;'><d:spLocks noGrp='1' noResize='true' f:keep='&#68;'><f:opaque/></d:spLocks><f:spLocks noTextEdit='1'/></p:cNvSpPr><p:nvPr f:app='kept'><f:before/><p:ph type='pic' idx='4'/><f:after/></p:nvPr></p:nvSpPr><p:spPr/><f:tail/></p:sp>"#;
        let shape = CT_Shape::from_xml(raw).unwrap();
        let before = shape.to_xml().unwrap();
        let transform = CT_Transform2D::from_xml(
            br#"<a:xfrm><a:off x="12700" y="25400"/><a:ext cx="1270000" cy="635000"/></a:xfrm>"#,
        )
        .unwrap();
        let picture = shape.fill_picture("rIdImage", transform).unwrap();
        assert_eq!(shape.to_xml().unwrap(), before);
        assert_eq!(picture.placeholder, shape.placeholder);
        assert_eq!(picture.non_visual_id(), Some(19));
        assert_eq!(picture.non_visual_name(), Some("Imported"));
        assert!(picture.shape_properties.custom_geometry.is_none());
        assert_eq!(
            picture
                .shape_properties
                .preset_geometry
                .as_ref()
                .unwrap()
                .preset,
            "rect"
        );
        let saved = picture.to_xml().unwrap();
        let xml = std::str::from_utf8(&saved).unwrap();
        for token in [
            "<p:cNvPr id='19' name='Imported' descr='&#65;'><f:meta data='&#66;'/></p:cNvPr>",
            "<d:picLocks noGrp='1' noResize='true' f:keep='&#68;'><f:opaque/></d:picLocks>",
            "<f:spLocks noTextEdit='1'/>",
            "<f:before/>",
            "<f:after/>",
            "<f:tail/>",
        ] {
            assert!(xml.contains(token), "{token}");
        }
        assert!(!xml.contains("<d:spLocks"));
        assert!(xml.find("nvPicPr").unwrap() < xml.find("blipFill").unwrap());
        assert!(xml.find("blipFill").unwrap() < xml.find("<p:spPr").unwrap());
        assert_eq!(xml.matches("<a:prstGeom prst=\"rect\">").count(), 1);
        assert!(xml.find("<a:xfrm").unwrap() < xml.find("<a:prstGeom").unwrap());
        assert_eq!(
            super::CT_Picture::from_xml(&saved)
                .unwrap()
                .to_xml()
                .unwrap(),
            saved
        );
        for geometry in [
            "<a:prstGeom prst='trapezoid'><a:avLst/><f:geometry code='&#69;'/></a:prstGeom>",
            "<a:custGeom><a:pathLst><a:path w='100' h='100'><a:moveTo><a:pt x='0' y='0'/></a:moveTo><a:lnTo><a:pt x='100' y='100'/></a:lnTo></a:path></a:pathLst><f:geometry code='&#69;'/></a:custGeom>",
        ] {
            let source = String::from_utf8(raw.to_vec()).unwrap().replace("<p:spPr/>", &format!("<p:spPr f:mode='kept'><f:before code='&#70;'/>{geometry}<f:after code='&#71;'/></p:spPr>"));
            let imported = CT_Shape::from_xml(source.as_bytes()).unwrap();
            let unchanged = imported.to_xml().unwrap();
            let transform = picture.shape_properties.transform.clone().unwrap();
            let mut expected = (*imported.shape_properties).clone();
            expected.transform = Some(transform.clone());
            let converted = imported.fill_picture("rIdImage", transform).unwrap();
            assert_eq!(converted.shape_properties, expected);
            assert_eq!(imported.to_xml().unwrap(), unchanged);
            let saved = converted.to_xml().unwrap();
            let xml = std::str::from_utf8(&saved).unwrap();
            assert!(!xml.contains("prst=\"rect\""));
            for raw in [
                "<f:before code='&#70;'/>",
                "<f:geometry code='&#69;'/>",
                "<f:after code='&#71;'/>",
            ] {
                assert!(xml.contains(raw), "{raw}");
            }
            assert_eq!(
                super::CT_Picture::from_xml(&saved)
                    .unwrap()
                    .to_xml()
                    .unwrap(),
                saved
            );
        }
        for owned in ["txBox='1'", "txBox='0'"] {
            let source = String::from_utf8(raw.to_vec())
                .unwrap()
                .replace("<p:cNvSpPr f:keep", &format!("<p:cNvSpPr {owned} f:keep"));
            let refused = CT_Shape::from_xml(source.as_bytes()).unwrap();
            let before = refused.to_xml().unwrap();
            assert!(
                refused
                    .fill_picture("rIdImage", CT_Transform2D::default())
                    .is_err()
            );
            assert_eq!(refused.to_xml().unwrap(), before);
        }
        let source = String::from_utf8(raw.to_vec())
            .unwrap()
            .replace("noResize='true'", "noTextEdit='true'");
        let refused = CT_Shape::from_xml(source.as_bytes()).unwrap();
        assert!(
            refused
                .fill_picture("rIdImage", CT_Transform2D::default())
                .is_err()
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

    #[test]
    fn moving_a_child_passes_only_the_children_between_its_indices() {
        let shape = |id: u32| {
            format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="S{id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"#
            )
        };
        let xml = format!(
            r#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}<x:ink xmlns:x="urn:producer" x:kept="between"/>{}{}<p:extLst><p:ext uri="{{kept}}"/></p:extLst></p:spTree>"#,
            shape(2),
            shape(3),
            shape(4)
        );
        let mut tree = CT_ShapeTree::from_xml(xml.as_bytes()).unwrap();
        let ids = |tree: &CT_ShapeTree| {
            tree.children
                .iter()
                .filter_map(ShapeTreeChild::non_visual_id)
                .collect::<Vec<_>>()
        };
        let order = |tree: &CT_ShapeTree| {
            let text = String::from_utf8(tree.to_xml().unwrap()).unwrap();
            let mut marks = [r#"id="2""#, r#"id="3""#, r#"id="4""#, "<x:ink", "<p:extLst"]
                .map(|mark| (text.find(mark).unwrap(), mark));
            marks.sort();
            marks.map(|(_, mark)| mark)
        };

        tree.move_child(0, 2).unwrap();
        assert_eq!(ids(&tree), [3, 4, 2]);
        assert_eq!(
            order(&tree),
            ["<x:ink", r#"id="3""#, r#"id="4""#, r#"id="2""#, "<p:extLst"]
        );
        tree.move_child(2, 1).unwrap();
        assert_eq!(ids(&tree), [3, 2, 4]);
        tree.move_child(1, 0).unwrap();
        assert_eq!(
            order(&tree),
            ["<x:ink", r#"id="2""#, r#"id="3""#, r#"id="4""#, "<p:extLst"]
        );
        let before = tree.to_xml().unwrap();
        tree.move_child(1, 1).unwrap();
        assert!(tree.move_child(0, 3).is_err());
        assert!(tree.move_child(3, 0).is_err());
        assert_eq!(tree.to_xml().unwrap(), before);
    }

    #[test]
    fn effect_reference_index_changes_only_a_present_style() {
        let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>"#;
        let mut shape = CT_Shape::from_xml(xml).unwrap();

        assert!(shape.set_effect_reference_index(0));

        let written = String::from_utf8(shape.to_xml().unwrap()).unwrap();
        assert!(written.contains(
            r#"<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>"#
        ));
        let mut unstyled = CT_Shape::new_preset(3, "Shape 3", "rect", Default::default()).unwrap();
        let before = unstyled.clone();
        assert!(!unstyled.set_effect_reference_index(0));
        assert_eq!(unstyled, before);
    }
}

#[cfg(test)]
mod fx214_frame_controls {
    use super::*;
    const ROOT: &str = r#"<q:spTree xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:foreign"><!--before--><q:nvGrpSpPr><q:cNvPr id="1"/><q:cNvGrpSpPr/><q:nvPr/></q:nvGrpSpPr><?preserve after-nonvisual?><q:grpSpPr/>"#;
    fn owner(children: &str) -> String {
        format!("{ROOT}{children}</q:spTree>")
    }
    fn alternate(depth: usize, payload: &str) -> String {
        format!(
            "{}{}{}",
            "<mc:AlternateContent><mc:Fallback>".repeat(depth),
            payload,
            "</mc:Fallback></mc:AlternateContent>".repeat(depth)
        )
    }
    #[test]
    fn fx214_qualified_empty_mixed_and_foreign_owners_preserve_source() {
        let payload = r#"<f:grpSp f:keep='&#65;'><f:grpSp/></f:grpSp><!--between--><f:opaque/>"#;
        let source = owner(payload);
        let parsed = CT_ShapeTree::from_xml(source.as_bytes()).unwrap();
        let saved = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(saved.contains(payload));
        assert!(saved.contains("<!--before-->"));
        assert!(saved.contains("<?preserve after-nonvisual?>"));
        let scope = owner(&alternate(32, payload));
        let parsed = CT_ShapeTree::from_xml(scope.as_bytes()).unwrap();
        assert_eq!(validate_shape_nesting(&parsed.children).unwrap(), 32);
        assert!(
            String::from_utf8(parsed.to_xml().unwrap())
                .unwrap()
                .contains(&alternate(32, payload))
        );
        let source = owner(&alternate(32, "<mc:AlternateContent/>"));
        let error = CT_ShapeTree::from_xml(source.as_bytes())
            .unwrap_err()
            .to_string();
        assert!(error.contains("limit 32, offending depth 33"), "{error}");
        let rebound = owner(r#"<q:grpSp xmlns:q="urn:foreign"><q:grpSp/></q:grpSp>"#);
        assert!(
            CT_ShapeTree::from_xml(rebound.as_bytes())
                .unwrap()
                .children
                .is_empty()
        );
    }
    #[test]
    fn fx214_cursor_work_and_shared_compatibility_source_are_linear() {
        let leaf = r#"<f:opaque f:keep='&#65;'><!--leaf--></f:opaque>"#;
        let input = |breadth| owner(&alternate(32, &leaf.repeat(breadth)));
        let one = input(1);
        let small = input(64);
        let large = input(128);
        let (_, base_events, base_bytes) = parse_shape_frames(
            one.as_bytes(),
            &[],
            ShapeParseRoot::Group(GroupKind::ShapeTree),
        )
        .unwrap();
        let (small_owner, small_events, small_bytes) = parse_shape_frames(
            small.as_bytes(),
            &[],
            ShapeParseRoot::Group(GroupKind::ShapeTree),
        )
        .unwrap();
        let (_, large_events, large_bytes) = parse_shape_frames(
            large.as_bytes(),
            &[],
            ShapeParseRoot::Group(GroupKind::ShapeTree),
        )
        .unwrap();
        assert_eq!(
            large_events - base_events,
            (small_events - base_events) * 127 / 63
        );
        assert_eq!(
            large_bytes - base_bytes,
            (small_bytes - base_bytes) * 127 / 63
        );
        let ParsedShapeOwner::Group(group) = small_owner else {
            panic!("group root");
        };
        let ShapeTreeChild::AlternateContent(first) = &group.children[0] else {
            panic!("AC root");
        };
        let mut current = first.as_ref();
        let mut count = 1;
        while let Some(ShapeTreeChild::AlternateContent(next)) =
            current.selected_fallback().and_then(|v| v.first())
        {
            assert!(std::sync::Arc::ptr_eq(&first.raw_source, &next.raw_source));
            assert_eq!(next.raw_source.len(), small.len());
            current = next;
            count += 1;
        }
        assert_eq!(count, 32);
        assert_eq!(first.raw_xml(), alternate(32, &leaf.repeat(64)).as_bytes());
    }
}

#[cfg(test)]
#[test]
fn fx216_style_materialization_prepares_actual_scopes_and_neutralizes_last() {
    let mut shape = CT_Shape::new_preset(2, "source", "rect", CT_Transform2D::default()).unwrap();
    shape.set_default_style().unwrap();
    let mut xml = String::from_utf8(shape.to_xml().unwrap()).unwrap();
    xml = xml
        .replace(
            "<p:style>",
            r#"<p:style xmlns:z="urn:source" xmlns:f="urn:foreign" z:meaning="z:item">"#,
        )
        .replace(
            "<a:fillRef idx=\"3\">",
            "<f:fillRef idx='99' keep='&#65;'/><a:fillRef idx=\"3\">",
        )
        .replace("</a:fillRef>", "<z:payload meaning='z:item'/></a:fillRef>");
    let mut shape = CT_Shape::from_xml(xml.as_bytes()).unwrap();
    shape.set_effect_reference_index(1);
    let before = shape.to_xml().unwrap();
    let prepared = shape.style_for_materialization().unwrap().unwrap();
    assert_eq!(prepared.line_reference.index, 1);
    assert_eq!(prepared.fill_reference.index, 3);
    assert_eq!(prepared.effect_reference.index, 1);
    assert_eq!(
        prepared.font_reference.index,
        oxml_drawing::style_ref::FontCollectionIndex::Minor
    );
    assert_eq!(shape.to_xml().unwrap(), before);
    shape.shape_properties.fill = Some(
        oxml_drawing::fill::Fill::from_xml(
            br#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
        )
        .unwrap(),
    );
    shape.materialize_style_references().unwrap();
    let style = shape.style().unwrap();
    assert_eq!(
        (
            style.line_reference.index,
            style.fill_reference.index,
            style.effect_reference.index
        ),
        (0, 0, 0)
    );
    assert_eq!(
        style.font_reference.index,
        oxml_drawing::style_ref::FontCollectionIndex::None
    );
    let xml = shape.to_xml().unwrap();
    assert!(String::from_utf8_lossy(&xml).contains("idx='99' keep='&#65;'"));
    assert!(String::from_utf8_lossy(&xml).contains("meaning='z:item'"));
    let mut reader = quick_xml::NsReader::from_reader(xml.as_slice());
    let mut payload = false;
    loop {
        let (namespace, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Empty(start) if start.local_name().as_ref() == b"payload" => {
                assert!(
                    matches!(namespace, quick_xml::name::ResolveResult::Bound(uri) if uri.as_ref() == b"urn:source")
                );
                payload = true;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    assert!(payload);
    assert_eq!(CT_Shape::from_xml(&xml).unwrap().to_xml().unwrap(), xml);
    for bad in [
        DEFAULT_STYLE.replace(
            "<a:fillRef idx=\"3\">",
            "<a:fillRef idx=\"3\" z:meaning=\"z:item\">",
        ),
        DEFAULT_STYLE.replace(
            "<a:schemeClr val=\"accent1\"/>",
            "<f:schemeClr val=\"accent1\"/>",
        ),
    ] {
        let mut original =
            CT_Shape::new_preset(2, "negative", "rect", CT_Transform2D::default()).unwrap();
        original.set_default_style().unwrap();
        let source = String::from_utf8(original.to_xml().unwrap())
            .unwrap()
            .replace(
                DEFAULT_STYLE,
                &bad.replace(
                    "<p:style>",
                    r#"<p:style xmlns:z="urn:source" xmlns:f="urn:foreign">"#,
                ),
            );
        let mut original = CT_Shape::from_xml(source.as_bytes()).unwrap();
        let before = original.to_xml().unwrap();
        assert!(original.style_for_materialization().is_err());
        assert!(original.materialize_style_references().is_err());
        assert_eq!(original.to_xml().unwrap(), before);
    }
}

#[test]
fn fx216_notes_tree_refuses_local_opaque_default_namespace_rebinding() {
    let tree = |attrs: &str, raw: &str| {
        CT_ShapeTree::from_xml(format!(r#"<p:spTree xmlns:p="{P_NS}" {attrs}><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{raw}</p:spTree>"#).as_bytes()).unwrap()
    };
    let mut local = tree(
        "",
        "<p:extLst><p:ext uri=\"local\"><opaque/></p:ext></p:extLst>",
    );
    let master = tree("xmlns=\"urn:source\"", "");
    let before = local.to_xml().unwrap();
    let source_before = master.to_xml().unwrap();
    assert!(local.materialize_notes_master_surface(&master).is_err());
    assert_eq!(local.to_xml().unwrap(), before);
    assert_eq!(master.to_xml().unwrap(), source_before);
}

#[test]
fn fx216_notes_tree_refuses_expanded_attribute_alias_conflicts() {
    let tree = |attrs: &str| {
        CT_ShapeTree::from_xml(format!(r#"<p:spTree xmlns:p="{P_NS}" xmlns:mc="{MC_NS}" xmlns:x="urn:meta" xmlns:y="urn:meta" mc:Ignorable="x y" {attrs}><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>"#).as_bytes()).unwrap()
    };
    let mut local = tree("y:keep=\"A\"");
    let before = local.to_xml().unwrap();
    let conflict = tree("x:keep=\"B\"");
    assert!(local.materialize_notes_master_surface(&conflict).is_err());
    assert_eq!(local.to_xml().unwrap(), before);
    let equal = tree("x:keep=\"A\"");
    local.materialize_notes_master_surface(&equal).unwrap();
    let xml = local.to_xml().unwrap();
    let raw = std::str::from_utf8(&xml).unwrap();
    assert!(raw.contains("y:keep=\"A\"") && !raw.contains("x:keep="));
    CT_ShapeTree::from_xml(&xml).unwrap();
}

#[test]
fn fx216_notes_tree_refuses_two_distinct_final_extension_owners() {
    let tree = |name: &str| {
        CT_ShapeTree::from_xml(format!(r#"<p:spTree xmlns:p="{P_NS}"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:extLst><p:ext uri="{name}"/></p:extLst></p:spTree>"#).as_bytes()).unwrap()
    };
    let mut local = tree("local");
    let master = tree("source");
    let before = local.to_xml().unwrap();
    assert!(local.materialize_notes_master_surface(&master).is_err());
    assert_eq!(local.to_xml().unwrap(), before);
}

#[test]
fn fx216_notes_tree_preserves_inherited_raw_boundary_before_its_actual_member() {
    let shape = |id, name| {
        CT_Shape::new_textbox(id, name, CT_Transform2D::default())
            .unwrap()
            .to_xml()
            .unwrap()
    };
    let first = String::from_utf8(shape(2, "First source")).unwrap();
    let second = String::from_utf8(shape(3, "Second source")).unwrap();
    let shell = |body: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:z="urn:source"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{body}</p:spTree>"#
        )
    };
    let mut local = CT_ShapeTree::from_xml(shell("").as_bytes()).unwrap();
    let master = CT_ShapeTree::from_xml(shell(&format!("<z:before keep='&#65;'/>{first}<z:between keep='&#66;'/>{second}<p:extLst><p:ext uri=\"source\"/></p:extLst>")).as_bytes()).unwrap();
    local.materialize_notes_master_surface(&master).unwrap();
    let xml = local.to_xml().unwrap();
    let raw = std::str::from_utf8(&xml).unwrap();
    assert!(raw.find("<z:before keep='&#65;'/>").unwrap() < raw.find("First source").unwrap());
    assert!(raw.find("First source").unwrap() < raw.find("<z:between keep='&#66;'/>").unwrap());
    assert!(raw.find("<z:between keep='&#66;'/>").unwrap() < raw.find("Second source").unwrap());
    assert_eq!(raw.matches("<p:extLst>").count(), 1);
    assert!(raw.find("Second source").unwrap() < raw.find("<p:extLst>").unwrap());
}

#[test]
fn fx216_notes_tree_keeps_equal_raw_children_at_distinct_boundaries() {
    let decoration = String::from_utf8(
        CT_Shape::new_textbox(2, "Source decoration", CT_Transform2D::default())
            .unwrap()
            .to_xml()
            .unwrap(),
    )
    .unwrap();
    let shell = |body: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:z="urn:meta"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{body}</p:spTree>"#
        )
    };
    let mut local = CT_ShapeTree::from_xml(shell("<z:marker/>").as_bytes()).unwrap();
    let master =
        CT_ShapeTree::from_xml(shell(&format!("<z:marker/>{decoration}")).as_bytes()).unwrap();
    local.materialize_notes_master_surface(&master).unwrap();
    let bytes = local.to_xml().unwrap();
    let xml = std::str::from_utf8(&bytes).unwrap();
    let positions: Vec<_> = xml.match_indices("<z:marker/>").map(|(at, _)| at).collect();
    assert_eq!(
        positions.len(),
        2,
        "both physical opaque boundaries must survive"
    );
    let member = xml.find("Source decoration").unwrap();
    assert!(positions[0] < member && member < positions[1]);
}

#[test]
fn fx216_notes_tree_preserves_source_raw_multiplicity_at_one_boundary() {
    let decoration = String::from_utf8(
        CT_Shape::new_textbox(2, "Source decoration", CT_Transform2D::default())
            .unwrap()
            .to_xml()
            .unwrap(),
    )
    .unwrap();
    let shell = |body: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:z="urn:meta"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{body}</p:spTree>"#
        )
    };
    for local_body in ["", "<z:marker/>"] {
        let mut local = CT_ShapeTree::from_xml(shell(local_body).as_bytes()).unwrap();
        let master = CT_ShapeTree::from_xml(
            shell(&format!("<z:marker/><z:marker/>{decoration}")).as_bytes(),
        )
        .unwrap();
        local.materialize_notes_master_surface(&master).unwrap();
        let bytes = local.to_xml().unwrap();
        let xml = std::str::from_utf8(&bytes).unwrap();
        let positions: Vec<_> = xml.match_indices("<z:marker/>").map(|(at, _)| at).collect();
        assert_eq!(
            positions.len(),
            if local_body.is_empty() { 2 } else { 3 },
            "source-owned multiplicity must survive"
        );
        let member = xml.find("Source decoration").unwrap();
        assert!(positions[0] < member && positions[1] < member);
        if !local_body.is_empty() {
            assert!(positions[2] > member);
        }
    }
}

#[test]
fn fx216_notes_tree_preserves_both_raw_sequences_at_one_boundary() {
    let shell = |body: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:z="urn:meta"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{body}</p:spTree>"#
        )
    };
    for (local_body, source_body, expected) in [
        ("<z:b/><z:a/>", "<z:a/><z:b/>", vec!["b", "a", "b"]),
        ("<z:a/>", "<z:b/><z:a/>", vec!["a", "b", "a"]),
        ("<z:a/><z:b/>", "<z:a/><z:b/>", vec!["a", "b"]),
        ("<z:a/>", "<z:a/><z:a/>", vec!["a", "a"]),
    ] {
        let mut local = CT_ShapeTree::from_xml(shell(local_body).as_bytes()).unwrap();
        let master = CT_ShapeTree::from_xml(shell(source_body).as_bytes()).unwrap();
        local.materialize_notes_master_surface(&master).unwrap();
        let bytes = local.to_xml().unwrap();
        let actual: Vec<_> = local
            .raw_children
            .at(0)
            .map(|raw| std::str::from_utf8(raw).unwrap())
            .collect();
        let expected: Vec<_> = expected.iter().map(|name| format!("<z:{name}/>")).collect();
        assert_eq!(actual, expected, "both original raw sequences must survive");
        let reopened = CT_ShapeTree::from_xml(&bytes).unwrap();
        assert_eq!(reopened.to_xml().unwrap(), bytes);
    }
}

#[test]
fn fx216_notes_tree_preserves_header_comment_order_and_multiplicity() {
    let shell = |header: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}">{header}<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>"#
        )
    };
    for (local_header, source_header, expected) in [
        (
            "<!--B--><!--A-->",
            "<!--A--><!--B-->",
            "<!--B--><!--A--><!--B-->",
        ),
        ("<!--A-->", "<!--B--><!--A-->", "<!--A--><!--B--><!--A-->"),
        ("", "<!--A--><!--A-->", "<!--A--><!--A-->"),
        ("<!--A--><!--B-->", "<!--A--><!--B-->", "<!--A--><!--B-->"),
    ] {
        let mut local = CT_ShapeTree::from_xml(shell(local_header).as_bytes()).unwrap();
        let master = CT_ShapeTree::from_xml(shell(source_header).as_bytes()).unwrap();
        local.materialize_notes_master_surface(&master).unwrap();
        assert_eq!(
            local.header_comments[0].concat(),
            expected.as_bytes(),
            "source header sequence must survive"
        );
        let bytes = local.to_xml().unwrap();
        assert_eq!(
            CT_ShapeTree::from_xml(&bytes).unwrap().to_xml().unwrap(),
            bytes
        );
    }
}

#[test]
fn fx216_notes_tree_keeps_actual_empty_and_field_content_without_master_hints() {
    let shape = |id: u32, kind: &str, text: &str| {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{kind}"/><p:cNvSpPr/><p:nvPr><p:ph type="{kind}" idx="{id}"/></p:nvPr></p:nvSpPr><p:spPr/>{text}</p:sp>"#
        )
    };
    let body =
        |text: &str| format!(r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p>{text}</a:p></p:txBody>"#);
    let shell = |text: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:z="urn:source"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{text}</p:spTree>"#
        )
    };
    let field = r#"<a:fld id="{A47B5AC0-96CA-4D3E-87D1-C4D13FDF4080}" type="slidenum"><a:rPr lang="en-US"/><a:t></a:t></a:fld>"#;
    let mut local = CT_ShapeTree::from_xml(
        shell(&format!(
            "{}{}{}",
            shape(2, "body", ""),
            shape(3, "sldNum", &body(field)),
            shape(4, "ftr", &body(""))
        ))
        .as_bytes(),
    )
    .unwrap();
    let before_field = match &local.children[1] {
        ShapeTreeChild::Shape(shape) => shape.text_body.clone(),
        _ => unreachable!(),
    };
    let master = CT_ShapeTree::from_xml(
        shell(&format!(
            "{}{}{}<z:before/>{}",
            shape(2, "body", &body("<a:r><a:t>MASTER SAMPLE</a:t></a:r>")),
            shape(3, "sldNum", &body("<a:r><a:t>MASTER NUMBER</a:t></a:r>")),
            shape(5, "hdr", &body("<a:r><a:t>MASTER HEADER</a:t></a:r>")),
            String::from_utf8(
                CT_Shape::new_textbox(8, "Decoration", CT_Transform2D::default())
                    .unwrap()
                    .to_xml()
                    .unwrap()
            )
            .unwrap()
        ))
        .as_bytes(),
    )
    .unwrap();
    local.materialize_notes_master_surface(&master).unwrap();
    assert_eq!(local.children.len(), 4);
    if let ShapeTreeChild::Shape(shape) = &local.children[0] {
        assert!(shape.text_body.is_none());
    } else {
        panic!("actual body owner");
    }
    if let ShapeTreeChild::Shape(shape) = &local.children[1] {
        assert_eq!(shape.text_body, before_field);
    } else {
        panic!("actual field owner");
    }
    if let ShapeTreeChild::Shape(shape) = &local.children[2] {
        assert!(
            shape.text_body.as_ref().unwrap().paragraphs()[0]
                .runs
                .is_empty()
        );
    } else {
        panic!("actual empty body owner");
    }
    let raw = String::from_utf8(local.to_xml().unwrap()).unwrap();
    assert!(!raw.contains("MASTER"));
    assert!(raw.find("<z:before/>").unwrap() < raw.find("Decoration").unwrap());
}

#[test]
fn fx216_notes_tree_refuses_references_to_uninstantiated_master_placeholders() {
    let shell = |body: &str| {
        format!(
            r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{body}</p:spTree>"#
        )
    };
    let placeholder = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Header hint"/><p:cNvSpPr/><p:nvPr><p:ph type="hdr" idx="0"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#;
    let connector = r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Source connector"/><p:cNvCxnSpPr><a:stCxn id="2" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#;
    let master =
        CT_ShapeTree::from_xml(shell(&format!("{placeholder}{connector}")).as_bytes()).unwrap();
    let master_before = master.to_xml().unwrap();
    let mut missing = CT_ShapeTree::from_xml(shell("").as_bytes()).unwrap();
    let before = missing.to_xml().unwrap();
    assert!(
        missing
            .materialize_notes_master_surface(&master)
            .unwrap_err()
            .to_string()
            .contains("uninstantiated source notes-master placeholder")
    );
    assert_eq!(missing.to_xml().unwrap(), before);
    let mut matched =
        CT_ShapeTree::from_xml(shell(&placeholder.replace("id=\"2\"", "id=\"42\"")).as_bytes())
            .unwrap();
    matched.materialize_notes_master_surface(&master).unwrap();
    let xml = matched.to_xml().unwrap();
    let ids = shape_id_occurrences(&xml, NamespaceBindings::default()).unwrap();
    assert!(
        ids.iter()
            .any(|value| !value.defines_shape && value.id == 42)
    );
    assert!(
        ids.iter()
            .any(|value| value.defines_shape && value.id == 42)
    );
    assert_eq!(master.to_xml().unwrap(), master_before);
}
#[cfg(test)]
#[test]
fn fx281_enum_hidden_dispatch_is_qualified_and_source_preserving() {
    let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="12"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>"#;
    let shape =
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"#;
    let picture = r#"<p:pic><p:nvPicPr><p:cNvPr id="3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>"#;
    let connector = r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#;
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shape}</p:grpSp>"#
    );
    let alternate = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="c">{frame}</mc:Choice><mc:Fallback>{shape}</mc:Fallback></mc:AlternateContent>"#
    );
    for (attribute, expected) in [
        ("hidden='1'", true),
        ("hidden='&#116;rue'", true),
        ("hidden='false'", false),
        ("xmlns:f='urn:foreign' f:hidden='1'", false),
    ] {
        let content = format!("{shape}{picture}{connector}{group}{frame}{alternate}")
            .replace("<p:cNvPr ", &format!("<p:cNvPr {attribute} "));
        let xml = format!(
            r#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:nvGrpSpPr/><p:grpSpPr/>{content}</p:spTree>"#
        );
        let tree = CT_ShapeTree::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(tree.children.len(), 6);
        let before = tree.to_xml().unwrap();
        assert!(
            tree.children.iter().all(|child| child.hidden() == expected),
            "{attribute}"
        );
        assert_eq!(before, tree.to_xml().unwrap());
    }
}
