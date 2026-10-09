use std::io::Write;
use std::ops::Range;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_drawing::fill::BlipFill;
use oxml_drawing::namespace::A_NS;
use oxml_drawing::order::OrderedRawChildren;
use oxml_drawing::shape_props::CT_ShapeProperties;
use oxml_drawing::xfrm::CT_Transform2D;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::namespace::{
    FIXED_SHAPE_TREE_PREFIXES, MC_NS, NamespaceBindings, P_NS, R_NS, non_visual_drawing_id,
    non_visual_drawing_name, root_attributes, self_contained_attributes,
    set_non_visual_drawing_name,
};
use crate::placeholder::{ApplicationProperties, CT_Placeholder, parse_application_properties};

pub type Result<T> = std::result::Result<T, OxmlError>;
type RawAttributes = Vec<(String, String)>;
const P14_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_MEDIA_EXTENSION_URI: &str = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/// The media family attached to a picture shape.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MediaKind {
    Audio,
    Video,
}

/// The package location carried by the Office 2010 media extension.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum MediaSource {
    Embedded { relationship_id: String },
    Linked { relationship_id: String },
}

impl MediaSource {
    pub fn relationship_id(&self) -> &str {
        match self {
            Self::Embedded { relationship_id } | Self::Linked { relationship_id } => {
                relationship_id
            }
        }
    }
}

/// The typed media projection of one retained `p:nvPr` subtree.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PictureMedia {
    pub kind: MediaKind,
    pub source: MediaSource,
    pub poster_relationship_id: Option<String>,
}

/// A typed `p:pic` shape-tree child.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_Picture {
    pub placeholder: Option<CT_Placeholder>,
    pub media: Option<PictureMedia>,
    pub blip_fill: Option<BlipFill>,
    pub shape_properties: CT_ShapeProperties,
    raw: Box<PictureRaw>,
}

#[derive(Clone, Debug, PartialEq)]
struct PictureRaw {
    raw_attributes: RawAttributes,
    non_visual_attributes: RawAttributes,
    non_visual_children: OrderedRawChildren,
    non_visual_drawing_properties: Vec<u8>,
    non_visual_name: Option<String>,
    non_visual_picture_properties: Vec<u8>,
    application_properties_attributes: RawAttributes,
    application_properties_raw_children: OrderedRawChildren,
    standard_media_relationship_id: Option<String>,
    office_media_relationship_ids: Vec<String>,
    media_trim_start_ms: Option<u64>,
    media_trim_end_ms: Option<u64>,
    raw_children: OrderedRawChildren,
}

struct ParsedNonVisualPicture {
    placeholder: Option<CT_Placeholder>,
    media: Option<PictureMedia>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
    non_visual_drawing_properties: Vec<u8>,
    non_visual_name: Option<String>,
    non_visual_picture_properties: Vec<u8>,
    application_properties_attributes: RawAttributes,
    application_properties_raw_children: OrderedRawChildren,
    standard_media_relationship_id: Option<String>,
    office_media_relationship_ids: Vec<String>,
    media_trim_start_ms: Option<u64>,
    media_trim_end_ms: Option<u64>,
}

struct ParsedPictureMedia {
    media: Option<PictureMedia>,
    standard_relationship_id: Option<String>,
    office_relationship_ids: Vec<String>,
    trim_start_ms: Option<u64>,
    trim_end_ms: Option<u64>,
}

struct ParsedP14Media {
    source: Option<MediaSource>,
    relationship_ids: Vec<String>,
    trim_start_ms: Option<u64>,
    trim_end_ms: Option<u64>,
}

impl CT_Picture {
    /// Creates a relationship-backed picture with canonical non-visual shells.
    ///
    /// The picture carries a rectangle preset geometry, as PowerPoint draws
    /// nothing for a picture whose shape properties have no geometry.
    pub fn new(
        id: u32,
        name: &str,
        relationship_id: &str,
        transform: CT_Transform2D,
    ) -> Result<Self> {
        let name = quick_xml::escape::escape(name);
        let relationship_id = quick_xml::escape::escape(relationship_id);
        let transform_xml = transform.to_xml().map_err(drawing_error)?;
        let transform_xml = std::str::from_utf8(&transform_xml)?;
        Self::from_xml(
            format!(
                r#"<p:pic xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"><p:nvPicPr><p:cNvPr id="{id}" name="{name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{transform_xml}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
            )
            .as_bytes(),
        )
    }

    /// Creates the picture that fills a picture placeholder, as python-pptx
    /// `insert_picture` writes it.
    ///
    /// The picture keeps the placeholder's `p:ph`, so it inherits its
    /// position and size from the layout through an empty `p:spPr`. `crop`
    /// holds the left, top, right, and bottom `a:srcRect` insets in
    /// thousandths of a percent, and four zeros write none.
    pub fn new_placeholder(
        id: u32,
        name: &str,
        relationship_id: &str,
        placeholder: &CT_Placeholder,
        crop: [i32; 4],
    ) -> Result<Self> {
        let name = quick_xml::escape::escape(name);
        let relationship_id = quick_xml::escape::escape(relationship_id);
        let mut placeholder_xml = Vec::new();
        placeholder.write_xml(&mut Writer::new(&mut placeholder_xml), false)?;
        let placeholder_xml = std::str::from_utf8(&placeholder_xml)?;
        let source_rect = if crop == [0; 4] {
            String::new()
        } else {
            let [left, top, right, bottom] = crop;
            let edge = |name: &str, value: i32| {
                if value == 0 {
                    String::new()
                } else {
                    format!(r#" {name}="{value}""#)
                }
            };
            format!(
                "<a:srcRect{}{}{}{}/>",
                edge("l", left),
                edge("t", top),
                edge("r", right),
                edge("b", bottom)
            )
        };
        Self::from_xml(
            format!(
                r#"<p:pic xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"><p:nvPicPr><p:cNvPr id="{id}" name="{name}"/><p:cNvPicPr><a:picLocks noGrp="1" noChangeAspect="1"/></p:cNvPicPr><p:nvPr>{placeholder_xml}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="{relationship_id}"/>{source_rect}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>"#
            )
            .as_bytes(),
        )
    }

    /// Creates a media picture with standard and Office 2010 relationships.
    ///
    /// Its poster frame carries the same rectangle preset geometry as a
    /// picture made by [`CT_Picture::new`].
    #[allow(clippy::too_many_arguments)]
    pub fn new_media(
        id: u32,
        name: &str,
        kind: MediaKind,
        source: MediaSource,
        microsoft_relationship_id: &str,
        poster_relationship_id: &str,
        trim_start_ms: Option<u64>,
        trim_end_ms: Option<u64>,
        transform: CT_Transform2D,
    ) -> Result<Self> {
        let name = quick_xml::escape::escape(name);
        let source_relationship_id = quick_xml::escape::escape(source.relationship_id());
        let microsoft_relationship_id = quick_xml::escape::escape(microsoft_relationship_id);
        let poster_relationship_id = quick_xml::escape::escape(poster_relationship_id);
        let transform_xml = transform.to_xml().map_err(drawing_error)?;
        let transform_xml = std::str::from_utf8(&transform_xml)?;
        let media_tag = match kind {
            MediaKind::Audio => "audioFile",
            MediaKind::Video => "videoFile",
        };
        let extension_attribute = match source {
            MediaSource::Embedded { .. } => "embed",
            MediaSource::Linked { .. } => "link",
        };
        let trim = if trim_start_ms.is_some() || trim_end_ms.is_some() {
            let start = trim_start_ms
                .map(|value| format!(r#" st="{value}""#))
                .unwrap_or_default();
            let end = trim_end_ms
                .map(|value| format!(r#" end="{value}""#))
                .unwrap_or_default();
            format!(r#"><p14:trim{start}{end}/></p14:media>"#)
        } else {
            "/>".to_owned()
        };
        Self::from_xml(
            format!(
                r#"<p:pic xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p14="{P14_NS}"><p:nvPicPr><p:cNvPr id="{id}" name="{name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:{media_tag} r:link="{source_relationship_id}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}"><p14:media r:{extension_attribute}="{microsoft_relationship_id}"{trim}</p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="{poster_relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{transform_xml}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
            )
            .as_bytes(),
        )
    }

    pub(crate) fn non_visual_id(&self) -> Option<u32> {
        non_visual_drawing_id(&self.raw.non_visual_drawing_properties)
    }

    pub(crate) fn non_visual_name(&self) -> Option<&str> {
        self.raw.non_visual_name.as_deref()
    }

    /// Returns the standard audio or video relationship paired with `p14:media`.
    pub fn standard_media_relationship_id(&self) -> Option<&str> {
        self.raw.standard_media_relationship_id.as_deref()
    }

    /// Returns every relationship carried by the Office 2010 media extension.
    pub fn office_media_relationship_ids(&self) -> &[String] {
        &self.raw.office_media_relationship_ids
    }

    /// Returns the rounded millisecond trim bounds from `p14:trim`.
    pub fn media_trim_bounds(&self) -> (Option<u64>, Option<u64>) {
        (self.raw.media_trim_start_ms, self.raw.media_trim_end_ms)
    }

    /// Rewrites only this picture's standard and Office media relationship ids.
    pub fn replace_media_relationship_ids(
        &mut self,
        standard_relationship_id: &str,
        source: MediaSource,
    ) -> Result<()> {
        self.media
            .as_ref()
            .ok_or_else(|| OxmlError::MissingElement("picture media".to_owned()))?;
        self.standard_media_relationship_id()
            .ok_or_else(|| OxmlError::MissingElement("standard media relationship".to_owned()))?;
        let rewritten =
            replace_standard_media_relationship(&self.to_xml()?, standard_relationship_id)?;
        let rewritten = if self.office_media_relationship_ids().is_empty() {
            rewritten
        } else {
            replace_p14_media_source(&rewritten, &source)?
        };
        *self = Self::from_xml(&rewritten)?;
        Ok(())
    }

    pub(crate) fn drawing_properties_mut(&mut self) -> &mut Vec<u8> {
        &mut self.raw.non_visual_drawing_properties
    }

    pub(crate) fn drawing_properties(&self) -> &[u8] {
        &self.raw.non_visual_drawing_properties
    }

    /// Changes the producer-facing non-visual picture name.
    pub fn set_name(&mut self, name: &str) -> Result<()> {
        set_non_visual_drawing_name(&mut self.raw.non_visual_drawing_properties, name)?;
        self.raw.non_visual_name = Some(name.to_owned());
        Ok(())
    }

    /// Parses a complete `p:pic` with any prefix bound to PresentationML.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_fragment(xml, &[])
    }

    pub(crate) fn from_fragment(xml: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces =
                        NamespaceBindings::from_entries(inherited).with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"pic"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
                    return Self::from_reader(&mut reader, &start, &namespaces, inherited);
                }
                Event::Empty(start) => {
                    return Err(OxmlError::MissingElement(format!(
                        "{} requires p:nvPicPr, a blip-fill choice, and p:spPr",
                        String::from_utf8_lossy(start.name().as_ref())
                    )));
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:pic".to_owned())),
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
        let mut blip_fill = None;
        let mut shape_properties = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut state = 0usize;
        let mut has_alternate_blip = false;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_element(reader, &child)?;
                    capture_picture_child(
                        &name,
                        uri,
                        raw,
                        &child_namespaces,
                        &mut non_visual,
                        &mut blip_fill,
                        &mut shape_properties,
                        &mut raw_children,
                        &mut state,
                        &mut has_alternate_blip,
                    )?;
                }
                Event::Empty(child) => {
                    let child_namespaces = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let uri = child_namespaces.element_uri(child.name().as_ref());
                    let raw = capture_empty_element(&child)?;
                    capture_picture_child(
                        &name,
                        uri,
                        raw,
                        &child_namespaces,
                        &mut non_visual,
                        &mut blip_fill,
                        &mut shape_properties,
                        &mut raw_children,
                        &mut state,
                        &mut has_alternate_blip,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"pic" => break,
                Event::Eof => return Err(OxmlError::MissingElement("closing p:pic".to_owned())),
                _ => {}
            }
            buffer.clear();
        }

        if blip_fill.is_none() && !has_alternate_blip {
            return Err(OxmlError::MissingElement(
                "p:pic blip-fill choice".to_owned(),
            ));
        }
        let non_visual = required(non_visual, "p:nvPicPr")?;
        let mut media = non_visual.media;
        if let Some(media) = &mut media {
            media.poster_relationship_id = blip_fill
                .as_ref()
                .and_then(|fill| fill.blip.as_ref())
                .and_then(|blip| blip.embed.as_ref().or(blip.link.as_ref()))
                .cloned();
        }
        Ok(Self {
            placeholder: non_visual.placeholder,
            media,
            blip_fill,
            shape_properties: required(shape_properties, "p:spPr")?,
            raw: Box::new(PictureRaw {
                raw_attributes: self_contained_attributes(
                    start,
                    FIXED_SHAPE_TREE_PREFIXES,
                    inherited,
                )?,
                non_visual_attributes: non_visual.raw_attributes,
                non_visual_children: non_visual.raw_children,
                non_visual_drawing_properties: non_visual.non_visual_drawing_properties,
                non_visual_name: non_visual.non_visual_name,
                non_visual_picture_properties: non_visual.non_visual_picture_properties,
                application_properties_attributes: non_visual.application_properties_attributes,
                application_properties_raw_children: non_visual.application_properties_raw_children,
                standard_media_relationship_id: non_visual.standard_media_relationship_id,
                office_media_relationship_ids: non_visual.office_media_relationship_ids,
                media_trim_start_ms: non_visual.media_trim_start_ms,
                media_trim_end_ms: non_visual.media_trim_end_ms,
                raw_children,
            }),
        })
    }

    /// Serialises a self-contained picture with fixed modelled prefixes.
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
        let mut start = BytesStart::new("p:pic");
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
        if let Some(blip_fill) = &self.blip_fill {
            blip_fill
                .write_xml_as(writer, "p:blipFill")
                .map_err(drawing_error)?;
        }
        emit_raw(writer, self.raw.raw_children.at(2))?;
        self.shape_properties
            .write_xml_as(writer, "p:spPr")
            .map_err(drawing_error)?;
        for boundary in 3..=5 {
            emit_raw(writer, self.raw.raw_children.at(boundary))?;
        }
        writer.write_event(Event::End(BytesEnd::new("p:pic")))?;
        Ok(())
    }

    fn write_non_visual_properties<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:nvPicPr");
        push_attributes(&mut start, &self.raw.non_visual_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw.non_visual_children.at(0))?;
        writer
            .get_mut()
            .write_all(&self.raw.non_visual_drawing_properties)?;
        emit_raw(writer, self.raw.non_visual_children.at(1))?;
        writer
            .get_mut()
            .write_all(&self.raw.non_visual_picture_properties)?;
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
        writer.write_event(Event::End(BytesEnd::new("p:nvPicPr")))?;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn capture_picture_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    non_visual: &mut Option<ParsedNonVisualPicture>,
    blip_fill: &mut Option<BlipFill>,
    shape_properties: &mut Option<CT_ShapeProperties>,
    raw_children: &mut OrderedRawChildren,
    state: &mut usize,
    has_alternate_blip: &mut bool,
) -> Result<()> {
    if uri == Some(P_NS) && matches!(name, b"nvPicPr" | b"blipFill" | b"spPr") {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
    }
    match (uri, name) {
        (Some(P_NS), b"nvPicPr") if *state == 0 && non_visual.is_none() => {
            *non_visual = Some(parse_non_visual_picture(&raw, namespaces)?);
            *state = 1;
        }
        (Some(P_NS), b"blipFill") if *state == 1 && blip_fill.is_none() && !*has_alternate_blip => {
            validate_blip_fill(&raw, namespaces)?;
            *blip_fill = Some(BlipFill::from_xml(&raw).map_err(drawing_error)?);
            *state = 2;
        }
        (Some(MC_NS), b"AlternateContent")
            if *state == 1 && blip_fill.is_none() && !*has_alternate_blip =>
        {
            namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
            if !alternate_contains_blip_fill(&raw, namespaces)? {
                return Err(OxmlError::InvalidValue(
                    "p:pic alternate content does not contain a p:blipFill choice".to_owned(),
                ));
            }
            raw_children.push(1, raw);
            *has_alternate_blip = true;
        }
        (Some(P_NS), b"spPr") if matches!(*state, 1 | 2) && shape_properties.is_none() => {
            if blip_fill.is_none() && !*has_alternate_blip {
                return Err(OxmlError::InvalidValue(
                    "p:spPr must follow the p:pic blip-fill choice".to_owned(),
                ));
            }
            *shape_properties = Some(CT_ShapeProperties::from_xml(&raw).map_err(drawing_error)?);
            *state = 3;
        }
        (Some(P_NS), b"style") if *state == 3 => {
            raw_children.push(3, raw);
            *state = 4;
        }
        (Some(P_NS), b"extLst") if matches!(*state, 3 | 4) => {
            raw_children.push(4, raw);
            *state = 5;
        }
        (Some(P_NS), b"nvPicPr" | b"blipFill" | b"spPr" | b"style" | b"extLst") => {
            return Err(OxmlError::InvalidValue(
                "p:pic children are outside schema order or duplicated".to_owned(),
            ));
        }
        (Some(MC_NS), b"AlternateContent") => {
            return Err(OxmlError::InvalidValue(
                "p:pic blip-fill choice is duplicated or outside schema order".to_owned(),
            ));
        }
        _ => raw_children.push(*state, raw),
    }
    Ok(())
}

fn alternate_contains_blip_fill(xml: &[u8], inherited: &NamespaceBindings) -> Result<bool> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                loop {
                    buffer.clear();
                    match reader.read_event_into(&mut buffer)? {
                        Event::Start(branch) => {
                            let branch_namespaces = namespaces.with_start(&branch)?;
                            let is_branch = branch_namespaces.element_uri(branch.name().as_ref())
                                == Some(MC_NS)
                                && matches!(
                                    local_name(branch.name().as_ref()),
                                    b"Choice" | b"Fallback"
                                );
                            if is_branch
                                && branch_contains_blip_fill(
                                    &mut reader,
                                    &branch,
                                    &branch_namespaces,
                                )?
                            {
                                return Ok(true);
                            }
                            if !is_branch {
                                capture_element(&mut reader, &branch)?;
                            }
                        }
                        Event::Empty(_) => {}
                        Event::End(end)
                            if local_name(end.name().as_ref()) == b"AlternateContent" =>
                        {
                            return Ok(false);
                        }
                        Event::Eof => {
                            return Err(OxmlError::MissingElement(
                                "closing mc:AlternateContent".to_owned(),
                            ));
                        }
                        _ => {}
                    }
                }
            }
            Event::Empty(_) => return Ok(false),
            Event::Eof => {
                return Err(OxmlError::MissingElement("mc:AlternateContent".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn branch_contains_blip_fill(
    reader: &mut Reader<&[u8]>,
    branch: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<bool> {
    let end_name = local_name(branch.name().as_ref()).to_vec();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_namespaces = namespaces.with_start(&child)?;
                if child_namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"blipFill"
                {
                    return Ok(true);
                }
                capture_element(reader, &child)?;
            }
            Event::Empty(child) => {
                let child_namespaces = namespaces.with_start(&child)?;
                if child_namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"blipFill"
                {
                    return Ok(true);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == end_name => return Ok(false),
            Event::Eof => {
                return Err(OxmlError::MissingElement(format!(
                    "closing mc:{}",
                    String::from_utf8_lossy(&end_name)
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_non_visual_picture(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<ParsedNonVisualPicture> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                let raw_attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?;
                let mut drawing = None;
                let mut drawing_name = None;
                let mut picture = None;
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
                            capture_non_visual_child(
                                &name,
                                uri,
                                raw,
                                candidate_name,
                                &child_namespaces,
                                &mut drawing,
                                &mut drawing_name,
                                &mut picture,
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
                            capture_non_visual_child(
                                &name,
                                uri,
                                raw,
                                candidate_name,
                                &child_namespaces,
                                &mut drawing,
                                &mut drawing_name,
                                &mut picture,
                                &mut application,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::End(end) if local_name(end.name().as_ref()) == b"nvPicPr" => break,
                        Event::Eof => {
                            return Err(OxmlError::MissingElement("closing p:nvPicPr".to_owned()));
                        }
                        _ => {}
                    }
                }
                let application = required(application, "p:nvPr")?;
                let drawing = required(drawing, "p:cNvPr")?;
                let parsed_media = parse_picture_media(xml, inherited)?;
                return Ok(ParsedNonVisualPicture {
                    placeholder: application.placeholder,
                    media: parsed_media.media,
                    raw_attributes,
                    raw_children,
                    non_visual_name: drawing_name,
                    non_visual_drawing_properties: drawing,
                    non_visual_picture_properties: required(picture, "p:cNvPicPr")?,
                    application_properties_attributes: application.raw_attributes,
                    application_properties_raw_children: application.raw_children,
                    standard_media_relationship_id: parsed_media.standard_relationship_id,
                    office_media_relationship_ids: parsed_media.office_relationship_ids,
                    media_trim_start_ms: parsed_media.trim_start_ms,
                    media_trim_end_ms: parsed_media.trim_end_ms,
                });
            }
            Event::Empty(_) => {
                return Err(OxmlError::MissingElement(
                    "p:nvPicPr requires p:cNvPr, p:cNvPicPr, and p:nvPr".to_owned(),
                ));
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:nvPicPr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_picture_media(
    non_visual_xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<ParsedPictureMedia> {
    let mut reader = Reader::from_reader(non_visual_xml);
    let mut buffer = Vec::new();
    let mut nv_picture_namespaces = None;
    let mut application_namespaces = None;
    let mut kind = None;
    let mut standard_relationship_id = None;
    let mut extension = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if nv_picture_namespaces.is_none() => {
                nv_picture_namespaces = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) if application_namespaces.is_none() => {
                let scope = nv_picture_namespaces
                    .as_ref()
                    .expect("non-visual picture root")
                    .with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"nvPr"
                {
                    application_namespaces = Some(scope);
                } else {
                    capture_element(&mut reader, &child)?;
                }
            }
            Event::Start(child) => {
                let scope = application_namespaces
                    .as_ref()
                    .expect("application properties root")
                    .with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(A_NS)
                    && matches!(
                        local_name(child.name().as_ref()),
                        b"audioFile" | b"videoFile"
                    )
                {
                    if kind.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "picture has more than one audio or video attachment".to_owned(),
                        ));
                    }
                    kind = Some(if local_name(child.name().as_ref()) == b"audioFile" {
                        MediaKind::Audio
                    } else {
                        MediaKind::Video
                    });
                    standard_relationship_id = relationship_attribute(&child, &scope, b"link")?;
                }
                if extension.is_none()
                    && scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"extLst"
                {
                    extension = find_p14_media_extension(
                        &raw,
                        application_namespaces
                            .as_ref()
                            .expect("application properties root"),
                    )?;
                }
            }
            Event::Empty(child) if application_namespaces.is_none() => {
                let scope = nv_picture_namespaces
                    .as_ref()
                    .expect("non-visual picture root")
                    .with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"nvPr"
                {
                    return Ok(ParsedPictureMedia {
                        media: None,
                        standard_relationship_id: None,
                        office_relationship_ids: Vec::new(),
                        trim_start_ms: None,
                        trim_end_ms: None,
                    });
                }
            }
            Event::Empty(child) => {
                let scope = application_namespaces
                    .as_ref()
                    .expect("application properties root")
                    .with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(A_NS)
                    && matches!(
                        local_name(child.name().as_ref()),
                        b"audioFile" | b"videoFile"
                    )
                {
                    if kind.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "picture has more than one audio or video attachment".to_owned(),
                        ));
                    }
                    kind = Some(if local_name(child.name().as_ref()) == b"audioFile" {
                        MediaKind::Audio
                    } else {
                        MediaKind::Video
                    });
                    standard_relationship_id = relationship_attribute(&child, &scope, b"link")?;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"nvPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let Some(kind) = kind else {
        return Ok(ParsedPictureMedia {
            media: None,
            standard_relationship_id: None,
            office_relationship_ids: Vec::new(),
            trim_start_ms: None,
            trim_end_ms: None,
        });
    };
    let source = extension
        .as_ref()
        .and_then(|media| media.source.clone())
        .or_else(|| {
            standard_relationship_id
                .clone()
                .map(|relationship_id| MediaSource::Linked { relationship_id })
        });
    Ok(ParsedPictureMedia {
        media: source.map(|source| PictureMedia {
            kind,
            source,
            poster_relationship_id: None,
        }),
        standard_relationship_id,
        office_relationship_ids: extension
            .as_ref()
            .map(|media| media.relationship_ids.clone())
            .unwrap_or_default(),
        trim_start_ms: extension.as_ref().and_then(|media| media.trim_start_ms),
        trim_end_ms: extension.as_ref().and_then(|media| media.trim_end_ms),
    })
}

fn find_p14_media_extension(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<ParsedP14Media>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let scope = root.as_ref().expect("extension root").with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"ext"
                    && unqualified_attribute(&child, b"uri")?.as_deref()
                        == Some(P14_MEDIA_EXTENSION_URI)
                {
                    return find_direct_p14_media(&raw, root.as_ref().expect("extension root"));
                }
            }
            Event::Empty(_) => {}
            Event::End(_) | Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn find_direct_p14_media(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<ParsedP14Media>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let scope = root
                    .as_ref()
                    .expect("media extension root")
                    .with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P14_NS)
                    && local_name(child.name().as_ref()) == b"media"
                {
                    let (source, relationship_ids) = p14_media_sources(&child, &scope)?;
                    let (trim_start_ms, trim_end_ms) =
                        parse_p14_trim(&raw, root.as_ref().expect("media extension root"))?;
                    return Ok(Some(ParsedP14Media {
                        source,
                        relationship_ids,
                        trim_start_ms,
                        trim_end_ms,
                    }));
                }
            }
            Event::Empty(child) => {
                let scope = root
                    .as_ref()
                    .expect("media extension root")
                    .with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P14_NS)
                    && local_name(child.name().as_ref()) == b"media"
                {
                    let (source, relationship_ids) = p14_media_sources(&child, &scope)?;
                    return Ok(Some(ParsedP14Media {
                        source,
                        relationship_ids,
                        trim_start_ms: None,
                        trim_end_ms: None,
                    }));
                }
            }
            Event::End(_) | Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn unqualified_attribute(start: &BytesStart<'_>, expected: &[u8]) -> Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() == expected {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn parse_p14_trim(
    media_xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<(Option<u64>, Option<u64>)> {
    let mut reader = Reader::from_reader(media_xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut trim = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("p14:media root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P14_NS)
                    && local_name(child.name().as_ref()) == b"trim"
                {
                    if trim.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "p14:media has more than one p14:trim child".to_owned(),
                        ));
                    }
                    trim = Some(parse_trim_attributes(&child)?);
                }
                capture_element(&mut reader, &child)?;
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("p14:media root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P14_NS)
                    && local_name(child.name().as_ref()) == b"trim"
                {
                    if trim.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "p14:media has more than one p14:trim child".to_owned(),
                        ));
                    }
                    trim = Some(parse_trim_attributes(&child)?);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"media" => {
                return Ok(trim.unwrap_or((None, None)));
            }
            Event::Eof => return Ok(trim.unwrap_or((None, None))),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_trim_attributes(start: &BytesStart<'_>) -> Result<(Option<u64>, Option<u64>)> {
    let mut raw_start = None;
    let mut raw_end = None;
    for attribute in start.attributes() {
        let attribute = attribute?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        match attribute.key.as_ref() {
            b"st" => raw_start = Some(parse_trim_value(&value, "st")?),
            b"end" => raw_end = Some(parse_trim_value(&value, "end")?),
            _ => {}
        }
    }
    Ok((
        raw_start.map(round_trim_value),
        raw_end.map(round_trim_value),
    ))
}

fn parse_trim_value(value: &str, name: &str) -> Result<f64> {
    let value = value
        .parse::<f64>()
        .map_err(|_| OxmlError::InvalidValue(format!("invalid p14:trim {name} value {value}")))?;
    if !value.is_finite() || value < 0.0 || value > u64::MAX as f64 {
        return Err(OxmlError::InvalidValue(format!(
            "p14:trim {name} value is out of range"
        )));
    }
    Ok(value)
}

fn round_trim_value(value: f64) -> u64 {
    value.round() as u64
}

fn p14_media_sources(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<(Option<MediaSource>, Vec<String>)> {
    let embedded = relationship_attribute(start, namespaces, b"embed")?;
    let linked = relationship_attribute(start, namespaces, b"link")?;
    let mut relationship_ids = Vec::with_capacity(2);
    if let Some(relationship_id) = &embedded {
        relationship_ids.push(relationship_id.clone());
    }
    if let Some(relationship_id) = &linked
        && !relationship_ids.contains(relationship_id)
    {
        relationship_ids.push(relationship_id.clone());
    }
    let source = linked
        .map(|relationship_id| MediaSource::Linked { relationship_id })
        .or_else(|| embedded.map(|relationship_id| MediaSource::Embedded { relationship_id }));
    Ok((source, relationship_ids))
}

fn replace_standard_media_relationship(xml: &[u8], relationship_id: &str) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut scopes = vec![NamespaceBindings::default()];
    let mut elements: Vec<(Option<String>, Vec<u8>)> = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let scope = scopes
                    .last()
                    .expect("picture namespace scope")
                    .with_start(&start)?;
                let uri = scope.element_uri(start.name().as_ref()).map(str::to_owned);
                let local = local_name(start.name().as_ref()).to_vec();
                if is_direct_standard_media_element(&elements, uri.as_deref(), &local) {
                    return replace_relationship_attribute_value(
                        xml,
                        event_start..reader.buffer_position() as usize,
                        &start,
                        &scope,
                        b"link",
                        relationship_id,
                    );
                }
                scopes.push(scope);
                elements.push((uri, local));
            }
            Event::Empty(start) => {
                let scope = scopes
                    .last()
                    .expect("picture namespace scope")
                    .with_start(&start)?;
                let uri = scope.element_uri(start.name().as_ref());
                let start_name = start.name();
                let local = local_name(start_name.as_ref());
                if is_direct_standard_media_element(&elements, uri, local) {
                    return replace_relationship_attribute_value(
                        xml,
                        event_start..reader.buffer_position() as usize,
                        &start,
                        &scope,
                        b"link",
                        relationship_id,
                    );
                }
            }
            Event::End(_) => {
                scopes.pop();
                elements.pop();
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "standard audio or video relationship".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn is_direct_standard_media_element(
    elements: &[(Option<String>, Vec<u8>)],
    uri: Option<&str>,
    local: &[u8],
) -> bool {
    uri == Some(A_NS)
        && matches!(local, b"audioFile" | b"videoFile")
        && elements.len() == 3
        && elements[0].0.as_deref() == Some(P_NS)
        && elements[0].1.as_slice() == b"pic"
        && elements[1].0.as_deref() == Some(P_NS)
        && elements[1].1.as_slice() == b"nvPicPr"
        && elements[2].0.as_deref() == Some(P_NS)
        && elements[2].1.as_slice() == b"nvPr"
}

fn replace_relationship_attribute_value(
    xml: &[u8],
    tag_range: Range<usize>,
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
    expected_local: &[u8],
    value: &str,
) -> Result<Vec<u8>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        if namespaces.attribute_uri(name) != Some(R_NS) || local_name(name) != expected_local {
            continue;
        }
        let attribute_value = attribute.value.as_ref();
        let relative_start = (attribute_value.as_ptr() as usize)
            .checked_sub(start.as_ptr() as usize)
            .filter(|offset| {
                offset
                    .checked_add(attribute_value.len())
                    .is_some_and(|end| end <= start.len())
            })
            .ok_or_else(|| {
                OxmlError::InvalidValue("media relationship value is outside its tag".to_owned())
            })?;
        let value_start = tag_range
            .start
            .checked_add(1)
            .and_then(|offset| offset.checked_add(relative_start))
            .ok_or_else(|| {
                OxmlError::InvalidValue("media relationship range overflowed".to_owned())
            })?;
        let value_end = value_start
            .checked_add(attribute_value.len())
            .ok_or_else(|| {
                OxmlError::InvalidValue("media relationship range overflowed".to_owned())
            })?;
        if value_end > tag_range.end || xml.get(value_start..value_end) != Some(attribute_value) {
            return Err(OxmlError::InvalidValue(
                "media relationship range did not match its tag".to_owned(),
            ));
        }
        let escaped = quick_xml::escape::escape(value);
        let mut rewritten = Vec::with_capacity(
            xml.len()
                .saturating_add(escaped.len().saturating_sub(attribute_value.len())),
        );
        rewritten.extend_from_slice(&xml[..value_start]);
        rewritten.extend_from_slice(escaped.as_bytes());
        rewritten.extend_from_slice(&xml[value_end..]);
        return Ok(rewritten);
    }
    Err(OxmlError::MissingElement(
        "standard media relationship attribute".to_owned(),
    ))
}

fn replace_p14_media_source(xml: &[u8], source: &MediaSource) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut scopes = vec![NamespaceBindings::default()];
    let mut elements: Vec<(Option<String>, Vec<u8>, bool)> = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let scope = scopes
                    .last()
                    .expect("picture namespace scope")
                    .with_start(&start)?;
                let uri = scope.element_uri(start.name().as_ref()).map(str::to_owned);
                let local = local_name(start.name().as_ref()).to_vec();
                if elements
                    .last()
                    .is_some_and(|(_, _, media_extension)| *media_extension)
                    && uri.as_deref() == Some(P14_NS)
                    && local.as_slice() == b"media"
                {
                    return replace_media_source_attributes(
                        xml,
                        event_start..reader.buffer_position() as usize,
                        &start,
                        &scope,
                        source,
                    );
                }
                let is_media_extension = uri.as_deref() == Some(P_NS)
                    && local.as_slice() == b"ext"
                    && elements.len() == 4
                    && elements[0].0.as_deref() == Some(P_NS)
                    && elements[0].1.as_slice() == b"pic"
                    && elements[1].0.as_deref() == Some(P_NS)
                    && elements[1].1.as_slice() == b"nvPicPr"
                    && elements[2].0.as_deref() == Some(P_NS)
                    && elements[2].1.as_slice() == b"nvPr"
                    && elements[3].0.as_deref() == Some(P_NS)
                    && elements[3].1.as_slice() == b"extLst"
                    && unqualified_attribute(&start, b"uri")?.as_deref()
                        == Some(P14_MEDIA_EXTENSION_URI);
                scopes.push(scope);
                elements.push((uri, local, is_media_extension));
            }
            Event::Empty(start) => {
                let scope = scopes
                    .last()
                    .expect("picture namespace scope")
                    .with_start(&start)?;
                if elements
                    .last()
                    .is_some_and(|(_, _, media_extension)| *media_extension)
                    && scope.element_uri(start.name().as_ref()) == Some(P14_NS)
                    && local_name(start.name().as_ref()) == b"media"
                {
                    return replace_media_source_attributes(
                        xml,
                        event_start..reader.buffer_position() as usize,
                        &start,
                        &scope,
                        source,
                    );
                }
            }
            Event::End(_) => {
                scopes.pop();
                elements.pop();
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("p14:media".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn replace_media_source_attributes(
    xml: &[u8],
    tag_range: Range<usize>,
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
    source: &MediaSource,
) -> Result<Vec<u8>> {
    let mut removals = Vec::new();
    let mut relationship_prefix = None;
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        if namespaces.attribute_uri(name) != Some(R_NS)
            || !matches!(local_name(name), b"embed" | b"link")
        {
            continue;
        }
        if relationship_prefix.is_none() {
            relationship_prefix = name
                .iter()
                .position(|byte| *byte == b':')
                .map(|position| name[..position].to_vec());
        }
        let element_address = start.as_ptr() as usize;
        let name_start = (attribute.key.as_ref().as_ptr() as usize)
            .checked_sub(element_address)
            .and_then(|offset| tag_range.start.checked_add(1 + offset))
            .ok_or_else(|| {
                OxmlError::InvalidValue("media attribute range overflowed".to_owned())
            })?;
        let value = attribute.value.as_ref();
        let value_start = (value.as_ptr() as usize)
            .checked_sub(element_address)
            .and_then(|offset| tag_range.start.checked_add(1 + offset))
            .ok_or_else(|| {
                OxmlError::InvalidValue("media attribute range overflowed".to_owned())
            })?;
        let mut removal_start = name_start;
        while removal_start > tag_range.start + 1 && xml[removal_start - 1].is_ascii_whitespace() {
            removal_start -= 1;
        }
        let value_end = value_start.checked_add(value.len()).ok_or_else(|| {
            OxmlError::InvalidValue("media attribute range overflowed".to_owned())
        })?;
        let quote = *xml
            .get(value_start.wrapping_sub(1))
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| OxmlError::InvalidValue("unquoted media attribute".to_owned()))?;
        if xml.get(value_end) != Some(&quote) {
            return Err(OxmlError::InvalidValue(
                "unterminated media attribute".to_owned(),
            ));
        }
        removals.push(removal_start..value_end + 1);
    }
    removals.sort_unstable_by_key(|range| range.start);
    let relationship_prefix = relationship_prefix.ok_or_else(|| {
        OxmlError::MissingElement("p14:media relationship source attribute".to_owned())
    })?;
    let relationship_prefix = std::str::from_utf8(&relationship_prefix)?;

    let tag = xml
        .get(tag_range.clone())
        .ok_or_else(|| OxmlError::InvalidValue("media start tag is out of range".to_owned()))?;
    let mut rewritten_tag = Vec::with_capacity(tag.len() + source.relationship_id().len() + 20);
    let mut copied_through = tag_range.start;
    for removal in removals {
        if removal.start < copied_through || removal.end > tag_range.end {
            return Err(OxmlError::InvalidValue(
                "media attribute ranges overlap".to_owned(),
            ));
        }
        rewritten_tag.extend_from_slice(&xml[copied_through..removal.start]);
        copied_through = removal.end;
    }
    rewritten_tag.extend_from_slice(&xml[copied_through..tag_range.end]);
    let terminator = rewritten_tag
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| OxmlError::InvalidValue("media start tag has no terminator".to_owned()))?;
    let insertion = if rewritten_tag.get(terminator.wrapping_sub(1)) == Some(&b'/') {
        terminator - 1
    } else {
        terminator
    };
    let attribute_name = match source {
        MediaSource::Embedded { .. } => "embed",
        MediaSource::Linked { .. } => "link",
    };
    let attribute = format!(
        r#" {relationship_prefix}:{attribute_name}="{}""#,
        quick_xml::escape::escape(source.relationship_id())
    );
    rewritten_tag.splice(insertion..insertion, attribute.bytes());

    let mut rewritten = Vec::with_capacity(xml.len() + rewritten_tag.len() - tag.len());
    rewritten.extend_from_slice(&xml[..tag_range.start]);
    rewritten.extend_from_slice(&rewritten_tag);
    rewritten.extend_from_slice(&xml[tag_range.end..]);
    Ok(rewritten)
}

fn relationship_attribute(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
    expected_local: &[u8],
) -> Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        if local_name(name) == expected_local && namespaces.attribute_uri(name) == Some(R_NS) {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

#[allow(clippy::too_many_arguments)]
fn capture_non_visual_child(
    name: &[u8],
    uri: Option<&str>,
    raw: Vec<u8>,
    candidate_name: Option<String>,
    namespaces: &NamespaceBindings,
    drawing: &mut Option<Vec<u8>>,
    drawing_name: &mut Option<String>,
    picture: &mut Option<Vec<u8>>,
    application: &mut Option<ApplicationProperties>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if uri == Some(P_NS) && matches!(name, b"cNvPr" | b"cNvPicPr" | b"nvPr") {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
    }
    match (uri, name) {
        (Some(P_NS), b"cNvPr") if *boundary == 0 && drawing.is_none() => {
            *drawing = Some(raw);
            *drawing_name = candidate_name;
            *boundary = 1;
        }
        (Some(P_NS), b"cNvPicPr") if *boundary == 1 && picture.is_none() => {
            *picture = Some(raw);
            *boundary = 2;
        }
        (Some(P_NS), b"nvPr") if *boundary == 2 && application.is_none() => {
            *application = Some(parse_application_properties(&raw, namespaces)?);
            *boundary = 3;
        }
        (Some(P_NS), b"cNvPr" | b"cNvPicPr" | b"nvPr") => {
            return Err(OxmlError::InvalidValue(
                "p:nvPicPr children must be p:cNvPr, p:cNvPicPr, and p:nvPr in order".to_owned(),
            ));
        }
        _ => raw_children.push(*boundary, raw),
    }
    Ok(())
}

fn validate_blip_fill(xml: &[u8], inherited: &NamespaceBindings) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                if namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    || local_name(start.name().as_ref()) != b"blipFill"
                {
                    return Err(unexpected(&start));
                }
                loop {
                    buffer.clear();
                    match reader.read_event_into(&mut buffer)? {
                        Event::Start(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            validate_blip_fill_child(&child, &child_namespaces)?;
                            capture_element(&mut reader, &child)?;
                        }
                        Event::Empty(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            validate_blip_fill_child(&child, &child_namespaces)?;
                        }
                        Event::End(end) if local_name(end.name().as_ref()) == b"blipFill" => {
                            return Ok(());
                        }
                        Event::Eof => {
                            return Err(OxmlError::MissingElement("closing p:blipFill".to_owned()));
                        }
                        _ => {}
                    }
                }
            }
            Event::Empty(start) => {
                let namespaces = inherited.with_start(&start)?;
                if namespaces.element_uri(start.name().as_ref()) == Some(P_NS)
                    && local_name(start.name().as_ref()) == b"blipFill"
                {
                    return Ok(());
                }
                return Err(unexpected(&start));
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:blipFill".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn validate_blip_fill_child(child: &BytesStart<'_>, namespaces: &NamespaceBindings) -> Result<()> {
    let qualified_name = child.name();
    let name = local_name(qualified_name.as_ref());
    if matches!(name, b"blip" | b"srcRect" | b"stretch" | b"tile")
        && namespaces.element_uri(child.name().as_ref()) != Some(A_NS)
    {
        return Err(OxmlError::InvalidValue(format!(
            "p:blipFill {} must use the DrawingML namespace",
            String::from_utf8_lossy(name)
        )));
    }
    if name == b"blip" && namespaces.element_uri(child.name().as_ref()) == Some(A_NS) {
        for attribute in child.attributes() {
            let attribute = attribute?;
            let attribute_name = attribute.key.as_ref();
            if matches!(local_name(attribute_name), b"embed" | b"link")
                && namespaces.attribute_uri(attribute_name) != Some(R_NS)
            {
                return Err(OxmlError::InvalidValue(format!(
                    "a:blip {} must use the Office relationships namespace",
                    String::from_utf8_lossy(attribute_name)
                )));
            }
        }
    }
    Ok(())
}

fn drawing_error(error: impl std::fmt::Display) -> OxmlError {
    OxmlError::InvalidValue(error.to_string())
}

fn required<T>(value: Option<T>, name: &str) -> Result<T> {
    value.ok_or_else(|| OxmlError::MissingElement(name.to_owned()))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn unexpected(start: &BytesStart<'_>) -> OxmlError {
    OxmlError::InvalidValue(format!(
        "unexpected PresentationML picture element: {}",
        String::from_utf8_lossy(start.name().as_ref())
    ))
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

#[cfg(test)]
mod tests {
    use oxml_core::units::Emu;
    use oxml_drawing::xfrm::{CT_Point2D, CT_PositiveSize2D, CT_Transform2D};

    use super::CT_Picture;

    #[test]
    fn picture_constructor_round_trips_in_schema_order() {
        let mut transform = CT_Transform2D::default();
        transform.offset = Some(CT_Point2D {
            x: Emu(10),
            y: Emu(20),
        });
        transform.extent = Some(CT_PositiveSize2D {
            cx: Emu(30),
            cy: Emu(40),
        });
        let picture = CT_Picture::new(7, "Picture & 7", "rId9", transform).unwrap();

        let xml = picture.to_xml().unwrap();
        let text = String::from_utf8(xml.clone()).unwrap();
        assert!(text.contains("<p:cNvPr id=\"7\" name=\"Picture &amp; 7\"/>"));
        assert!(text.contains("r:embed=\"rId9\""));
        assert!(text.contains("<a:off x=\"10\" y=\"20\"/>"));
        assert!(text.contains("<a:ext cx=\"30\" cy=\"40\"/>"));
        let non_visual = text.find("<p:nvPicPr").unwrap();
        let blip_fill = text.find("<p:blipFill").unwrap();
        let shape_properties = text.find("<p:spPr").unwrap();
        assert!(non_visual < blip_fill);
        assert!(blip_fill < shape_properties);
        assert!(
            text.contains("</a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>")
        );

        let reparsed = CT_Picture::from_xml(&xml).unwrap();
        assert_eq!(reparsed.non_visual_id(), Some(7));
        assert_eq!(
            reparsed.blip_fill.unwrap().blip.unwrap().embed.as_deref(),
            Some("rId9")
        );
        assert_eq!(
            reparsed.shape_properties.transform,
            picture.shape_properties.transform
        );
        assert_eq!(
            reparsed
                .shape_properties
                .preset_geometry
                .as_ref()
                .map(|geometry| geometry.preset.as_str()),
            Some("rect")
        );
    }
}
