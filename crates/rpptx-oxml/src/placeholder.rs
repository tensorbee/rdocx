use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_drawing::namespace::A_NS;
use oxml_drawing::order::OrderedRawChildren;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::namespace::{
    FIXED_SHAPE_TREE_PREFIXES, MC_NS, NamespaceBindings, P_NS, R_NS, all_attributes,
    root_attributes, self_contained_attributes,
};

pub type Result<T> = std::result::Result<T, OxmlError>;

/// The placeholder kinds accepted by PresentationML.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub enum PhType {
    Title,
    Body,
    CenteredTitle,
    Subtitle,
    DateTime,
    SlideNumber,
    Footer,
    Header,
    Object,
    Chart,
    Table,
    ClipArt,
    Diagram,
    Media,
    SlideImage,
    Picture,
    VerticalTitle,
    VerticalBody,
    VerticalObject,
    Other(String),
}

impl PhType {
    fn parse(value: &str) -> Self {
        match value {
            "title" => Self::Title,
            "body" => Self::Body,
            "ctrTitle" => Self::CenteredTitle,
            "subTitle" => Self::Subtitle,
            "dt" => Self::DateTime,
            "sldNum" => Self::SlideNumber,
            "ftr" => Self::Footer,
            "hdr" => Self::Header,
            "obj" => Self::Object,
            "chart" => Self::Chart,
            "tbl" => Self::Table,
            "clipArt" => Self::ClipArt,
            "dgm" => Self::Diagram,
            "media" => Self::Media,
            "sldImg" => Self::SlideImage,
            "pic" => Self::Picture,
            "vertTitle" => Self::VerticalTitle,
            "vertBody" => Self::VerticalBody,
            "vertObj" => Self::VerticalObject,
            other => Self::Other(other.to_owned()),
        }
    }

    /// Returns the `ST_PlaceholderType` token as written in the file.
    pub fn as_str(&self) -> &str {
        match self {
            Self::Title => "title",
            Self::Body => "body",
            Self::CenteredTitle => "ctrTitle",
            Self::Subtitle => "subTitle",
            Self::DateTime => "dt",
            Self::SlideNumber => "sldNum",
            Self::Footer => "ftr",
            Self::Header => "hdr",
            Self::Object => "obj",
            Self::Chart => "chart",
            Self::Table => "tbl",
            Self::ClipArt => "clipArt",
            Self::Diagram => "dgm",
            Self::Media => "media",
            Self::SlideImage => "sldImg",
            Self::Picture => "pic",
            Self::VerticalTitle => "vertTitle",
            Self::VerticalBody => "vertBody",
            Self::VerticalObject => "vertObj",
            Self::Other(value) => value,
        }
    }

    fn equivalent_to(&self, other: &Self) -> bool {
        matches!(
            (self, other),
            (
                Self::Title | Self::CenteredTitle,
                Self::Title | Self::CenteredTitle
            ) | (
                Self::Body | Self::Subtitle | Self::Object,
                Self::Body | Self::Subtitle | Self::Object
            )
        ) || self == other
    }
}

/// A parsed `p:ph` element with producer extensions retained.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_Placeholder {
    pub ph_type: Option<PhType>,
    pub idx: Option<u32>,
    raw_attributes: Vec<(String, String)>,
    raw_children: Vec<Vec<u8>>,
}

impl CT_Placeholder {
    /// Creates a minimal placeholder with its inheritance identity.
    pub fn new(ph_type: Option<PhType>, idx: Option<u32>) -> Self {
        Self {
            ph_type,
            idx,
            raw_attributes: Vec::new(),
            raw_children: Vec::new(),
        }
    }

    /// Parses a complete placeholder with any prefix bound to PresentationML.
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
                    validate_root(&start, &namespaces)?;
                    return Self::from_element(&mut reader, &start, inherited);
                }
                Event::Empty(start) => {
                    let namespaces =
                        NamespaceBindings::from_entries(inherited).with_start(&start)?;
                    validate_root(&start, &namespaces)?;
                    return Self::from_start(&start, Vec::new(), inherited);
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:ph".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_element(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        inherited: &[(String, String)],
    ) -> Result<Self> {
        let mut raw_children = Vec::new();
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => raw_children.push(capture_element(reader, &child)?),
                Event::Empty(child) => raw_children.push(capture_empty_element(&child)?),
                Event::End(end) if local_name(end.name().as_ref()) == b"ph" => break,
                Event::Eof => return Err(OxmlError::MissingElement("closing p:ph".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
        Self::from_start(start, raw_children, inherited)
    }

    fn from_start(
        start: &BytesStart<'_>,
        raw_children: Vec<Vec<u8>>,
        inherited: &[(String, String)],
    ) -> Result<Self> {
        let mut ph_type = None;
        let mut idx = None;
        let mut raw_attributes =
            self_contained_attributes(start, FIXED_SHAPE_TREE_PREFIXES, inherited)?;
        raw_attributes.retain(|(name, _)| !matches!(name.as_str(), "type" | "idx"));
        for (name, value) in all_attributes(start)? {
            match name.as_str() {
                "type" => ph_type = Some(PhType::parse(&value)),
                "idx" => {
                    idx = Some(value.parse::<u32>().map_err(|error| {
                        OxmlError::InvalidValue(format!("invalid p:ph@idx {value}: {error}"))
                    })?)
                }
                _ if is_xmlns(&name) || raw_attributes.iter().any(|(raw, _)| raw == &name) => {}
                _ => raw_attributes.push((name, value)),
            }
        }
        Ok(Self {
            ph_type,
            idx,
            raw_attributes,
            raw_children,
        })
    }

    /// Returns the explicit type or the PowerPoint body default.
    pub fn effective_type(&self) -> PhType {
        self.ph_type.clone().unwrap_or(PhType::Body)
    }

    /// Returns the presence-sensitive key used for inheritance matching.
    pub fn key(&self) -> PlaceholderKey {
        PlaceholderKey {
            ph_type: self.effective_type(),
            idx: self.idx,
        }
    }

    /// Serialises a self-contained placeholder with the fixed `p:` prefix.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer, true)?;
        Ok(writer.into_inner())
    }

    pub(crate) fn write_xml<W: Write>(
        &self,
        writer: &mut Writer<W>,
        declare_namespace: bool,
    ) -> Result<()> {
        let mut start = BytesStart::new("p:ph");
        if declare_namespace {
            start.push_attribute(("xmlns:p", P_NS));
            start.push_attribute(("xmlns:a", A_NS));
            start.push_attribute(("xmlns:r", R_NS));
            start.push_attribute(("xmlns:mc", MC_NS));
        }
        if let Some(ph_type) = &self.ph_type {
            start.push_attribute(("type", ph_type.as_str()));
        }
        if let Some(idx) = self.idx {
            let value = idx.to_string();
            start.push_attribute(("idx", value.as_str()));
        }
        for (name, value) in &self.raw_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }
        if self.raw_children.is_empty() {
            writer.write_event(Event::Empty(start))?;
            return Ok(());
        }
        writer.write_event(Event::Start(start))?;
        for child in &self.raw_children {
            writer.get_mut().write_all(child)?;
        }
        writer.write_event(Event::End(BytesEnd::new("p:ph")))?;
        Ok(())
    }
}

/// The placeholder identity used by slide, layout, and master matching.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct PlaceholderKey {
    pub ph_type: PhType,
    pub idx: Option<u32>,
}

pub(crate) struct ApplicationProperties {
    pub(crate) placeholder: Option<CT_Placeholder>,
    pub(crate) raw_attributes: Vec<(String, String)>,
    pub(crate) raw_children: OrderedRawChildren,
}

impl PlaceholderKey {
    /// Applies PowerPoint's index-first, type-fallback matching rule.
    pub fn matches(&self, other: &Self) -> bool {
        match (self.idx, other.idx) {
            (Some(left), Some(right)) => left == right,
            _ => self.ph_type.equivalent_to(&other.ph_type),
        }
    }
}

pub(crate) fn parse_application_properties(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<ApplicationProperties> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                let raw_attributes = root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?;
                let mut placeholder = None;
                let mut raw_children = OrderedRawChildren::default();
                let mut boundary = 0usize;
                loop {
                    buffer.clear();
                    match reader.read_event_into(&mut buffer)? {
                        Event::Start(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            let is_placeholder =
                                child_namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                                    && local_name(child.name().as_ref()) == b"ph";
                            let raw = capture_element(&mut reader, &child)?;
                            capture_application_child(
                                is_placeholder,
                                raw,
                                &child_namespaces,
                                &mut placeholder,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::Empty(child) => {
                            let child_namespaces = namespaces.with_start(&child)?;
                            let is_placeholder =
                                child_namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                                    && local_name(child.name().as_ref()) == b"ph";
                            let raw = capture_empty_element(&child)?;
                            capture_application_child(
                                is_placeholder,
                                raw,
                                &child_namespaces,
                                &mut placeholder,
                                &mut raw_children,
                                &mut boundary,
                            )?;
                        }
                        Event::End(end) if local_name(end.name().as_ref()) == b"nvPr" => break,
                        Event::Eof => {
                            return Err(OxmlError::MissingElement("closing p:nvPr".to_owned()));
                        }
                        _ => {}
                    }
                }
                return Ok(ApplicationProperties {
                    placeholder,
                    raw_attributes,
                    raw_children,
                });
            }
            Event::Empty(start) => {
                return Ok(ApplicationProperties {
                    placeholder: None,
                    raw_attributes: root_attributes(&start, FIXED_SHAPE_TREE_PREFIXES)?,
                    raw_children: OrderedRawChildren::default(),
                });
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:nvPr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn capture_application_child(
    is_placeholder: bool,
    raw: Vec<u8>,
    namespaces: &NamespaceBindings,
    placeholder: &mut Option<CT_Placeholder>,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if is_placeholder {
        namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)?;
        if placeholder.is_some() {
            return Err(OxmlError::InvalidValue("duplicate p:ph".to_owned()));
        }
        *placeholder = Some(CT_Placeholder::from_fragment(&raw, &namespaces.entries())?);
        *boundary = 1;
    } else {
        raw_children.push(*boundary, raw);
    }
    Ok(())
}

fn validate_root(start: &BytesStart<'_>, namespaces: &NamespaceBindings) -> Result<()> {
    if local_name(start.name().as_ref()) != b"ph"
        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
    {
        return Err(OxmlError::InvalidValue(
            "expected a PresentationML p:ph element".to_owned(),
        ));
    }
    namespaces.reject_writer_conflicts(FIXED_SHAPE_TREE_PREFIXES)
}

fn is_xmlns(name: &str) -> bool {
    name == "xmlns" || name.starts_with("xmlns:")
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}
