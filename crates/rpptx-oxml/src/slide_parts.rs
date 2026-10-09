use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_drawing::color::{ColorChoice, ColorMap, ColorMapSlot, ThemeColorSlot};
use oxml_drawing::fill::Fill;
use oxml_drawing::namespace::A_NS;
use oxml_drawing::order::OrderedRawChildren;
use oxml_drawing::text::CT_TextListStyle;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer};

use crate::namespace::{
    FIXED_MODEL_PREFIXES, MC_NS, NamespaceBindings, P_NS, R_NS, all_attributes, is_fixed_xmlns,
    root_attributes,
};

const P188_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const COMMENT_EXTENSION_URI: &str = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}";
use crate::shape_tree::CT_ShapeTree;
use crate::timing::{CT_SlideTransition, CT_Timing};

pub type Result<T> = std::result::Result<T, OxmlError>;
type RawAttributes = Vec<(String, String)>;

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ColorMapOverrideKind {
    Master,
    Override(ColorMap),
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_ColorMapOverride {
    pub kind: ColorMapOverrideKind,
    raw_attributes: RawAttributes,
    mapping_attributes: RawAttributes,
    mapping_children: OrderedRawChildren,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_CommonSlideData {
    pub name: Option<String>,
    pub background: Option<CT_Background>,
    pub shape_tree: CT_ShapeTree,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

/// The renderer-visible choice carried by one preserved `p:bg` subtree.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum BackgroundRendering {
    Properties(Option<Box<Fill>>),
    Reference {
        index: u32,
        color: Option<ColorChoice>,
    },
    Unsupported(&'static str),
}

/// A slide background whose original bytes remain the sole serialisation source.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_Background {
    raw_xml: Vec<u8>,
    rendering: BackgroundRendering,
    fill_range: Option<(usize, usize)>,
}

impl CT_Background {
    /// Builds a canonical direct-fill slide background.
    pub fn from_fill(fill: Fill) -> Result<Self> {
        let mut writer = Writer::new(Vec::new());
        writer.write_event(Event::Start(BytesStart::new("p:bg")))?;
        writer.write_event(Event::Start(BytesStart::new("p:bgPr")))?;
        let fill_start = writer.get_ref().len();
        fill.write_xml(&mut writer)
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        let fill_end = writer.get_ref().len();
        writer.write_event(Event::End(BytesEnd::new("p:bgPr")))?;
        writer.write_event(Event::End(BytesEnd::new("p:bg")))?;
        Ok(Self {
            raw_xml: writer.into_inner(),
            rendering: BackgroundRendering::Properties(Some(Box::new(fill))),
            fill_range: Some((fill_start, fill_end)),
        })
    }

    /// Replaces only an existing direct fill inside the preserved background.
    fn replace_fill(&mut self, fill: Fill) -> Result<()> {
        let Some((start, end)) = self.fill_range else {
            return Err(OxmlError::InvalidValue(
                "cannot replace a background that has no safe direct fill range".to_owned(),
            ));
        };
        let replacement = fill
            .to_xml()
            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
        self.raw_xml.splice(start..end, replacement.iter().copied());
        self.fill_range = Some((start, start + replacement.len()));
        self.rendering = BackgroundRendering::Properties(Some(Box::new(fill)));
        Ok(())
    }

    /// Returns the original subtree used as the sole serialisation source.
    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }

    /// Returns the typed projection used only by inheritance resolution.
    pub fn rendering(&self) -> &BackgroundRendering {
        &self.rendering
    }

    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"bg"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    let rendering = parse_background_rendering(&mut reader, &namespaces)?;
                    let fill_range =
                        if matches!(rendering, BackgroundRendering::Properties(Some(_))) {
                            direct_background_fill_range(xml, inherited)?
                        } else {
                            None
                        };
                    return Ok(Self {
                        raw_xml: xml.to_vec(),
                        rendering,
                        fill_range,
                    });
                }
                Event::Empty(start) => {
                    return Err(OxmlError::MissingElement(format!(
                        "{} requires p:bgPr or p:bgRef",
                        String::from_utf8_lossy(start.name().as_ref())
                    )));
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:bg".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CT_MasterTextStyles {
    pub title_style: CT_TextListStyle,
    pub body_style: CT_TextListStyle,
    pub other_style: CT_TextListStyle,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

/// Presence-sensitive header and footer visibility inherited by a slide.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_HeaderFooter {
    pub slide_number: Option<bool>,
    pub header: Option<bool>,
    pub footer: Option<bool>,
    pub date_time: Option<bool>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

impl CT_HeaderFooter {
    /// Creates a container with these flags, `None` keeping the enabled default.
    pub fn new(
        slide_number: Option<bool>,
        header: Option<bool>,
        footer: Option<bool>,
        date_time: Option<bool>,
    ) -> Self {
        Self {
            slide_number,
            header,
            footer,
            date_time,
            raw_attributes: RawAttributes::new(),
            raw_children: OrderedRawChildren::default(),
        }
    }

    pub fn slide_number_enabled(&self) -> bool {
        self.slide_number.unwrap_or(true)
    }

    pub fn footer_enabled(&self) -> bool {
        self.footer.unwrap_or(true)
    }

    pub fn header_enabled(&self) -> bool {
        self.header.unwrap_or(true)
    }

    pub fn date_time_enabled(&self) -> bool {
        self.date_time.unwrap_or(true)
    }
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_Slide {
    pub common_slide_data: CT_CommonSlideData,
    pub color_map_override: Option<CT_ColorMapOverride>,
    pub transition: Option<CT_SlideTransition>,
    pub timing: Option<CT_Timing>,
    pub show: Option<bool>,
    pub show_master_shapes: Option<bool>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_SlideLayout {
    pub common_slide_data: CT_CommonSlideData,
    pub color_map_override: Option<CT_ColorMapOverride>,
    pub timing: Option<CT_Timing>,
    pub transition: Option<CT_SlideTransition>,
    pub show_master_shapes: Option<bool>,
    pub header_footer: Option<CT_HeaderFooter>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug, PartialEq)]
pub struct CT_SlideMaster {
    pub common_slide_data: CT_CommonSlideData,
    pub color_map: ColorMap,
    pub transition: Option<CT_SlideTransition>,
    pub timing: Option<CT_Timing>,
    pub text_styles: Option<CT_MasterTextStyles>,
    pub header_footer: Option<CT_HeaderFooter>,
    raw_attributes: RawAttributes,
    color_map_attributes: RawAttributes,
    color_map_children: OrderedRawChildren,
    raw_children: OrderedRawChildren,
}

#[derive(Clone, Copy)]
enum RootKind {
    Slide,
    Layout,
    Master,
}

impl RootKind {
    const fn local_name(self) -> &'static [u8] {
        match self {
            Self::Slide => b"sld",
            Self::Layout => b"sldLayout",
            Self::Master => b"sldMaster",
        }
    }

    const fn tag(self) -> &'static str {
        match self {
            Self::Slide => "p:sld",
            Self::Layout => "p:sldLayout",
            Self::Master => "p:sldMaster",
        }
    }

    fn raw_boundary(self, name: &[u8], current: usize) -> (usize, usize) {
        let before = match self {
            Self::Slide => match name {
                b"transition" => Some(2),
                b"timing" => Some(3),
                b"extLst" => Some(4),
                _ => None,
            },
            Self::Layout => match name {
                b"transition" => Some(2),
                b"timing" => Some(3),
                b"hf" => Some(4),
                b"extLst" => Some(5),
                _ => None,
            },
            Self::Master => match name {
                b"sldLayoutIdLst" => Some(2),
                b"transition" => Some(3),
                b"timing" => Some(4),
                b"hf" => Some(5),
                b"extLst" => Some(7),
                _ => None,
            },
        };
        before.map_or((current, current), |before| {
            (current.max(before), current.max(before + 1))
        })
    }
}

#[derive(Default)]
struct ParsedRoot {
    common_slide_data: Option<CT_CommonSlideData>,
    color_map_override: Option<CT_ColorMapOverride>,
    color_map: Option<ParsedColorMap>,
    text_styles: Option<Box<CT_MasterTextStyles>>,
    header_footer: Option<CT_HeaderFooter>,
    transition: Option<CT_SlideTransition>,
    timing: Option<CT_Timing>,
    show: Option<bool>,
    show_master_shapes: Option<bool>,
    raw_attributes: RawAttributes,
    raw_children: OrderedRawChildren,
    boundary: usize,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct ParsedColorMap {
    pub(crate) value: ColorMap,
    pub(crate) raw_attributes: RawAttributes,
    pub(crate) raw_children: OrderedRawChildren,
}

impl CT_Slide {
    /// Creates a minimal slide that inherits its colour map from the master.
    pub fn new(shape_tree: CT_ShapeTree) -> Self {
        Self {
            common_slide_data: CT_CommonSlideData {
                name: None,
                background: None,
                shape_tree,
                raw_attributes: Vec::new(),
                raw_children: OrderedRawChildren::default(),
            },
            color_map_override: Some(CT_ColorMapOverride {
                kind: ColorMapOverrideKind::Master,
                raw_attributes: Vec::new(),
                mapping_attributes: Vec::new(),
                mapping_children: OrderedRawChildren::default(),
                raw_children: OrderedRawChildren::default(),
            }),
            transition: None,
            timing: None,
            show: None,
            show_master_shapes: None,
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        }
    }

    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let parsed = parse_root(xml, RootKind::Slide)?;
        Ok(Self {
            common_slide_data: required(parsed.common_slide_data, "p:cSld")?,
            color_map_override: parsed.color_map_override,
            transition: parsed.transition,
            timing: parsed.timing,
            show: parsed.show,
            show_master_shapes: parsed.show_master_shapes,
            raw_attributes: parsed.raw_attributes,
            raw_children: parsed.raw_children,
        })
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        write_slide_like(
            RootKind::Slide,
            &self.common_slide_data,
            self.color_map_override.as_ref(),
            self.transition.as_ref(),
            self.timing.as_ref(),
            self.show,
            self.show_master_shapes,
            None,
            &self.raw_attributes,
            &self.raw_children,
        )
    }

    /// Returns whether any retained slide XML contains a PresentationML placeholder.
    ///
    /// The scan covers typed children and preserved compatibility branches by
    /// resolving expanded element names across the complete serialized slide.
    pub fn contains_placeholder(&self) -> Result<bool> {
        contains_presentationml_placeholder(&self.to_xml()?)
    }

    /// Returns whether this slide is hidden from the slideshow.
    pub fn hidden(&self) -> bool {
        !self.show.unwrap_or(true)
    }

    /// Sets hidden state through the inverse `p:sld/@show` flag.
    pub fn set_hidden(&mut self, hidden: bool) {
        self.show = Some(!hidden);
    }

    /// Returns whether the slide carries any explicit `p:bg` element.
    pub fn has_explicit_background(&self) -> bool {
        self.common_slide_data.background.is_some()
    }

    /// Replaces an absent or direct-fill background without authoring `p:bgRef`.
    pub fn set_background(&mut self, fill: Fill) -> Result<()> {
        if let Some(background) = &mut self.common_slide_data.background {
            return background.replace_fill(fill);
        }
        self.common_slide_data.background = Some(CT_Background::from_fill(fill)?);
        Ok(())
    }

    /// Removes only a direct-fill background and preserves all other choices.
    pub fn clear_background(&mut self) {
        if self
            .common_slide_data
            .background
            .as_ref()
            .is_some_and(|background| {
                matches!(background.rendering(), BackgroundRendering::Properties(_))
            })
        {
            self.common_slide_data.background = None;
        }
    }

    /// Returns the modern-comment relationship referenced by the slide extension.
    pub fn modern_comment_relationship_id(&self) -> Result<Option<String>> {
        find_modern_comment_relationship(&self.to_xml()?)
    }

    /// Adds the schema-final modern-comment extension if it is absent.
    pub fn ensure_modern_comment_relationship(&mut self, relationship_id: &str) -> Result<()> {
        let xml = self.to_xml()?;
        if let Some(existing) = find_modern_comment_relationship(&xml)? {
            if existing == relationship_id {
                return Ok(());
            }
            return Err(OxmlError::InvalidValue(format!(
                "slide modern-comment extension references {existing}, not {relationship_id}"
            )));
        }
        let insertion = find_extension_list_closing(&xml)?;
        let extension = format!(
            "<p:ext uri=\"{COMMENT_EXTENSION_URI}\"><p188:commentRel xmlns:p188=\"{P188_NS}\" r:id=\"{relationship_id}\"/></p:ext>"
        );
        let rewritten = if let Some((start, end)) = insertion {
            let mut rewritten = Vec::with_capacity(xml.len() + extension.len());
            rewritten.extend_from_slice(&xml[..start]);
            if start == end {
                rewritten.extend_from_slice(extension.as_bytes());
            } else {
                let empty = &xml[start..end];
                let slash = empty
                    .windows(2)
                    .rposition(|window| window == b"/>")
                    .ok_or_else(|| {
                        OxmlError::InvalidValue("malformed empty p:extLst".to_owned())
                    })?;
                let name_end = empty[1..]
                    .iter()
                    .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
                    .map(|index| index + 1)
                    .ok_or_else(|| {
                        OxmlError::InvalidValue("malformed empty p:extLst".to_owned())
                    })?;
                rewritten.extend_from_slice(&empty[..slash]);
                rewritten.extend_from_slice(b">");
                rewritten.extend_from_slice(extension.as_bytes());
                rewritten.extend_from_slice(b"</");
                rewritten.extend_from_slice(&empty[1..name_end]);
                rewritten.extend_from_slice(b">");
            }
            rewritten.extend_from_slice(&xml[end..]);
            rewritten
        } else {
            let closing = xml
                .windows(b"</p:sld>".len())
                .position(|window| window == b"</p:sld>")
                .ok_or_else(|| OxmlError::MissingElement("closing p:sld".to_owned()))?;
            let mut rewritten = Vec::with_capacity(xml.len() + extension.len() + 21);
            rewritten.extend_from_slice(&xml[..closing]);
            rewritten.extend_from_slice(b"<p:extLst>");
            rewritten.extend_from_slice(extension.as_bytes());
            rewritten.extend_from_slice(b"</p:extLst>");
            rewritten.extend_from_slice(&xml[closing..]);
            rewritten
        };
        *self = Self::from_xml(&rewritten)?;
        Ok(())
    }
}

fn contains_presentationml_placeholder(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_resolved_event_into(&mut buffer)? {
            (
                ResolveResult::Bound(Namespace(namespace)),
                Event::Start(element) | Event::Empty(element),
            ) if namespace == P_NS.as_bytes() && local_name(element.name().as_ref()) == b"ph" => {
                return Ok(true);
            }
            (_, Event::Eof) => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn find_modern_comment_relationship(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut scopes = vec![NamespaceBindings::default()];
    let mut depth = 0usize;
    let mut root_extension_list_depth = None;
    let mut comment_extension_depth = None;
    let mut found = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let parent = scopes.last().ok_or_else(|| {
                    OxmlError::InvalidValue("missing comment namespace scope".to_owned())
                })?;
                let scope = parent.with_start(&element)?;
                if depth == 1
                    && local_name(element.name().as_ref()) == b"extLst"
                    && scope.element_uri(element.name().as_ref()) == Some(P_NS)
                {
                    root_extension_list_depth = Some(depth + 1);
                } else if root_extension_list_depth == Some(depth)
                    && local_name(element.name().as_ref()) == b"ext"
                    && scope.element_uri(element.name().as_ref()) == Some(P_NS)
                    && extension_uri(&element)?.as_deref() == Some(COMMENT_EXTENSION_URI)
                {
                    comment_extension_depth = Some(depth + 1);
                }
                if local_name(element.name().as_ref()) == b"commentRel"
                    && scope.element_uri(element.name().as_ref()) == Some(P188_NS)
                    && comment_extension_depth == Some(depth)
                {
                    let relationship = relationship_id_attribute(&element, &scope)?;
                    if found.replace(relationship).is_some() {
                        return Err(duplicate("modern comment relationship"));
                    }
                }
                scopes.push(scope);
                depth += 1;
            }
            Event::Empty(element) => {
                let parent = scopes.last().ok_or_else(|| {
                    OxmlError::InvalidValue("missing comment namespace scope".to_owned())
                })?;
                let scope = parent.with_start(&element)?;
                if local_name(element.name().as_ref()) == b"commentRel"
                    && scope.element_uri(element.name().as_ref()) == Some(P188_NS)
                    && comment_extension_depth == Some(depth)
                {
                    let relationship = relationship_id_attribute(&element, &scope)?;
                    if found.replace(relationship).is_some() {
                        return Err(duplicate("modern comment relationship"));
                    }
                }
            }
            Event::End(_) => {
                if depth == 0 || scopes.len() == 1 {
                    return Err(OxmlError::InvalidValue(
                        "slide XML has an unmatched closing tag".to_owned(),
                    ));
                }
                if comment_extension_depth == Some(depth) {
                    comment_extension_depth = None;
                }
                if root_extension_list_depth == Some(depth) {
                    root_extension_list_depth = None;
                }
                depth -= 1;
                scopes.pop();
            }
            Event::Eof if depth == 0 && scopes.len() == 1 => return Ok(found),
            Event::Eof => {
                return Err(OxmlError::InvalidValue(
                    "slide XML ended before its root closed".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn extension_uri(element: &BytesStart<'_>) -> Result<Option<String>> {
    Ok(all_attributes(element)?
        .into_iter()
        .find_map(|(name, value)| (name == "uri").then_some(value)))
}

fn relationship_id_attribute(
    element: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<String> {
    for (name, value) in all_attributes(element)? {
        if local_name(name.as_bytes()) == b"id"
            && namespaces.attribute_uri(name.as_bytes()) == Some(R_NS)
        {
            return Ok(value);
        }
    }
    Err(OxmlError::MissingElement(
        "p188:commentRel requires r:id".to_owned(),
    ))
}

fn find_extension_list_closing(xml: &[u8]) -> Result<Option<(usize, usize)>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut scopes = vec![NamespaceBindings::default()];
    let mut depth = 0usize;
    let mut extension_depth = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let parent = scopes.last().ok_or_else(|| {
                    OxmlError::InvalidValue("missing extension namespace scope".to_owned())
                })?;
                let scope = parent.with_start(&element)?;
                depth += 1;
                if depth == 2
                    && local_name(element.name().as_ref()) == b"extLst"
                    && scope.element_uri(element.name().as_ref()) == Some(P_NS)
                {
                    extension_depth = Some(depth);
                }
                scopes.push(scope);
            }
            Event::Empty(element) => {
                let parent = scopes.last().ok_or_else(|| {
                    OxmlError::InvalidValue("missing extension namespace scope".to_owned())
                })?;
                let scope = parent.with_start(&element)?;
                if depth == 1
                    && local_name(element.name().as_ref()) == b"extLst"
                    && scope.element_uri(element.name().as_ref()) == Some(P_NS)
                {
                    return Ok(Some((event_start, reader.buffer_position() as usize)));
                }
            }
            Event::End(_) => {
                if extension_depth == Some(depth) {
                    return Ok(Some((event_start, event_start)));
                }
                if depth == 0 || scopes.len() == 1 {
                    return Err(OxmlError::InvalidValue(
                        "slide XML has an unmatched closing tag".to_owned(),
                    ));
                }
                depth -= 1;
                scopes.pop();
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn direct_background_fill_range(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<(usize, usize)>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut scopes = vec![inherited.clone()];
    let mut depth = 0usize;
    let mut properties_depth = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let parent = scopes.last().ok_or_else(|| {
                    OxmlError::InvalidValue("missing background namespace scope".to_owned())
                })?;
                let scope = parent.with_start(&element)?;
                if properties_depth == Some(depth)
                    && scope.element_uri(element.name().as_ref()) == Some(A_NS)
                    && is_fill_name(local_name(element.name().as_ref()))
                {
                    reader.read_to_end(element.name())?;
                    return Ok(Some((event_start, reader.buffer_position() as usize)));
                }
                depth += 1;
                if depth == 2
                    && scope.element_uri(element.name().as_ref()) == Some(P_NS)
                    && local_name(element.name().as_ref()) == b"bgPr"
                {
                    properties_depth = Some(depth);
                }
                scopes.push(scope);
            }
            Event::Empty(element) => {
                let parent = scopes.last().ok_or_else(|| {
                    OxmlError::InvalidValue("missing background namespace scope".to_owned())
                })?;
                let scope = parent.with_start(&element)?;
                if properties_depth == Some(depth)
                    && scope.element_uri(element.name().as_ref()) == Some(A_NS)
                    && is_fill_name(local_name(element.name().as_ref()))
                {
                    return Ok(Some((event_start, reader.buffer_position() as usize)));
                }
            }
            Event::End(element) => {
                if properties_depth == Some(depth) && local_name(element.name().as_ref()) == b"bgPr"
                {
                    properties_depth = None;
                }
                if depth == 0 || scopes.len() == 1 {
                    return Err(OxmlError::InvalidValue(
                        "background XML has an unmatched closing tag".to_owned(),
                    ));
                }
                depth -= 1;
                scopes.pop();
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

impl CT_SlideLayout {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let parsed = parse_root(xml, RootKind::Layout)?;
        Ok(Self {
            common_slide_data: required(parsed.common_slide_data, "p:cSld")?,
            color_map_override: parsed.color_map_override,
            timing: parsed.timing,
            transition: parsed.transition,
            show_master_shapes: parsed.show_master_shapes,
            header_footer: parsed.header_footer,
            raw_attributes: parsed.raw_attributes,
            raw_children: parsed.raw_children,
        })
    }

    /// Returns the `type` attribute, such as `title` for a title slide layout.
    pub fn layout_type(&self) -> Option<&str> {
        self.raw_attributes
            .iter()
            .find(|(name, _)| name == "type")
            .map(|(_, value)| value.as_str())
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        write_slide_like(
            RootKind::Layout,
            &self.common_slide_data,
            self.color_map_override.as_ref(),
            self.transition.as_ref(),
            self.timing.as_ref(),
            None,
            self.show_master_shapes,
            self.header_footer.as_ref(),
            &self.raw_attributes,
            &self.raw_children,
        )
    }
}

impl CT_SlideMaster {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let parsed = parse_root(xml, RootKind::Master)?;
        let color_map = required(parsed.color_map, "p:clrMap")?;
        Ok(Self {
            common_slide_data: required(parsed.common_slide_data, "p:cSld")?,
            color_map: color_map.value,
            transition: parsed.transition,
            timing: parsed.timing,
            text_styles: parsed.text_styles.map(|styles| *styles),
            header_footer: parsed.header_footer,
            raw_attributes: parsed.raw_attributes,
            color_map_attributes: color_map.raw_attributes,
            color_map_children: color_map.raw_children,
            raw_children: parsed.raw_children,
        })
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        write_root_start(
            &mut writer,
            RootKind::Master,
            None,
            None,
            &self.raw_attributes,
        )?;
        emit_raw(&mut writer, self.raw_children.at(0))?;
        self.common_slide_data.write_xml(&mut writer)?;
        emit_raw(&mut writer, self.raw_children.at(1))?;
        write_color_map(
            &mut writer,
            "p:clrMap",
            &self.color_map,
            &self.color_map_attributes,
            &self.color_map_children,
        )?;
        emit_raw(&mut writer, self.raw_children.at(2))?;
        emit_raw(&mut writer, self.raw_children.at(3))?;
        if let Some(transition) = &self.transition {
            transition.write_xml(&mut writer)?;
        }
        emit_raw(&mut writer, self.raw_children.at(4))?;
        if let Some(timing) = &self.timing {
            timing.write_xml(&mut writer)?;
        }
        emit_raw(&mut writer, self.raw_children.at(5))?;
        if let Some(header_footer) = &self.header_footer {
            header_footer.write_xml(&mut writer)?;
        }
        emit_raw(&mut writer, self.raw_children.at(6))?;
        if let Some(styles) = &self.text_styles {
            styles.write_xml(&mut writer)?;
        }
        emit_raw(&mut writer, self.raw_children.at(7))?;
        emit_raw(&mut writer, self.raw_children.at(8))?;
        writer.write_event(Event::End(BytesEnd::new(RootKind::Master.tag())))?;
        Ok(writer.into_inner())
    }
}

/// The raw boundary that holds a master's `p:sldLayoutIdLst`.
const MASTER_LAYOUT_ID_BOUNDARY: usize = 2;

impl CT_SlideMaster {
    /// Returns the `p:sldLayoutIdLst` entries as `(id, relationship id)`
    /// pairs, in document order.
    pub fn slide_layout_ids(&self) -> Result<Vec<(u32, String)>> {
        let Some(list) = self
            .raw_children
            .at(MASTER_LAYOUT_ID_BOUNDARY)
            .find(|xml| raw_root_local_name(xml) == b"sldLayoutIdLst")
        else {
            return Ok(Vec::new());
        };
        Ok(layout_id_entries(list)?
            .into_iter()
            .map(|(id, relationship_id, _)| (id, relationship_id))
            .collect())
    }

    /// Replaces the `p:sldLayoutIdLst` entries. An entry whose id and
    /// relationship id were already listed keeps its original XML.
    pub fn set_slide_layout_ids(&mut self, entries: &[(u32, String)]) -> Result<()> {
        let existing = self
            .raw_children
            .at(MASTER_LAYOUT_ID_BOUNDARY)
            .find(|xml| raw_root_local_name(xml) == b"sldLayoutIdLst")
            .map(layout_id_entries)
            .transpose()?
            .unwrap_or_default();
        let mut list = b"<p:sldLayoutIdLst>".to_vec();
        for (id, relationship_id) in entries {
            if let Some((_, _, raw)) = existing.iter().find(|(old_id, old_relationship, _)| {
                old_id == id && old_relationship == relationship_id
            }) {
                list.extend_from_slice(raw);
                continue;
            }
            let mut writer = Writer::new(Vec::new());
            let mut entry = BytesStart::new("p:sldLayoutId");
            entry.push_attribute(("id", id.to_string().as_str()));
            entry.push_attribute(("r:id", relationship_id.as_str()));
            writer.write_event(Event::Empty(entry))?;
            list.extend_from_slice(&writer.into_inner());
        }
        list.extend_from_slice(b"</p:sldLayoutIdLst>");
        self.raw_children
            .retain(|xml| raw_root_local_name(xml) != b"sldLayoutIdLst");
        self.raw_children.push(MASTER_LAYOUT_ID_BOUNDARY, list);
        Ok(())
    }
}

/// The local name of a captured element's root tag.
fn raw_root_local_name(xml: &[u8]) -> &[u8] {
    let start = xml
        .iter()
        .position(|byte| *byte == b'<')
        .map_or(0, |at| at + 1);
    let rest = &xml[start..];
    let end = rest
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))
        .unwrap_or(rest.len());
    local_name(&rest[..end])
}

/// Reads each `p:sldLayoutId` of a captured list as its id, relationship id
/// and original XML.
fn layout_id_entries(list: &[u8]) -> Result<Vec<(u32, String, Vec<u8>)>> {
    let mut reader = Reader::from_reader(list);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut entries = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer)?.into_owned();
        match event {
            Event::Start(start) if depth == 1 => {
                let attributes = all_attributes(&start)?;
                let raw = capture_element(&mut reader, &start)?;
                if local_name(start.name().as_ref()) == b"sldLayoutId" {
                    entries.push(layout_id_entry(&attributes, raw)?);
                }
            }
            Event::Empty(start) if depth == 1 => {
                let attributes = all_attributes(&start)?;
                let raw = capture_empty_element(&start)?;
                if local_name(start.name().as_ref()) == b"sldLayoutId" {
                    entries.push(layout_id_entry(&attributes, raw)?);
                }
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(entries)
}

fn layout_id_entry(
    attributes: &[(String, String)],
    raw: Vec<u8>,
) -> Result<(u32, String, Vec<u8>)> {
    let id = attributes
        .iter()
        .find(|(name, _)| name == "id")
        .and_then(|(_, value)| value.parse::<u32>().ok())
        .ok_or_else(|| {
            OxmlError::InvalidValue("p:sldLayoutId requires a numeric @id".to_owned())
        })?;
    let relationship_id = attributes
        .iter()
        .find(|(name, _)| name.ends_with(":id"))
        .map(|(_, value)| value.clone())
        .ok_or_else(|| OxmlError::MissingElement("p:sldLayoutId/@r:id".to_owned()))?;
    Ok((id, relationship_id, raw))
}

fn parse_root(xml: &[u8], kind: RootKind) -> Result<ParsedRoot> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = NamespaceBindings::default().with_start(&start)?;
                if local_name(start.name().as_ref()) != kind.local_name()
                    || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                {
                    return Err(unexpected(&start));
                }
                namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
                return parse_root_children(&mut reader, &start, &namespaces, kind);
            }
            Event::Empty(start) => {
                let namespaces = NamespaceBindings::default().with_start(&start)?;
                if local_name(start.name().as_ref()) != kind.local_name()
                    || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                {
                    return Err(unexpected(&start));
                }
                return Err(OxmlError::MissingElement(format!(
                    "{} requires p:cSld",
                    kind.tag()
                )));
            }
            Event::Eof => return Err(OxmlError::MissingElement(kind.tag().to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_root_children(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    root_namespaces: &NamespaceBindings,
    kind: RootKind,
) -> Result<ParsedRoot> {
    let mut parsed = ParsedRoot {
        raw_attributes: root_attributes(start, FIXED_MODEL_PREFIXES)?,
        ..ParsedRoot::default()
    };
    if !matches!(kind, RootKind::Master) {
        parsed.show_master_shapes = parse_optional_bool_attribute(start, "showMasterSp")?;
        parsed
            .raw_attributes
            .retain(|(name, _)| name != "showMasterSp");
    }
    if matches!(kind, RootKind::Slide) {
        parsed.show = parse_optional_bool_attribute(start, "show")?;
        parsed.raw_attributes.retain(|(name, _)| name != "show");
    }
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let namespaces = root_namespaces.with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let namespace_uri = namespaces.element_uri(child.name().as_ref());
                let raw = capture_element(reader, &child)?;
                parsed.capture_child(&name, namespace_uri, &namespaces, raw, kind)?;
            }
            Event::Empty(child) => {
                let namespaces = root_namespaces.with_start(&child)?;
                let name = local_name(child.name().as_ref()).to_vec();
                let namespace_uri = namespaces.element_uri(child.name().as_ref());
                let raw = capture_empty_element(&child)?;
                parsed.capture_child(&name, namespace_uri, &namespaces, raw, kind)?;
            }
            Event::End(end) if local_name(end.name().as_ref()) == kind.local_name() => {
                return Ok(parsed);
            }
            Event::Eof => return Err(OxmlError::MissingElement(format!("closing {}", kind.tag()))),
            _ => {}
        }
        buffer.clear();
    }
}

impl ParsedRoot {
    fn capture_child(
        &mut self,
        name: &[u8],
        namespace_uri: Option<&str>,
        namespaces: &NamespaceBindings,
        raw: Vec<u8>,
        kind: RootKind,
    ) -> Result<()> {
        let is_p = namespace_uri == Some(P_NS);
        let is_mc = namespace_uri == Some(MC_NS);
        if is_p && name == b"cSld" {
            namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
            if self.common_slide_data.is_some() {
                return Err(duplicate("cSld"));
            }
            self.advance_modelled_child("cSld", 0, 1)?;
            self.common_slide_data = Some(CT_CommonSlideData::from_fragment(&raw, namespaces)?);
            return Ok(());
        }
        if is_p && name == b"clrMapOvr" && !matches!(kind, RootKind::Master) {
            namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
            if self.color_map_override.is_some() {
                return Err(duplicate("clrMapOvr"));
            }
            self.advance_modelled_child("clrMapOvr", 1, 2)?;
            self.color_map_override = Some(CT_ColorMapOverride::from_fragment(&raw, namespaces)?);
            return Ok(());
        }
        if is_p && name == b"clrMap" && matches!(kind, RootKind::Master) {
            namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
            if self.color_map.is_some() {
                return Err(duplicate("clrMap"));
            }
            self.advance_modelled_child("clrMap", 1, 2)?;
            self.color_map = Some(parse_color_map(&raw, namespaces)?);
            return Ok(());
        }
        if is_p && name == b"txStyles" && matches!(kind, RootKind::Master) {
            namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
            if self.text_styles.is_some() {
                return Err(duplicate("txStyles"));
            }
            self.advance_modelled_child("txStyles", 6, 7)?;
            self.text_styles = Some(Box::new(CT_MasterTextStyles::from_fragment(
                &raw, namespaces,
            )?));
            return Ok(());
        }
        if is_p && name == b"hf" && matches!(kind, RootKind::Layout | RootKind::Master) {
            namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
            if self.header_footer.is_some() {
                return Err(duplicate("hf"));
            }
            let (before, after) = match kind {
                RootKind::Layout => (4, 5),
                RootKind::Master => (5, 6),
                RootKind::Slide => unreachable!(),
            };
            self.advance_modelled_child("hf", before, after)?;
            self.header_footer = Some(CT_HeaderFooter::from_fragment(&raw, namespaces)?);
            return Ok(());
        }
        if is_p && name == b"transition" {
            let transition = CT_SlideTransition::from_fragment(&raw, namespaces)?;
            self.capture_transition(transition, kind)?;
            return Ok(());
        }
        if is_mc
            && name == b"AlternateContent"
            && let Some(transition) = CT_SlideTransition::from_alternate_content(&raw, namespaces)?
        {
            self.capture_transition(transition, kind)?;
            return Ok(());
        }
        if is_p && name == b"timing" {
            if self.timing.is_some() {
                return Err(duplicate("timing"));
            }
            let (before, after) = match kind {
                RootKind::Slide => (3, 4),
                RootKind::Layout => (3, 4),
                RootKind::Master => (4, 5),
            };
            if self.boundary > before {
                return Err(OxmlError::InvalidValue(
                    "p:timing is out of schema order".to_owned(),
                ));
            }
            self.timing = Some(CT_Timing::from_fragment(&raw, namespaces)?);
            self.boundary = after;
            return Ok(());
        }
        let (at, after) = if is_p {
            kind.raw_boundary(name, self.boundary)
        } else {
            (self.boundary, self.boundary)
        };
        self.raw_children.push(at, raw);
        self.boundary = after;
        Ok(())
    }

    fn capture_transition(&mut self, transition: CT_SlideTransition, kind: RootKind) -> Result<()> {
        if self.transition.is_some() {
            return Err(duplicate("transition"));
        }
        let (before, after) = match kind {
            RootKind::Slide => (2, 3),
            RootKind::Layout => (2, 3),
            RootKind::Master => (3, 4),
        };
        if self.boundary > before {
            return Err(OxmlError::InvalidValue(
                "p:transition is out of schema order".to_owned(),
            ));
        }
        self.transition = Some(transition);
        self.boundary = after;
        Ok(())
    }

    fn advance_modelled_child(&mut self, name: &str, before: usize, after: usize) -> Result<()> {
        if self.boundary > before {
            return Err(OxmlError::InvalidValue(format!(
                "p:{name} is out of schema order"
            )));
        }
        self.boundary = after;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn write_slide_like(
    kind: RootKind,
    common: &CT_CommonSlideData,
    color_override: Option<&CT_ColorMapOverride>,
    transition: Option<&CT_SlideTransition>,
    timing: Option<&CT_Timing>,
    show: Option<bool>,
    show_master_shapes: Option<bool>,
    header_footer: Option<&CT_HeaderFooter>,
    attributes: &RawAttributes,
    raw: &OrderedRawChildren,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    write_root_start(&mut writer, kind, show, show_master_shapes, attributes)?;
    emit_raw(&mut writer, raw.at(0))?;
    common.write_xml(&mut writer)?;
    emit_raw(&mut writer, raw.at(1))?;
    if let Some(color_override) = color_override {
        color_override.write_xml(&mut writer)?;
    }
    match kind {
        RootKind::Slide => {
            emit_raw(&mut writer, raw.at(2))?;
            if let Some(transition) = transition {
                transition.write_xml(&mut writer)?;
            }
            emit_raw(&mut writer, raw.at(3))?;
            if let Some(timing) = timing {
                timing.write_xml(&mut writer)?;
            }
            emit_raw(&mut writer, raw.at(4))?;
            emit_raw(&mut writer, raw.at(5))?;
        }
        RootKind::Layout => {
            emit_raw(&mut writer, raw.at(2))?;
            if let Some(transition) = transition {
                transition.write_xml(&mut writer)?;
            }
            emit_raw(&mut writer, raw.at(3))?;
            if let Some(timing) = timing {
                timing.write_xml(&mut writer)?;
            }
            emit_raw(&mut writer, raw.at(4))?;
            if let Some(header_footer) = header_footer {
                header_footer.write_xml(&mut writer)?;
            }
            emit_raw(&mut writer, raw.at(5))?;
            emit_raw(&mut writer, raw.at(6))?;
        }
        RootKind::Master => unreachable!(),
    }
    writer.write_event(Event::End(BytesEnd::new(kind.tag())))?;
    Ok(writer.into_inner())
}

fn write_root_start<W: Write>(
    writer: &mut Writer<W>,
    kind: RootKind,
    show: Option<bool>,
    show_master_shapes: Option<bool>,
    attributes: &RawAttributes,
) -> Result<()> {
    let mut root = BytesStart::new(kind.tag());
    root.push_attribute(("xmlns:p", P_NS));
    root.push_attribute(("xmlns:a", A_NS));
    root.push_attribute(("xmlns:r", R_NS));
    push_optional_bool_attribute(&mut root, "show", show);
    if let Some(show_master_shapes) = show_master_shapes {
        root.push_attribute(("showMasterSp", if show_master_shapes { "1" } else { "0" }));
    }
    push_attributes(&mut root, attributes);
    writer.write_event(Event::Start(root))?;
    Ok(())
}

impl CT_HeaderFooter {
    pub(crate) fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"hf"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    let mut value = Self::from_start(&start)?;
                    value.raw_children = capture_shell_children(&mut reader, b"hf")?;
                    return Ok(value);
                }
                Event::Empty(start) => return Self::from_start(&start),
                Event::Eof => return Err(OxmlError::MissingElement("p:hf".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        let mut raw_attributes = root_attributes(start, FIXED_MODEL_PREFIXES)?;
        raw_attributes
            .retain(|(name, _)| !matches!(name.as_str(), "sldNum" | "hdr" | "ftr" | "dt"));
        Ok(Self {
            slide_number: parse_optional_bool_attribute(start, "sldNum")?,
            header: parse_optional_bool_attribute(start, "hdr")?,
            footer: parse_optional_bool_attribute(start, "ftr")?,
            date_time: parse_optional_bool_attribute(start, "dt")?,
            raw_attributes,
            raw_children: OrderedRawChildren::default(),
        })
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:hf");
        push_optional_bool_attribute(&mut start, "sldNum", self.slide_number);
        push_optional_bool_attribute(&mut start, "hdr", self.header);
        push_optional_bool_attribute(&mut start, "ftr", self.footer);
        push_optional_bool_attribute(&mut start, "dt", self.date_time);
        push_attributes(&mut start, &self.raw_attributes);
        if self.raw_children.is_empty() {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;
            emit_raw(writer, self.raw_children.at(0))?;
            writer.write_event(Event::End(BytesEnd::new("p:hf")))?;
        }
        Ok(())
    }
}

fn parse_optional_bool_attribute(start: &BytesStart<'_>, name: &str) -> Result<Option<bool>> {
    let Some((_, value)) = all_attributes(start)?
        .into_iter()
        .find(|(attribute, _)| attribute == name)
    else {
        return Ok(None);
    };
    match value.as_str() {
        "true" | "1" => Ok(Some(true)),
        "false" | "0" => Ok(Some(false)),
        _ => Err(OxmlError::InvalidValue(format!(
            "invalid boolean @{name} value {value}"
        ))),
    }
}

fn push_optional_bool_attribute(
    start: &mut BytesStart<'_>,
    name: &'static str,
    value: Option<bool>,
) {
    if let Some(value) = value {
        start.push_attribute((name, if value { "1" } else { "0" }));
    }
}

impl CT_CommonSlideData {
    /// Sets a direct `p:bgPr` fill. An existing direct-fill background keeps
    /// its other children and only its fill changes, and any other
    /// background, such as a `p:bgRef` theme reference, is replaced whole.
    pub fn set_background_fill(&mut self, fill: Fill) -> Result<()> {
        if let Some(background) = &mut self.background
            && background.fill_range.is_some()
        {
            return background.replace_fill(fill);
        }
        self.background = Some(CT_Background::from_fill(fill)?);
        Ok(())
    }

    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = NamespaceBindings::default().with_start(&start)?;
                    if local_name(start.name().as_ref()) == b"cSld"
                        && namespaces.element_uri(start.name().as_ref()) == Some(P_NS)
                    {
                        return Self::from_reader(&mut reader, &start, &namespaces);
                    }
                    return Err(unexpected(&start));
                }
                Event::Empty(start) => {
                    let namespaces = NamespaceBindings::default().with_start(&start)?;
                    if local_name(start.name().as_ref()) != b"cSld"
                        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                    {
                        return Err(unexpected(&start));
                    }
                    return Err(OxmlError::MissingElement(
                        "p:cSld requires p:spTree".to_owned(),
                    ));
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:cSld".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    pub(crate) fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    return Self::from_reader(&mut reader, &start, &namespaces);
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:cSld".to_owned())),
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
        namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
        let attributes = all_attributes(start)?;
        let name = attributes
            .iter()
            .find(|(key, _)| key == "name")
            .map(|(_, value)| value.clone());
        let raw_attributes = attributes
            .into_iter()
            .filter(|(key, _)| key != "name" && !is_fixed_xmlns(key, FIXED_MODEL_PREFIXES))
            .collect();
        let mut background = None;
        let mut shape_tree = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0usize;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_ns = namespaces.with_start(&child)?;
                    let local = local_name(child.name().as_ref()).to_vec();
                    let is_p = child_ns.element_uri(child.name().as_ref()) == Some(P_NS);
                    let raw = capture_element(reader, &child)?;
                    capture_common_child(
                        &local,
                        is_p,
                        raw,
                        &mut background,
                        &mut shape_tree,
                        &child_ns,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::Empty(child) => {
                    let child_ns = namespaces.with_start(&child)?;
                    let local = local_name(child.name().as_ref()).to_vec();
                    let is_p = child_ns.element_uri(child.name().as_ref()) == Some(P_NS);
                    let raw = capture_empty_element(&child)?;
                    capture_common_child(
                        &local,
                        is_p,
                        raw,
                        &mut background,
                        &mut shape_tree,
                        &child_ns,
                        &mut raw_children,
                        &mut boundary,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"cSld" => break,
                Event::Eof => return Err(OxmlError::MissingElement("closing p:cSld".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
        Ok(Self {
            name,
            background,
            shape_tree: required(shape_tree, "p:spTree")?,
            raw_attributes,
            raw_children,
        })
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:cSld");
        if let Some(name) = &self.name {
            start.push_attribute(("name", name.as_str()));
        }
        push_attributes(&mut start, &self.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(background) = &self.background {
            writer.get_mut().write_all(background.raw_xml())?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        self.shape_tree.write_xml(writer)?;
        for boundary in 2..=5 {
            emit_raw(writer, self.raw_children.at(boundary))?;
        }
        writer.write_event(Event::End(BytesEnd::new("p:cSld")))?;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn capture_common_child(
    name: &[u8],
    is_p: bool,
    raw: Vec<u8>,
    background: &mut Option<CT_Background>,
    shape_tree: &mut Option<CT_ShapeTree>,
    namespaces: &NamespaceBindings,
    children: &mut OrderedRawChildren,
    boundary: &mut usize,
) -> Result<()> {
    if is_p && name == b"bg" {
        let parsed = CT_Background::from_fragment(&raw, namespaces)?;
        if background.replace(parsed).is_some() {
            return Err(duplicate("bg"));
        }
        *boundary = (*boundary).max(1);
    } else if is_p && name == b"spTree" {
        let parsed = CT_ShapeTree::from_fragment(&raw, &namespaces.entries())?;
        if shape_tree.replace(parsed).is_some() {
            return Err(duplicate("spTree"));
        }
        *boundary = (*boundary).max(2);
    } else {
        let before = if is_p {
            match name {
                b"custDataLst" => Some(2),
                b"controls" => Some(3),
                b"extLst" => Some(4),
                _ => None,
            }
        } else {
            None
        };
        if let Some(before) = before {
            let at = (*boundary).max(before);
            children.push(at, raw);
            *boundary = at.max(before + 1);
        } else {
            children.push(*boundary, raw);
        }
    }
    Ok(())
}

fn parse_background_rendering(
    reader: &mut Reader<&[u8]>,
    namespaces: &NamespaceBindings,
) -> Result<BackgroundRendering> {
    let mut rendering = None;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let is_p = child_ns.element_uri(child.name().as_ref()) == Some(P_NS);
                let raw = capture_element(reader, &child)?;
                let parsed = match (is_p, local_name(child.name().as_ref())) {
                    (true, b"bgPr") => Some(parse_background_properties(&raw, &child_ns)?),
                    (true, b"bgRef") => Some(parse_background_reference(&raw, &child_ns)?),
                    _ => None,
                };
                if let Some(parsed) = parsed
                    && rendering.replace(parsed).is_some()
                {
                    return Err(duplicate("background choice"));
                }
            }
            Event::Empty(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let is_p = child_ns.element_uri(child.name().as_ref()) == Some(P_NS);
                let raw = capture_empty_element(&child)?;
                let parsed = match (is_p, local_name(child.name().as_ref())) {
                    (true, b"bgPr") => Some(BackgroundRendering::Properties(None)),
                    (true, b"bgRef") => Some(parse_background_reference(&raw, &child_ns)?),
                    _ => None,
                };
                if let Some(parsed) = parsed
                    && rendering.replace(parsed).is_some()
                {
                    return Err(duplicate("background choice"));
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"bg" => break,
            Event::Eof => return Err(OxmlError::MissingElement("closing p:bg".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
    required(rendering, "p:bgPr or p:bgRef")
}

fn parse_background_properties(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<BackgroundRendering> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                if local_name(start.name().as_ref()) != b"bgPr"
                    || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                {
                    return Err(unexpected(&start));
                }
                return parse_background_fill_children(&mut reader, &namespaces);
            }
            Event::Empty(_) => return Ok(BackgroundRendering::Properties(None)),
            Event::Eof => return Err(OxmlError::MissingElement("p:bgPr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_background_fill_children(
    reader: &mut Reader<&[u8]>,
    namespaces: &NamespaceBindings,
) -> Result<BackgroundRendering> {
    let mut fill = None;
    let mut unsupported = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let is_drawing = child_ns.element_uri(child.name().as_ref()) == Some(A_NS);
                let qualified_name = child.name();
                let name = local_name(qualified_name.as_ref());
                let is_fill = is_drawing && is_fill_name(name);
                let raw = capture_element(reader, &child)?;
                if is_fill {
                    let parsed = Fill::from_xml(&raw).map_err(map_drawing_error)?;
                    if fill.replace(parsed).is_some() {
                        return Err(duplicate("background fill"));
                    }
                } else if is_drawing && name == b"grpFill" {
                    unsupported = true;
                }
            }
            Event::Empty(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let is_drawing = child_ns.element_uri(child.name().as_ref()) == Some(A_NS);
                let qualified_name = child.name();
                let name = local_name(qualified_name.as_ref());
                if is_drawing && is_fill_name(name) {
                    let raw = capture_empty_element(&child)?;
                    let parsed = Fill::from_xml(&raw).map_err(map_drawing_error)?;
                    if fill.replace(parsed).is_some() {
                        return Err(duplicate("background fill"));
                    }
                } else if is_drawing && name == b"grpFill" {
                    unsupported = true;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"bgPr" => break,
            Event::Eof => return Err(OxmlError::MissingElement("closing p:bgPr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
    if unsupported {
        Ok(BackgroundRendering::Unsupported("group fill"))
    } else {
        Ok(BackgroundRendering::Properties(fill.map(Box::new)))
    }
}

fn parse_background_reference(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<BackgroundRendering> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let namespaces = inherited.with_start(&start)?;
                if local_name(start.name().as_ref()) != b"bgRef"
                    || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
                {
                    return Err(unexpected(&start));
                }
                let index = required_attribute(&start, "idx")?.parse::<u32>()?;
                let (color, unsupported) =
                    parse_background_reference_color(&mut reader, &namespaces)?;
                return if unsupported {
                    Ok(BackgroundRendering::Unsupported("reference colour"))
                } else {
                    Ok(BackgroundRendering::Reference { index, color })
                };
            }
            Event::Empty(start) => {
                let index = required_attribute(&start, "idx")?.parse::<u32>()?;
                return Ok(BackgroundRendering::Reference { index, color: None });
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:bgRef".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_background_reference_color(
    reader: &mut Reader<&[u8]>,
    namespaces: &NamespaceBindings,
) -> Result<(Option<ColorChoice>, bool)> {
    let mut color = None;
    let mut unsupported = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let is_drawing = child_ns.element_uri(child.name().as_ref()) == Some(A_NS);
                let qualified_name = child.name();
                let name = local_name(qualified_name.as_ref());
                if is_drawing && is_color_name(name) {
                    let parsed =
                        ColorChoice::from_xml(reader, &child).map_err(map_drawing_error)?;
                    if color.replace(parsed).is_some() {
                        return Err(duplicate("background reference colour"));
                    }
                } else {
                    capture_element(reader, &child)?;
                    unsupported |= is_drawing && is_unmodelled_color_name(name);
                }
            }
            Event::Empty(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let is_drawing = child_ns.element_uri(child.name().as_ref()) == Some(A_NS);
                let qualified_name = child.name();
                let name = local_name(qualified_name.as_ref());
                if is_drawing && is_color_name(name) {
                    let parsed = ColorChoice::from_empty_xml(&child).map_err(map_drawing_error)?;
                    if color.replace(parsed).is_some() {
                        return Err(duplicate("background reference colour"));
                    }
                } else {
                    unsupported |= is_drawing && is_unmodelled_color_name(name);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"bgRef" => break,
            Event::Eof => return Err(OxmlError::MissingElement("closing p:bgRef".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
    Ok((color, unsupported))
}

fn is_fill_name(name: &[u8]) -> bool {
    matches!(
        name,
        b"noFill" | b"solidFill" | b"gradFill" | b"pattFill" | b"blipFill"
    )
}

fn is_color_name(name: &[u8]) -> bool {
    matches!(name, b"srgbClr" | b"schemeClr" | b"sysClr" | b"prstClr")
}

fn is_unmodelled_color_name(name: &[u8]) -> bool {
    matches!(name, b"scrgbClr" | b"hslClr")
}

fn required_attribute(start: &BytesStart<'_>, name: &str) -> Result<String> {
    all_attributes(start)?
        .into_iter()
        .find_map(|(key, value)| (key == name).then_some(value))
        .ok_or_else(|| {
            OxmlError::MissingElement(format!(
                "{} requires @{name}",
                String::from_utf8_lossy(start.name().as_ref())
            ))
        })
}

impl CT_ColorMapOverride {
    pub(crate) fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => break start.into_owned(),
                Event::Eof => return Err(OxmlError::MissingElement("p:clrMapOvr".to_owned())),
                _ => {}
            }
            buffer.clear();
        };
        let namespaces = inherited.with_start(&start)?;
        namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
        let raw_attributes = root_attributes(&start, FIXED_MODEL_PREFIXES)?;
        let mut kind = None;
        let mut mapping_attributes = Vec::new();
        let mut mapping_children = OrderedRawChildren::default();
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0usize;
        loop {
            buffer.clear();
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_ns = namespaces.with_start(&child)?;
                    let local = local_name(child.name().as_ref()).to_vec();
                    let is_a = child_ns.element_uri(child.name().as_ref()) == Some(A_NS);
                    let raw = capture_element(&mut reader, &child)?;
                    capture_override_choice(
                        &local,
                        is_a,
                        raw,
                        &mut kind,
                        &mut mapping_attributes,
                        &mut mapping_children,
                        &mut raw_children,
                        &mut boundary,
                        &child_ns,
                    )?;
                }
                Event::Empty(child) => {
                    let child_ns = namespaces.with_start(&child)?;
                    let local = local_name(child.name().as_ref()).to_vec();
                    let is_a = child_ns.element_uri(child.name().as_ref()) == Some(A_NS);
                    let raw = capture_empty_element(&child)?;
                    capture_override_choice(
                        &local,
                        is_a,
                        raw,
                        &mut kind,
                        &mut mapping_attributes,
                        &mut mapping_children,
                        &mut raw_children,
                        &mut boundary,
                        &child_ns,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"clrMapOvr" => break,
                Event::Eof => {
                    return Err(OxmlError::MissingElement("closing p:clrMapOvr".to_owned()));
                }
                _ => {}
            }
        }
        Ok(Self {
            kind: required(kind, "a:masterClrMapping or a:overrideClrMapping")?,
            raw_attributes,
            mapping_attributes,
            mapping_children,
            raw_children,
        })
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:clrMapOvr");
        push_attributes(&mut start, &self.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        match &self.kind {
            ColorMapOverrideKind::Master => {
                write_mapping_shell(
                    writer,
                    "a:masterClrMapping",
                    &self.mapping_attributes,
                    &self.mapping_children,
                )?;
            }
            ColorMapOverrideKind::Override(map) => {
                write_color_map(
                    writer,
                    "a:overrideClrMapping",
                    map,
                    &self.mapping_attributes,
                    &self.mapping_children,
                )?;
            }
        }
        emit_raw(writer, self.raw_children.at(1))?;
        writer.write_event(Event::End(BytesEnd::new("p:clrMapOvr")))?;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn capture_override_choice(
    name: &[u8],
    is_a: bool,
    raw: Vec<u8>,
    kind: &mut Option<ColorMapOverrideKind>,
    attributes: &mut RawAttributes,
    mapping_children: &mut OrderedRawChildren,
    raw_children: &mut OrderedRawChildren,
    boundary: &mut usize,
    namespaces: &NamespaceBindings,
) -> Result<()> {
    if is_a && matches!(name, b"masterClrMapping" | b"overrideClrMapping") {
        namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
        if kind.is_some() {
            return Err(duplicate("colour-map override choice"));
        }
        if name == b"overrideClrMapping" {
            let parsed = parse_color_map(&raw, namespaces)?;
            *attributes = parsed.raw_attributes;
            *mapping_children = parsed.raw_children;
            *kind = Some(ColorMapOverrideKind::Override(parsed.value));
        } else {
            let (attrs, children) = parse_raw_shell(&raw)?;
            *attributes = attrs;
            *mapping_children = children;
            *kind = Some(ColorMapOverrideKind::Master);
        }
        *boundary = 1;
    } else {
        raw_children.push(*boundary, raw);
    }
    Ok(())
}

impl CT_MasterTextStyles {
    fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => break start.into_owned(),
                Event::Eof => return Err(OxmlError::MissingElement("p:txStyles".to_owned())),
                _ => {}
            }
            buffer.clear();
        };
        let namespaces = inherited.with_start(&start)?;
        namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
        let raw_attributes = root_attributes(&start, FIXED_MODEL_PREFIXES)?;
        let mut styles: [Option<CT_TextListStyle>; 3] = [None, None, None];
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0usize;
        loop {
            buffer.clear();
            match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let child_ns = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let is_p = child_ns.element_uri(child.name().as_ref()) == Some(P_NS);
                    let raw = capture_element(&mut reader, &child)?;
                    capture_text_style(
                        &name,
                        is_p,
                        raw,
                        &mut styles,
                        &mut raw_children,
                        &mut boundary,
                        &child_ns,
                    )?;
                }
                Event::Empty(child) => {
                    let child_ns = namespaces.with_start(&child)?;
                    let name = local_name(child.name().as_ref()).to_vec();
                    let is_p = child_ns.element_uri(child.name().as_ref()) == Some(P_NS);
                    let raw = capture_empty_element(&child)?;
                    capture_text_style(
                        &name,
                        is_p,
                        raw,
                        &mut styles,
                        &mut raw_children,
                        &mut boundary,
                        &child_ns,
                    )?;
                }
                Event::End(end) if local_name(end.name().as_ref()) == b"txStyles" => break,
                Event::Eof => {
                    return Err(OxmlError::MissingElement("closing p:txStyles".to_owned()));
                }
                _ => {}
            }
        }
        let [title, body, other] = styles;
        Ok(Self {
            title_style: required(title, "p:titleStyle")?,
            body_style: required(body, "p:bodyStyle")?,
            other_style: required(other, "p:otherStyle")?,
            raw_attributes,
            raw_children,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("p:txStyles");
        push_attributes(&mut start, &self.raw_attributes);
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, self.raw_children.at(0))?;
        self.title_style
            .write_xml_as(writer, "p:titleStyle")
            .map_err(map_drawing_error)?;
        emit_raw(writer, self.raw_children.at(1))?;
        self.body_style
            .write_xml_as(writer, "p:bodyStyle")
            .map_err(map_drawing_error)?;
        emit_raw(writer, self.raw_children.at(2))?;
        self.other_style
            .write_xml_as(writer, "p:otherStyle")
            .map_err(map_drawing_error)?;
        emit_raw(writer, self.raw_children.at(3))?;
        writer.write_event(Event::End(BytesEnd::new("p:txStyles")))?;
        Ok(())
    }
}

fn capture_text_style(
    name: &[u8],
    is_p: bool,
    raw: Vec<u8>,
    styles: &mut [Option<CT_TextListStyle>; 3],
    children: &mut OrderedRawChildren,
    boundary: &mut usize,
    namespaces: &NamespaceBindings,
) -> Result<()> {
    let index = if is_p {
        match name {
            b"titleStyle" => Some(0),
            b"bodyStyle" => Some(1),
            b"otherStyle" => Some(2),
            _ => None,
        }
    } else {
        None
    };
    if let Some(index) = index {
        namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
        if styles[index].is_some() {
            return Err(duplicate(std::str::from_utf8(name).unwrap_or("text style")));
        }
        styles[index] = Some(CT_TextListStyle::from_xml(&raw).map_err(map_drawing_error)?);
        *boundary = (*boundary).max(index + 1);
    } else {
        children.push(*boundary, raw);
    }
    Ok(())
}

pub(crate) fn parse_color_map(
    xml: &[u8],
    namespaces: &NamespaceBindings,
) -> Result<ParsedColorMap> {
    namespaces.reject_writer_conflicts(FIXED_MODEL_PREFIXES)?;
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let value = color_map_from_start(&start)?;
                let raw_attributes = color_map_raw_attributes(&start)?;
                let raw_children =
                    capture_shell_children(&mut reader, local_name(start.name().as_ref()))?;
                return Ok(ParsedColorMap {
                    value,
                    raw_attributes,
                    raw_children,
                });
            }
            Event::Empty(start) => {
                return Ok(ParsedColorMap {
                    value: color_map_from_start(&start)?,
                    raw_attributes: color_map_raw_attributes(&start)?,
                    raw_children: OrderedRawChildren::default(),
                });
            }
            Event::Eof => return Err(OxmlError::MissingElement("colour map".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn color_map_from_start(start: &BytesStart<'_>) -> Result<ColorMap> {
    let attrs = all_attributes(start)?;
    let value = |name: &str| -> Result<ThemeColorSlot> {
        let raw = attrs
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
            .ok_or_else(|| OxmlError::MissingElement(format!("colour map requires @{name}")))?;
        parse_theme_slot(raw)
            .ok_or_else(|| OxmlError::InvalidValue(format!("invalid colour-map slot {raw}")))
    };
    Ok(ColorMap::new(
        value("bg1")?,
        value("tx1")?,
        value("bg2")?,
        value("tx2")?,
        value("accent1")?,
        value("accent2")?,
        value("accent3")?,
        value("accent4")?,
        value("accent5")?,
        value("accent6")?,
        value("hlink")?,
        value("folHlink")?,
    ))
}

fn parse_theme_slot(value: &str) -> Option<ThemeColorSlot> {
    Some(match value {
        "dk1" => ThemeColorSlot::Dark1,
        "lt1" => ThemeColorSlot::Light1,
        "dk2" => ThemeColorSlot::Dark2,
        "lt2" => ThemeColorSlot::Light2,
        "accent1" => ThemeColorSlot::Accent1,
        "accent2" => ThemeColorSlot::Accent2,
        "accent3" => ThemeColorSlot::Accent3,
        "accent4" => ThemeColorSlot::Accent4,
        "accent5" => ThemeColorSlot::Accent5,
        "accent6" => ThemeColorSlot::Accent6,
        "hlink" => ThemeColorSlot::Hyperlink,
        "folHlink" => ThemeColorSlot::FollowedHyperlink,
        _ => return None,
    })
}

fn color_map_raw_attributes(start: &BytesStart<'_>) -> Result<RawAttributes> {
    const KNOWN: [&str; 12] = [
        "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    Ok(all_attributes(start)?
        .into_iter()
        .filter(|(name, _)| {
            !KNOWN.contains(&name.as_str()) && !is_fixed_xmlns(name, FIXED_MODEL_PREFIXES)
        })
        .collect())
}

pub(crate) fn write_color_map<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    map: &ColorMap,
    attributes: &RawAttributes,
    children: &OrderedRawChildren,
) -> Result<()> {
    let mut start = BytesStart::new(tag);
    for (name, slot) in color_map_slots() {
        start.push_attribute((name, map.theme_slot(slot).as_str()));
    }
    push_attributes(&mut start, attributes);
    if children.is_empty() {
        writer.write_event(Event::Empty(start))?;
    } else {
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, children.at(0))?;
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
    }
    Ok(())
}

fn color_map_slots() -> [(&'static str, ColorMapSlot); 12] {
    [
        ("bg1", ColorMapSlot::Background1),
        ("tx1", ColorMapSlot::Text1),
        ("bg2", ColorMapSlot::Background2),
        ("tx2", ColorMapSlot::Text2),
        ("accent1", ColorMapSlot::Accent1),
        ("accent2", ColorMapSlot::Accent2),
        ("accent3", ColorMapSlot::Accent3),
        ("accent4", ColorMapSlot::Accent4),
        ("accent5", ColorMapSlot::Accent5),
        ("accent6", ColorMapSlot::Accent6),
        ("hlink", ColorMapSlot::Hyperlink),
        ("folHlink", ColorMapSlot::FollowedHyperlink),
    ]
}

fn parse_raw_shell(xml: &[u8]) -> Result<(RawAttributes, OrderedRawChildren)> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let attrs = root_attributes(&start, FIXED_MODEL_PREFIXES)?;
                let children =
                    capture_shell_children(&mut reader, local_name(start.name().as_ref()))?;
                return Ok((attrs, children));
            }
            Event::Empty(start) => {
                return Ok((
                    root_attributes(&start, FIXED_MODEL_PREFIXES)?,
                    OrderedRawChildren::default(),
                ));
            }
            Event::Eof => return Err(OxmlError::MissingElement("mapping element".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn capture_shell_children(
    reader: &mut Reader<&[u8]>,
    root_name: &[u8],
) -> Result<OrderedRawChildren> {
    let mut children = OrderedRawChildren::default();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => children.push(0, capture_element(reader, &child)?),
            Event::Empty(child) => children.push(0, capture_empty_element(&child)?),
            event @ (Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::DocType(_)) => children.push(0, capture_event(event)?),
            Event::End(end) if local_name(end.name().as_ref()) == root_name => return Ok(children),
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "closing mapping element".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn write_mapping_shell<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    attributes: &RawAttributes,
    children: &OrderedRawChildren,
) -> Result<()> {
    let mut start = BytesStart::new(tag);
    push_attributes(&mut start, attributes);
    if children.is_empty() {
        writer.write_event(Event::Empty(start))?;
    } else {
        writer.write_event(Event::Start(start))?;
        emit_raw(writer, children.at(0))?;
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
    }
    Ok(())
}

fn push_attributes(start: &mut BytesStart<'_>, attributes: &RawAttributes) {
    for (name, value) in attributes {
        start.push_attribute((name.as_str(), value.as_str()));
    }
}

fn capture_event(event: Event<'_>) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event.into_owned())?;
    Ok(writer.into_inner())
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

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn required<T>(value: Option<T>, name: &str) -> Result<T> {
    value.ok_or_else(|| OxmlError::MissingElement(name.to_owned()))
}

fn duplicate(name: &str) -> OxmlError {
    OxmlError::InvalidValue(format!("duplicate PresentationML {name}"))
}

fn unexpected(element: &BytesStart<'_>) -> OxmlError {
    OxmlError::UnexpectedElement(String::from_utf8_lossy(element.name().as_ref()).into_owned())
}

fn map_drawing_error(error: impl ToString) -> OxmlError {
    OxmlError::InvalidValue(error.to_string())
}

#[cfg(test)]
mod stack_tests {
    use super::ParsedRoot;

    #[test]
    fn slide_root_parse_state_keeps_large_master_text_styles_off_stack() {
        assert!(
            std::mem::size_of::<ParsedRoot>() < 4 * 1024,
            "slide root parse state must remain small enough for nested producer content"
        );
    }
}
