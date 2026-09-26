//! Typed projections over PresentationML timing and slide-transition XML.
//!
//! Supported values are exposed for timeline evaluation. The original subtree
//! remains the serialization source so unsupported siblings, attributes, and
//! relationship-bearing extensions survive byte for byte.

use std::collections::HashMap;
use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::xml_text::resolve_entity;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::namespace::{MC_NS, NamespaceBindings, P_NS, all_attributes};

pub type Result<T> = std::result::Result<T, OxmlError>;

const P159_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
const P14_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

struct SelectedFragment {
    xml: Vec<u8>,
    parent_namespaces: NamespaceBindings,
    range: std::ops::Range<usize>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingDuration {
    Finite(u64),
    Indefinite,
}

impl TimingDuration {
    fn parse(value: Option<&str>) -> Result<Self> {
        match value {
            None | Some("indefinite") => Ok(Self::Indefinite),
            Some(value) => value
                .parse::<u64>()
                .map(Self::Finite)
                .map_err(|_| OxmlError::InvalidValue(format!("invalid timing duration {value}"))),
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingFill {
    Hold,
    Remove,
    Freeze,
    Transition,
    Other(String),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingRestart {
    Always,
    WhenNotActive,
    Never,
    Other(String),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingNodeType {
    TimingRoot,
    MainSequence,
    ClickEffect,
    WithEffect,
    AfterEffect,
    Other(String),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingEvent {
    OnBegin,
    OnEnd,
    OnClick,
    OnNext,
    OnPrevious,
    OnStopAudio,
    Other(String),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingTarget {
    Shape(u32),
    Slide,
    TimeNode(u32),
    Unsupported,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingCondition {
    pub event: Option<TimingEvent>,
    pub delay: TimingDuration,
    pub target: TimingTarget,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CommonTimeNode {
    pub id: u32,
    pub group_id: Option<u32>,
    pub duration: TimingDuration,
    pub fill: Option<TimingFill>,
    pub restart: Option<TimingRestart>,
    pub node_type: Option<TimingNodeType>,
    pub preset_id: Option<u32>,
    pub preset_class: Option<String>,
    pub preset_subtype: Option<u32>,
    pub start_conditions: Vec<TimingCondition>,
    pub end_conditions: Vec<TimingCondition>,
    pub children: Vec<TimingNode>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingContainer {
    pub common: CommonTimeNode,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingSequence {
    pub common: CommonTimeNode,
    pub concurrent: Option<bool>,
    pub next_action: Option<String>,
    pub previous_conditions: Vec<TimingCondition>,
    pub next_conditions: Vec<TimingCondition>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingSet {
    pub common: CommonTimeNode,
    pub target: TimingTarget,
    pub attribute_name: Option<String>,
    pub value: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingAnimate {
    pub common: CommonTimeNode,
    pub target: TimingTarget,
    pub attribute_name: Option<String>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub by: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingEffect {
    pub common: CommonTimeNode,
    pub target: TimingTarget,
    pub transition: Option<String>,
    pub filter: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingMotionPath {
    pub common: CommonTimeNode,
    pub target: TimingTarget,
    pub path: Option<String>,
    pub origin: Option<String>,
}

/// Whether retained media stays visible after playback stops.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MediaDisplayPolicy {
    HideWhenStopped,
    ShowWhenStopped,
}

/// The supported authored playback trigger subset.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MediaPlaybackTrigger {
    Automatic,
    OnClick,
    InClickSequence,
}

/// A supported media command carried by `p:cmd`.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum MediaCommandKind {
    Play { from_ms: Option<u64> },
    Pause,
    Stop,
    Seek { position_ms: u64 },
    Other(String),
}

/// The typed common-media-node projection shared by audio and video.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CommonMediaNode {
    pub common: CommonTimeNode,
    pub target: TimingTarget,
    pub volume: Option<u32>,
    pub looped: bool,
    pub display: Option<MediaDisplayPolicy>,
}

/// A typed `p:audio` or `p:video` timing node.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingMedia {
    pub common: CommonMediaNode,
    pub full_screen: Option<bool>,
}

/// A typed `p:cmd` targeting a media shape.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingMediaCommand {
    pub common: CommonTimeNode,
    pub target: TimingTarget,
    pub command: MediaCommandKind,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingUnsupported {
    pub local_name: String,
    raw_xml: Vec<u8>,
}

impl TimingUnsupported {
    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TimingNode {
    Parallel(TimingContainer),
    Sequence(TimingSequence),
    Set(TimingSet),
    Animate(TimingAnimate),
    Effect(TimingEffect),
    Motion(TimingMotionPath),
    Audio(TimingMedia),
    Video(TimingMedia),
    MediaCommand(TimingMediaCommand),
    Unsupported(TimingUnsupported),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TimingBuild {
    pub kind: String,
    pub shape_id: Option<u32>,
    pub group_id: Option<u32>,
    pub build_mode: Option<String>,
    pub build_level: Option<u32>,
    pub reverse: Option<bool>,
    pub advance_automatically: Option<TimingDuration>,
    pub animate_background: Option<bool>,
    pub auto_update_animated_background: Option<bool>,
    pub ui_expand: Option<bool>,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug)]
pub struct CT_Timing {
    nodes: Vec<TimingNode>,
    builds: Vec<TimingBuild>,
    condition_target_presence: HashMap<(u32, bool, usize), bool>,
    raw_xml: Vec<u8>,
    inherited_namespaces: Vec<(String, String)>,
}

impl PartialEq for CT_Timing {
    fn eq(&self, other: &Self) -> bool {
        self.nodes == other.nodes && self.builds == other.builds && self.raw_xml == other.raw_xml
    }
}

impl Eq for CT_Timing {}

impl CT_Timing {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_fragment(xml, &NamespaceBindings::default())
    }

    pub(crate) fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    validate_element(&start, &namespaces, b"timing")?;
                    let (nodes, builds) = parse_timing_children(&mut reader, &namespaces)?;
                    let condition_target_presence =
                        collect_condition_target_presence(xml, inherited)?;
                    return Ok(Self {
                        nodes,
                        builds,
                        condition_target_presence,
                        raw_xml: xml.to_vec(),
                        inherited_namespaces: inherited.entries(),
                    });
                }
                Event::Empty(start) => {
                    let namespaces = inherited.with_start(&start)?;
                    validate_element(&start, &namespaces, b"timing")?;
                    return Ok(Self {
                        nodes: Vec::new(),
                        builds: Vec::new(),
                        condition_target_presence: HashMap::new(),
                        raw_xml: xml.to_vec(),
                        inherited_namespaces: inherited.entries(),
                    });
                }
                Event::Eof => return Err(OxmlError::MissingElement("p:timing".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    pub fn to_xml(&self) -> Vec<u8> {
        self.raw_xml.clone()
    }

    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }

    pub fn nodes(&self) -> &[TimingNode] {
        &self.nodes
    }

    pub fn builds(&self) -> &[TimingBuild] {
        &self.builds
    }

    /// Returns whether any `spid` attribute in the retained subtree names `shape_id`.
    ///
    /// This covers animation targets and build lists, typed or not. A subtree
    /// that no longer reads cleanly counts as referencing every shape.
    pub fn references_shape(&self, shape_id: u32) -> bool {
        let mut reader = Reader::from_reader(self.raw_xml.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start) | Event::Empty(start)) => {
                    if all_attributes(&start).is_ok_and(|attributes| {
                        attributes.iter().any(|(name, value)| {
                            name.rsplit(':').next() == Some("spid")
                                && value.trim().parse::<u32>() == Ok(shape_id)
                        })
                    }) {
                        return true;
                    }
                }
                Ok(Event::Eof) => return false,
                Err(_) => return true,
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Returns the first typed audio or video node targeting `shape_id`.
    pub fn media_for_shape(&self, shape_id: u32) -> Option<&TimingMedia> {
        find_media_in_nodes(&self.nodes, shape_id)
    }

    /// Returns typed media commands targeting `shape_id` in timing order.
    pub fn media_commands_for_shape(&self, shape_id: u32) -> Vec<&TimingMediaCommand> {
        let mut commands = Vec::new();
        collect_media_commands(&self.nodes, shape_id, &mut commands);
        commands
    }

    /// Classifies the first media command using its complete timing ancestry.
    pub fn media_playback_trigger_for_shape(&self, shape_id: u32) -> Option<MediaPlaybackTrigger> {
        find_media_trigger(&self.nodes, shape_id, TriggerContext::default())
    }

    /// Appends one retained media node and its initial play command.
    pub fn add_media(
        &mut self,
        video: bool,
        shape_id: u32,
        volume: u32,
        looped: bool,
        show_when_stopped: bool,
        trigger: MediaPlaybackTrigger,
    ) -> Result<()> {
        validate_media_volume(volume)?;
        let inherited = NamespaceBindings::from_entries(&self.inherited_namespaces);
        let ids = next_timing_ids(&self.raw_xml, &inherited, 2)?;
        let media_xml =
            authored_media_xml(video, shape_id, ids[0], volume, looped, show_when_stopped);
        let command_xml = authored_media_command_xml(shape_id, ids[1], trigger);
        let inserted = format!("{media_xml}{command_xml}");
        let raw = append_to_timing_list(&self.raw_xml, &inherited, inserted.as_bytes())?;
        *self = Self::from_fragment(&raw, &inherited)?;
        Ok(())
    }

    /// Removes typed audio, video, and command nodes owned by `shape_id`.
    pub fn remove_media(&mut self, shape_id: u32) -> Result<()> {
        let inherited = NamespaceBindings::from_entries(&self.inherited_namespaces);
        let mut ranges = Vec::new();
        collect_owned_media_ranges(&self.raw_xml, &inherited, 0, shape_id, &mut ranges)?;
        ranges.sort_unstable_by_key(|range| range.start);
        let mut raw = self.raw_xml.clone();
        for range in ranges.into_iter().rev() {
            raw.drain(range);
        }
        *self = Self::from_fragment(&raw, &inherited)?;
        Ok(())
    }

    /// Returns whether one parsed common-node condition contains a target element.
    pub fn condition_has_explicit_target(
        &self,
        node_id: u32,
        end_condition: bool,
        index: usize,
    ) -> Option<bool> {
        self.condition_target_presence
            .get(&(node_id, end_condition, index))
            .copied()
    }

    /// Changes one common time-node duration without rewriting unrelated XML.
    pub fn set_node_duration(&mut self, id: u32, duration: TimingDuration) -> Result<()> {
        let inherited = NamespaceBindings::from_entries(&self.inherited_namespaces);
        let value = match duration {
            TimingDuration::Finite(value) => value.to_string(),
            TimingDuration::Indefinite => "indefinite".to_owned(),
        };
        let raw = rewrite_modeled_node_duration(&self.raw_xml, &inherited, id, &value)?;
        let replacement = Self::from_fragment(&raw, &inherited)?;
        *self = replacement;
        Ok(())
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.get_mut().write_all(&self.raw_xml)?;
        Ok(())
    }
}

#[derive(Clone, Copy, Default)]
struct TriggerContext {
    explicit_click: bool,
    click_sequence: bool,
}

fn find_media_trigger(
    nodes: &[TimingNode],
    shape_id: u32,
    inherited: TriggerContext,
) -> Option<MediaPlaybackTrigger> {
    for node in nodes {
        let mut context = inherited;
        if let Some(common) = timing_node_common(node) {
            context.explicit_click |= common.start_conditions.iter().any(|condition| {
                matches!(
                    condition.event,
                    Some(TimingEvent::OnClick | TimingEvent::OnNext)
                )
            });
            context.click_sequence |= matches!(common.node_type, Some(TimingNodeType::ClickEffect));
        }
        if let TimingNode::MediaCommand(command) = node
            && command.target == TimingTarget::Shape(shape_id)
        {
            return Some(if context.explicit_click {
                MediaPlaybackTrigger::OnClick
            } else if context.click_sequence {
                MediaPlaybackTrigger::InClickSequence
            } else {
                MediaPlaybackTrigger::Automatic
            });
        }
        if let Some(common) = timing_node_common(node)
            && let Some(trigger) = find_media_trigger(&common.children, shape_id, context)
        {
            return Some(trigger);
        }
    }
    None
}

fn find_media_in_nodes(nodes: &[TimingNode], shape_id: u32) -> Option<&TimingMedia> {
    for node in nodes {
        match node {
            TimingNode::Audio(media) | TimingNode::Video(media)
                if media.common.target == TimingTarget::Shape(shape_id) =>
            {
                return Some(media);
            }
            _ => {}
        }
        if let Some(common) = timing_node_common(node)
            && let Some(media) = find_media_in_nodes(&common.children, shape_id)
        {
            return Some(media);
        }
    }
    None
}

fn collect_media_commands<'a>(
    nodes: &'a [TimingNode],
    shape_id: u32,
    commands: &mut Vec<&'a TimingMediaCommand>,
) {
    for node in nodes {
        if let TimingNode::MediaCommand(command) = node
            && command.target == TimingTarget::Shape(shape_id)
        {
            commands.push(command);
        }
        if let Some(common) = timing_node_common(node) {
            collect_media_commands(&common.children, shape_id, commands);
        }
    }
}

fn timing_node_common(node: &TimingNode) -> Option<&CommonTimeNode> {
    match node {
        TimingNode::Parallel(node) => Some(&node.common),
        TimingNode::Sequence(node) => Some(&node.common),
        TimingNode::Set(node) => Some(&node.common),
        TimingNode::Animate(node) => Some(&node.common),
        TimingNode::Effect(node) => Some(&node.common),
        TimingNode::Motion(node) => Some(&node.common),
        TimingNode::Audio(node) | TimingNode::Video(node) => Some(&node.common.common),
        TimingNode::MediaCommand(node) => Some(&node.common),
        TimingNode::Unsupported(_) => None,
    }
}

fn validate_media_volume(volume: u32) -> Result<()> {
    if volume > 100_000 {
        return Err(OxmlError::InvalidValue(
            "media volume must be between 0 and 100000".to_owned(),
        ));
    }
    Ok(())
}

fn next_timing_ids(xml: &[u8], inherited: &NamespaceBindings, count: usize) -> Result<Vec<u32>> {
    let mut maximum = 0u32;
    collect_schema_timing_ids(xml, inherited, &mut maximum)?;
    (1..=count)
        .map(|offset| {
            maximum
                .checked_add(offset as u32)
                .ok_or_else(|| OxmlError::InvalidValue("timing ids are exhausted".to_owned()))
        })
        .collect()
}

fn collect_schema_timing_ids(
    xml: &[u8],
    inherited: &NamespaceBindings,
    maximum: &mut u32,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing id root");
                let scope = parent.with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"tnLst"
                {
                    collect_schema_node_list_ids(&raw, parent, maximum)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_schema_node_list_ids(
    xml: &[u8],
    inherited: &NamespaceBindings,
    maximum: &mut u32,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing node list id root");
                let scope = parent.with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && is_schema_timing_node(&local)
                {
                    collect_schema_node_ids(&raw, parent, &local, maximum)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn is_schema_timing_node(local: &[u8]) -> bool {
    matches!(
        local,
        b"par"
            | b"seq"
            | b"excl"
            | b"anim"
            | b"animClr"
            | b"animEffect"
            | b"animMotion"
            | b"animRot"
            | b"animScale"
            | b"cmd"
            | b"set"
            | b"audio"
            | b"video"
    )
}

fn collect_schema_node_ids(
    xml: &[u8],
    inherited: &NamespaceBindings,
    node_local: &[u8],
    maximum: &mut u32,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing node id root");
                let scope = parent.with_start(&child)?;
                let child_name = child.name();
                let local = local_name(child_name.as_ref());
                let direct_common =
                    matches!(node_local, b"par" | b"seq" | b"excl") && local == b"cTn";
                let behavior_common = matches!(
                    node_local,
                    b"anim"
                        | b"animClr"
                        | b"animEffect"
                        | b"animMotion"
                        | b"animRot"
                        | b"animScale"
                        | b"cmd"
                        | b"set"
                ) && local == b"cBhvr";
                let media_common =
                    matches!(node_local, b"audio" | b"video") && local == b"cMediaNode";
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) != Some(P_NS) {
                    buffer.clear();
                    continue;
                }
                if direct_common {
                    collect_schema_common_time_node_ids(&raw, parent, maximum)?;
                } else if behavior_common || media_common {
                    collect_schema_common_owner_ids(&raw, parent, maximum)?;
                }
            }
            Event::Empty(child) => {
                let parent = root.as_ref().expect("timing node id root");
                let scope = parent.with_start(&child)?;
                let child_name = child.name();
                let local = local_name(child_name.as_ref());
                if matches!(node_local, b"par" | b"seq" | b"excl")
                    && local == b"cTn"
                    && scope.element_uri(child.name().as_ref()) == Some(P_NS)
                {
                    collect_schema_common_time_node_ids(
                        &capture_empty_element(&child)?,
                        parent,
                        maximum,
                    )?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_schema_common_owner_ids(
    xml: &[u8],
    inherited: &NamespaceBindings,
    maximum: &mut u32,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing common owner id root");
                let scope = parent.with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn"
                {
                    collect_schema_common_time_node_ids(&raw, parent, maximum)?;
                }
            }
            Event::Empty(child) => {
                let parent = root.as_ref().expect("timing common owner id root");
                let scope = parent.with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn"
                {
                    collect_schema_common_time_node_ids(
                        &capture_empty_element(&child)?,
                        parent,
                        maximum,
                    )?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_schema_common_time_node_ids(
    xml: &[u8],
    inherited: &NamespaceBindings,
    maximum: &mut u32,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                if let Some(id) = attribute(&all_attributes(&start)?, "id")
                    .and_then(|value| value.parse::<u32>().ok())
                {
                    *maximum = (*maximum).max(id);
                }
                root = Some(inherited.with_start(&start)?);
            }
            Event::Empty(start) if root.is_none() => {
                if let Some(id) = attribute(&all_attributes(&start)?, "id")
                    .and_then(|value| value.parse::<u32>().ok())
                {
                    *maximum = (*maximum).max(id);
                }
                return Ok(());
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("common time node id root");
                let scope = parent.with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && matches!(
                        local_name(child.name().as_ref()),
                        b"childTnLst" | b"subTnLst"
                    )
                {
                    collect_schema_node_list_ids(&raw, parent, maximum)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

#[allow(clippy::too_many_arguments)]
fn authored_media_xml(
    video: bool,
    shape_id: u32,
    node_id: u32,
    volume: u32,
    looped: bool,
    show_when_stopped: bool,
) -> String {
    let tag = if video { "video" } else { "audio" };
    let repeat = if looped {
        " repeatCount=\"indefinite\""
    } else {
        ""
    };
    let show = if show_when_stopped { "1" } else { "0" };
    format!(
        r#"<p:{tag} xmlns:p="{P_NS}"><p:cMediaNode vol="{volume}" showWhenStopped="{show}"><p:cTn id="{node_id}" dur="indefinite"{repeat}/><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cMediaNode></p:{tag}>"#
    )
}

fn authored_media_command_xml(
    shape_id: u32,
    node_id: u32,
    trigger: MediaPlaybackTrigger,
) -> String {
    let (node_type, condition) = match trigger {
        MediaPlaybackTrigger::Automatic => ("", r#"<p:cond delay="0"/>"#),
        MediaPlaybackTrigger::OnClick => (
            r#" nodeType="clickEffect""#,
            r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="SHAPE"/></p:tgtEl></p:cond>"#,
        ),
        MediaPlaybackTrigger::InClickSequence => {
            (r#" nodeType="clickEffect""#, r#"<p:cond delay="0"/>"#)
        }
    };
    let condition = condition.replace("SHAPE", &shape_id.to_string());
    format!(
        r#"<p:cmd xmlns:p="{P_NS}" type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="{node_id}" dur="1" fill="hold"{node_type}><p:stCondLst>{condition}</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr></p:cmd>"#
    )
}

fn append_to_timing_list(
    xml: &[u8],
    inherited: &NamespaceBindings,
    inserted: &[u8],
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut later_child_start = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let scope = inherited.with_start(&start)?;
                validate_element(&start, &scope, b"timing")?;
                root = Some(scope);
            }
            Event::Empty(start) if root.is_none() => {
                let scope = inherited.with_start(&start)?;
                validate_element(&start, &scope, b"timing")?;
                let range = start_tag_range(xml, reader.buffer_position() as usize)?;
                let list = authored_timing_list(inserted);
                return expand_empty_element(xml, range, start.name().as_ref(), list.as_bytes());
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("timing root").with_start(&child)?;
                let range_start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"tnLst"
                {
                    let closing = root_closing_tag_start(&raw)?;
                    return insert_bytes(xml, range_start + closing, inserted);
                }
                if later_child_start.is_none()
                    && scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && matches!(local_name(child.name().as_ref()), b"bldLst" | b"extLst")
                {
                    later_child_start = Some(range_start);
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("timing root").with_start(&child)?;
                let range = event_start..reader.buffer_position() as usize;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"tnLst"
                {
                    return expand_empty_element(xml, range, child.name().as_ref(), inserted);
                }
                if later_child_start.is_none()
                    && scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && matches!(local_name(child.name().as_ref()), b"bldLst" | b"extLst")
                {
                    later_child_start = Some(range.start);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"timing" => {
                let insertion = later_child_start.unwrap_or(event_start);
                let list = authored_timing_list(inserted);
                return insert_bytes(xml, insertion, list.as_bytes());
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing p:timing".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn root_closing_tag_start(xml: &[u8]) -> Result<usize> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OxmlError::InvalidValue("timing list has an unmatched end tag".to_owned())
                })?;
                if depth == 0 {
                    return Ok(event_start);
                }
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("closing p:tnLst".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn authored_timing_list(inserted: &[u8]) -> String {
    format!(
        r#"<p:tnLst xmlns:p="{P_NS}">{}</p:tnLst>"#,
        String::from_utf8_lossy(inserted)
    )
}

fn insert_bytes(xml: &[u8], at: usize, inserted: &[u8]) -> Result<Vec<u8>> {
    if at > xml.len() {
        return Err(OxmlError::InvalidValue(
            "timing insertion point is outside the source".to_owned(),
        ));
    }
    let mut output = Vec::with_capacity(xml.len() + inserted.len());
    output.extend_from_slice(&xml[..at]);
    output.extend_from_slice(inserted);
    output.extend_from_slice(&xml[at..]);
    Ok(output)
}

fn expand_empty_element(
    xml: &[u8],
    range: std::ops::Range<usize>,
    qualified_name: &[u8],
    inserted: &[u8],
) -> Result<Vec<u8>> {
    let slash = xml[range.clone()]
        .iter()
        .rposition(|byte| *byte == b'/')
        .map(|offset| range.start + offset)
        .ok_or_else(|| OxmlError::InvalidValue("empty timing element lacks slash".to_owned()))?;
    let mut output = Vec::with_capacity(xml.len() + inserted.len() + qualified_name.len() + 2);
    output.extend_from_slice(&xml[..slash]);
    output.push(b'>');
    output.extend_from_slice(inserted);
    output.extend_from_slice(b"</");
    output.extend_from_slice(qualified_name);
    output.push(b'>');
    output.extend_from_slice(&xml[range.end..]);
    Ok(output)
}

fn collect_owned_media_ranges(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    shape_id: u32,
    ranges: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing traversal root");
                let scope = parent.with_start(&child)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"tnLst"
                {
                    collect_owned_media_node_list(&raw, parent, base + start, shape_id, ranges)?;
                }
            }
            Event::Empty(_) => {}
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_owned_media_node_list(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    shape_id: u32,
    ranges: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing node list root");
                let scope = parent.with_start(&child)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) != Some(P_NS) {
                    buffer.clear();
                    continue;
                }
                let node = parse_timing_node(&raw, parent, &scope)?;
                let owned = match &node {
                    TimingNode::Audio(media) | TimingNode::Video(media) => {
                        media.common.target == TimingTarget::Shape(shape_id)
                    }
                    TimingNode::MediaCommand(command) => {
                        command.target == TimingTarget::Shape(shape_id)
                    }
                    _ => false,
                };
                if owned {
                    ranges.push(base + start..base + reader.buffer_position() as usize);
                } else if !matches!(node, TimingNode::Unsupported(_)) {
                    collect_owned_media_node_children(
                        &raw,
                        parent,
                        base + start,
                        shape_id,
                        ranges,
                    )?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_owned_media_node_children(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    shape_id: u32,
    ranges: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut root_local = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root_local = local_name(start.name().as_ref()).to_vec();
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing node root");
                let scope = parent.with_start(&child)?;
                let child_name = child.name();
                let local = local_name(child_name.as_ref());
                let direct_common =
                    matches!(root_local.as_slice(), b"par" | b"seq") && local == b"cTn";
                let behavior_common = matches!(
                    root_local.as_slice(),
                    b"set" | b"anim" | b"animEffect" | b"animMotion" | b"cmd"
                ) && local == b"cBhvr";
                let media_common =
                    matches!(root_local.as_slice(), b"audio" | b"video") && local == b"cMediaNode";
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) != Some(P_NS) {
                    buffer.clear();
                    continue;
                }
                if direct_common {
                    collect_owned_media_common_time_node(
                        &raw,
                        parent,
                        base + start,
                        shape_id,
                        ranges,
                    )?;
                } else if behavior_common || media_common {
                    collect_owned_media_common_owner(&raw, parent, base + start, shape_id, ranges)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_owned_media_common_owner(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    shape_id: u32,
    ranges: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("timing common owner root");
                let scope = parent.with_start(&child)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn"
                {
                    collect_owned_media_common_time_node(
                        &raw,
                        parent,
                        base + start,
                        shape_id,
                        ranges,
                    )?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_owned_media_common_time_node(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    shape_id: u32,
    ranges: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let parent = root.as_ref().expect("common time node root");
                let scope = parent.with_start(&child)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"childTnLst"
                {
                    collect_owned_media_node_list(&raw, parent, base + start, shape_id, ranges)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TransitionSpeed {
    Slow,
    Medium,
    Fast,
    Other(String),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TransitionEffect {
    Cut,
    Fade,
    Wipe,
    Push,
    Zoom,
    Morph,
    Other(String),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MorphMetadata {
    pub option: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TransitionParameter {
    pub name: String,
    pub value: String,
}

#[allow(non_camel_case_types)]
#[derive(Clone, Debug)]
pub struct CT_SlideTransition {
    pub speed: Option<TransitionSpeed>,
    pub duration_ms: Option<u64>,
    pub advance_on_click: Option<bool>,
    pub advance_after_ms: Option<u64>,
    pub effect: Option<TransitionEffect>,
    pub effect_parameters: Vec<TransitionParameter>,
    pub morph: Option<MorphMetadata>,
    raw_xml: Vec<u8>,
    inherited_namespaces: Vec<(String, String)>,
}

impl PartialEq for CT_SlideTransition {
    fn eq(&self, other: &Self) -> bool {
        self.speed == other.speed
            && self.duration_ms == other.duration_ms
            && self.advance_on_click == other.advance_on_click
            && self.advance_after_ms == other.advance_after_ms
            && self.effect == other.effect
            && self.effect_parameters == other.effect_parameters
            && self.morph == other.morph
            && self.raw_xml == other.raw_xml
    }
}

impl Eq for CT_SlideTransition {}

impl CT_SlideTransition {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::from_fragment(xml, &NamespaceBindings::default())
    }

    pub(crate) fn from_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        Self::from_projected_fragment(xml, xml, inherited, inherited)
    }

    pub(crate) fn from_alternate_content(
        xml: &[u8],
        inherited: &NamespaceBindings,
    ) -> Result<Option<Self>> {
        let Some(selected) = selected_transition_fragment(xml, inherited)? else {
            return Ok(None);
        };
        Self::from_projected_fragment(&selected.xml, xml, &selected.parent_namespaces, inherited)
            .map(Some)
    }

    fn from_projected_fragment(
        projection_xml: &[u8],
        retained_xml: &[u8],
        projection_inherited: &NamespaceBindings,
        retained_inherited: &NamespaceBindings,
    ) -> Result<Self> {
        let mut reader = Reader::from_reader(projection_xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let namespaces = projection_inherited.with_start(&start)?;
                    validate_element(&start, &namespaces, b"transition")?;
                    let attributes = all_attributes(&start)?;
                    let speed = attribute(&attributes, "spd").map(parse_speed);
                    let duration_ms = parse_optional_u64(
                        namespaced_attribute(&attributes, &namespaces, P14_NS, "dur"),
                        "dur",
                    )?;
                    let advance_on_click =
                        parse_optional_bool(attribute(&attributes, "advClick"), "advClick")?;
                    let advance_after_ms =
                        parse_optional_u64(attribute(&attributes, "advTm"), "advTm")?;
                    let (effect, effect_parameters, morph) =
                        parse_transition_children(&mut reader, &namespaces)?;
                    return Ok(Self {
                        speed,
                        duration_ms,
                        advance_on_click,
                        advance_after_ms,
                        effect,
                        effect_parameters,
                        morph,
                        raw_xml: retained_xml.to_vec(),
                        inherited_namespaces: retained_inherited.entries(),
                    });
                }
                Event::Empty(start) => {
                    let namespaces = projection_inherited.with_start(&start)?;
                    validate_element(&start, &namespaces, b"transition")?;
                    let attributes = all_attributes(&start)?;
                    return Ok(Self {
                        speed: attribute(&attributes, "spd").map(parse_speed),
                        duration_ms: parse_optional_u64(
                            namespaced_attribute(&attributes, &namespaces, P14_NS, "dur"),
                            "dur",
                        )?,
                        advance_on_click: parse_optional_bool(
                            attribute(&attributes, "advClick"),
                            "advClick",
                        )?,
                        advance_after_ms: parse_optional_u64(
                            attribute(&attributes, "advTm"),
                            "advTm",
                        )?,
                        effect: None,
                        effect_parameters: Vec::new(),
                        morph: None,
                        raw_xml: retained_xml.to_vec(),
                        inherited_namespaces: retained_inherited.entries(),
                    });
                }
                Event::Eof => {
                    return Err(OxmlError::MissingElement("p:transition".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    pub fn to_xml(&self) -> Vec<u8> {
        self.raw_xml.clone()
    }

    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }

    /// Changes the transition speed while preserving every child subtree.
    pub fn set_speed(&mut self, speed: TransitionSpeed) -> Result<()> {
        let value = match &speed {
            TransitionSpeed::Slow => "slow",
            TransitionSpeed::Medium => "med",
            TransitionSpeed::Fast => "fast",
            TransitionSpeed::Other(value) => value,
        };
        let inherited = NamespaceBindings::from_entries(&self.inherited_namespaces);
        let raw = rewrite_selected_transition_fragment(
            &self.raw_xml,
            &inherited,
            P_NS,
            b"transition",
            "spd",
            value,
        )?;
        let replacement = Self::from_retained_fragment(&raw, &inherited)?;
        *self = replacement;
        Ok(())
    }

    /// Changes existing morph option metadata without authoring a new choice.
    pub fn set_morph_option(&mut self, option: &str) -> Result<()> {
        let inherited = NamespaceBindings::from_entries(&self.inherited_namespaces);
        let raw = rewrite_selected_transition_fragment(
            &self.raw_xml,
            &inherited,
            P159_NS,
            b"morph",
            "option",
            option,
        )?;
        let replacement = Self::from_retained_fragment(&raw, &inherited)?;
        *self = replacement;
        Ok(())
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.get_mut().write_all(&self.raw_xml)?;
        Ok(())
    }

    fn from_retained_fragment(xml: &[u8], inherited: &NamespaceBindings) -> Result<Self> {
        if root_is(xml, inherited, MC_NS, b"AlternateContent")? {
            return Self::from_alternate_content(xml, inherited)?.ok_or_else(|| {
                OxmlError::MissingElement("p:transition in mc:AlternateContent".to_owned())
            });
        }
        Self::from_fragment(xml, inherited)
    }
}

fn parse_timing_children(
    reader: &mut Reader<&[u8]>,
    namespaces: &NamespaceBindings,
) -> Result<(Vec<TimingNode>, Vec<TimingBuild>)> {
    let mut nodes = Vec::new();
    let mut builds = Vec::new();
    let mut seen_tn_list = false;
    let mut seen_build_list = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(reader, &child)?;
                if child_ns.element_uri(child.name().as_ref()) == Some(P_NS) && local == b"tnLst" {
                    if seen_tn_list {
                        return Err(duplicate("tnLst"));
                    }
                    if seen_build_list {
                        return Err(out_of_order("tnLst"));
                    }
                    seen_tn_list = true;
                    nodes = parse_node_list(&raw, namespaces, b"tnLst")?;
                } else if child_ns.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local == b"bldLst"
                {
                    if seen_build_list {
                        return Err(duplicate("bldLst"));
                    }
                    seen_build_list = true;
                    builds = parse_build_list(&raw, namespaces)?;
                }
            }
            Event::Empty(child) => {
                let child_ns = namespaces.with_start(&child)?;
                let child_name = child.name();
                let local = local_name(child_name.as_ref());
                if child_ns.element_uri(child.name().as_ref()) == Some(P_NS) && local == b"tnLst" {
                    if seen_tn_list {
                        return Err(duplicate("tnLst"));
                    }
                    if seen_build_list {
                        return Err(out_of_order("tnLst"));
                    }
                    seen_tn_list = true;
                } else if child_ns.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local == b"bldLst"
                {
                    if seen_build_list {
                        return Err(duplicate("bldLst"));
                    }
                    seen_build_list = true;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"timing" => {
                return Ok((nodes, builds));
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing p:timing".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_node_list(
    xml: &[u8],
    inherited: &NamespaceBindings,
    expected: &[u8],
) -> Result<Vec<TimingNode>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut nodes = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let namespaces = inherited.with_start(&start)?;
                validate_element(&start, &namespaces, expected)?;
                root = Some(namespaces);
            }
            Event::Start(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                nodes.push(parse_timing_node(
                    &raw,
                    root.as_ref().expect("root set"),
                    &namespaces,
                )?);
            }
            Event::Empty(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                let raw = capture_empty_element(&child)?;
                nodes.push(parse_timing_node(
                    &raw,
                    root.as_ref().expect("root set"),
                    &namespaces,
                )?);
            }
            Event::End(end) if local_name(end.name().as_ref()) == expected => return Ok(nodes),
            Event::Eof => return Ok(nodes),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_timing_node(
    xml: &[u8],
    inherited: &NamespaceBindings,
    element_namespaces: &NamespaceBindings,
) -> Result<TimingNode> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) | Event::Empty(start) => {
                let start_name = start.name();
                let local = local_name(start_name.as_ref());
                let is_p = element_namespaces.element_uri(start.name().as_ref()) == Some(P_NS);
                if !is_p {
                    return Ok(unsupported(local, xml));
                }
                return match local {
                    b"par" => Ok(TimingNode::Parallel(TimingContainer {
                        common: parse_container_common(xml, inherited)?,
                    })),
                    b"seq" => Ok(TimingNode::Sequence(parse_sequence(xml, inherited)?)),
                    b"set" => {
                        let behavior = parse_behavior(xml, inherited)?;
                        Ok(TimingNode::Set(TimingSet {
                            common: behavior.common,
                            target: behavior.target,
                            attribute_name: behavior.attribute_name,
                            value: parse_set_value(xml, inherited)?,
                        }))
                    }
                    b"anim" => {
                        let behavior = parse_behavior(xml, inherited)?;
                        let attrs = all_attributes(&start)?;
                        Ok(TimingNode::Animate(TimingAnimate {
                            common: behavior.common,
                            target: behavior.target,
                            attribute_name: behavior.attribute_name,
                            from: attribute(&attrs, "from").map(str::to_owned),
                            to: attribute(&attrs, "to").map(str::to_owned),
                            by: attribute(&attrs, "by").map(str::to_owned),
                        }))
                    }
                    b"animEffect" => {
                        let behavior = parse_behavior(xml, inherited)?;
                        let attrs = all_attributes(&start)?;
                        Ok(TimingNode::Effect(TimingEffect {
                            common: behavior.common,
                            target: behavior.target,
                            transition: attribute(&attrs, "transition").map(str::to_owned),
                            filter: attribute(&attrs, "filter").map(str::to_owned),
                        }))
                    }
                    b"animMotion" => {
                        let behavior = parse_behavior(xml, inherited)?;
                        let attrs = all_attributes(&start)?;
                        Ok(TimingNode::Motion(TimingMotionPath {
                            common: behavior.common,
                            target: behavior.target,
                            path: attribute(&attrs, "path").map(str::to_owned),
                            origin: attribute(&attrs, "origin").map(str::to_owned),
                        }))
                    }
                    b"audio" => Ok(TimingNode::Audio(parse_media(xml, inherited, false)?)),
                    b"video" => Ok(TimingNode::Video(parse_media(xml, inherited, true)?)),
                    b"cmd" => match parse_media_command(xml, inherited) {
                        Ok(command) => Ok(TimingNode::MediaCommand(command)),
                        Err(OxmlError::MissingElement(_)) => Ok(unsupported(local, xml)),
                        Err(error) => Err(error),
                    },
                    _ => Ok(unsupported(local, xml)),
                };
            }
            Event::Eof => return Err(OxmlError::MissingElement("timing node".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_media(xml: &[u8], inherited: &NamespaceBindings, video: bool) -> Result<TimingMedia> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut full_screen = None;
    let mut common = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let attributes = all_attributes(&start)?;
                full_screen = parse_optional_bool(attribute(&attributes, "fullScrn"), "fullScrn")?;
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("media root").with_start(&child)?;
                let is_common = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cMediaNode";
                let raw = capture_element(&mut reader, &child)?;
                if is_common {
                    set_once(
                        &mut common,
                        parse_common_media_node(&raw, root.as_ref().expect("media root"))?,
                        "cMediaNode",
                    )?;
                }
            }
            Event::Empty(_) if root.is_none() => {
                return Err(OxmlError::MissingElement("p:cMediaNode".to_owned()));
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("media root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cMediaNode"
                {
                    return Err(OxmlError::MissingElement("p:cTn".to_owned()));
                }
            }
            Event::End(end)
                if local_name(end.name().as_ref()) == if video { b"video" } else { b"audio" } =>
            {
                return Ok(TimingMedia {
                    common: common
                        .ok_or_else(|| OxmlError::MissingElement("p:cMediaNode".to_owned()))?,
                    full_screen,
                });
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing media node".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_common_media_node(xml: &[u8], inherited: &NamespaceBindings) -> Result<CommonMediaNode> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut attributes = Vec::new();
    let mut common = None;
    let mut common_attributes = None;
    let mut target = None;
    let mut boundary = 0u8;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                attributes = all_attributes(&start)?;
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root
                    .as_ref()
                    .expect("common media root")
                    .with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = match local.as_slice() {
                        b"cTn" => 1,
                        b"tgtEl" => 2,
                        _ => 0,
                    };
                    if rank != 0 && rank < boundary {
                        return Err(out_of_order(&String::from_utf8_lossy(&local)));
                    }
                    boundary = boundary.max(rank);
                    match local.as_slice() {
                        b"cTn" => {
                            let parsed = parse_common_time_node(
                                &raw,
                                root.as_ref().expect("common media root"),
                            )?;
                            set_once(&mut common, parsed, "cTn")?;
                            common_attributes = Some(all_attributes(&child)?);
                        }
                        b"tgtEl" => set_once(
                            &mut target,
                            parse_target(&raw, root.as_ref().expect("common media root"))?,
                            "tgtEl",
                        )?,
                        _ => {}
                    }
                }
            }
            Event::Empty(child) => {
                let scope = root
                    .as_ref()
                    .expect("common media root")
                    .with_start(&child)?;
                let child_name = child.name();
                let local = local_name(child_name.as_ref());
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) && local == b"cTn" {
                    let parsed = parse_common_time_node(
                        &capture_empty_element(&child)?,
                        root.as_ref().expect("common media root"),
                    )?;
                    set_once(&mut common, parsed, "cTn")?;
                    common_attributes = Some(all_attributes(&child)?);
                    boundary = 1;
                } else if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local == b"tgtEl"
                {
                    set_once(&mut target, TimingTarget::Unsupported, "tgtEl")?;
                    boundary = 2;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"cMediaNode" => {
                let common = common.ok_or_else(|| OxmlError::MissingElement("p:cTn".to_owned()))?;
                let common_attributes = common_attributes
                    .ok_or_else(|| OxmlError::MissingElement("p:cTn".to_owned()))?;
                let volume = parse_optional_u32(attribute(&attributes, "vol"), "vol")?;
                if volume.is_some_and(|volume| volume > 100_000) {
                    return Err(OxmlError::InvalidValue(
                        "media volume must be between 0 and 100000".to_owned(),
                    ));
                }
                let show_when_stopped = parse_optional_bool(
                    attribute(&attributes, "showWhenStopped"),
                    "showWhenStopped",
                )?;
                let looped = attribute(&common_attributes, "repeatCount")
                    .is_some_and(|value| value == "indefinite");
                let display = show_when_stopped
                    .or(parse_optional_bool(
                        attribute(&common_attributes, "display"),
                        "display",
                    )?)
                    .map(|show| {
                        if show {
                            MediaDisplayPolicy::ShowWhenStopped
                        } else {
                            MediaDisplayPolicy::HideWhenStopped
                        }
                    });
                return Ok(CommonMediaNode {
                    common,
                    target: target.unwrap_or(TimingTarget::Unsupported),
                    volume,
                    looped,
                    display,
                });
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("closing p:cMediaNode".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_media_command(xml: &[u8], inherited: &NamespaceBindings) -> Result<TimingMediaCommand> {
    let behavior = parse_behavior(xml, inherited)?;
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) | Event::Empty(start) => {
                let attributes = all_attributes(&start)?;
                let lexical = attribute(&attributes, "cmd").unwrap_or_default();
                return Ok(TimingMediaCommand {
                    common: behavior.common,
                    target: behavior.target,
                    command: parse_media_command_kind(lexical)?,
                });
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:cmd".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_media_command_kind(value: &str) -> Result<MediaCommandKind> {
    match value {
        "play" => Ok(MediaCommandKind::Play { from_ms: None }),
        "pause" | "togglePause" => Ok(MediaCommandKind::Pause),
        "stop" => Ok(MediaCommandKind::Stop),
        _ if value.starts_with("playFrom(") && value.ends_with(')') => {
            let value = &value[9..value.len() - 1];
            Ok(parse_media_seconds(value)
                .map(|from_ms| MediaCommandKind::Play {
                    from_ms: Some(from_ms),
                })
                .unwrap_or_else(|_| MediaCommandKind::Other(format!("playFrom({value})"))))
        }
        _ if value.starts_with("seek(") && value.ends_with(')') => {
            let value = &value[5..value.len() - 1];
            Ok(parse_media_seconds(value)
                .map(|position_ms| MediaCommandKind::Seek { position_ms })
                .unwrap_or_else(|_| MediaCommandKind::Other(format!("seek({value})"))))
        }
        _ => Ok(MediaCommandKind::Other(value.to_owned())),
    }
}

fn parse_media_seconds(value: &str) -> Result<u64> {
    let seconds = value
        .parse::<f64>()
        .map_err(|_| OxmlError::InvalidValue(format!("invalid media command offset {value}")))?;
    let milliseconds = seconds * 1_000.0;
    if !milliseconds.is_finite() || milliseconds < 0.0 || milliseconds > u64::MAX as f64 {
        return Err(OxmlError::InvalidValue(format!(
            "media command offset is out of range: {value}"
        )));
    }
    Ok(milliseconds.round() as u64)
}

struct ParsedBehavior {
    common: CommonTimeNode,
    target: TimingTarget,
    attribute_name: Option<String>,
}

fn parse_behavior(xml: &[u8], inherited: &NamespaceBindings) -> Result<ParsedBehavior> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut behavior = None;
    let mut passed_behavior_slot = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let is_behavior = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cBhvr";
                let raw = capture_element(&mut reader, &child)?;
                if is_behavior {
                    if passed_behavior_slot {
                        return Err(out_of_order("cBhvr"));
                    }
                    set_once(
                        &mut behavior,
                        parse_common_behavior(&raw, root.as_ref().expect("root"))?,
                        "cBhvr",
                    )?;
                } else if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    passed_behavior_slot = true;
                }
            }
            Event::Empty(_) if root.is_none() => {
                return Err(OxmlError::MissingElement("p:cBhvr".to_owned()));
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cBhvr"
                {
                    if behavior.is_some() {
                        return Err(duplicate("cBhvr"));
                    }
                    return Err(OxmlError::MissingElement("p:cTn".to_owned()));
                }
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    passed_behavior_slot = true;
                }
            }
            Event::End(_) => {
                return behavior.ok_or_else(|| OxmlError::MissingElement("p:cBhvr".to_owned()));
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:cBhvr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_common_behavior(xml: &[u8], inherited: &NamespaceBindings) -> Result<ParsedBehavior> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut common = None;
    let mut target = None;
    let mut attribute_name = None;
    let mut boundary = 0u8;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = behavior_child_rank(&local);
                    if rank != 0 && rank < boundary {
                        return Err(out_of_order(&String::from_utf8_lossy(&local)));
                    }
                    boundary = boundary.max(rank);
                    match local.as_slice() {
                        b"cTn" => set_once(
                            &mut common,
                            parse_common_time_node(&raw, root.as_ref().expect("root"))?,
                            "cTn",
                        )?,
                        b"tgtEl" => set_once(
                            &mut target,
                            parse_target(&raw, root.as_ref().expect("root"))?,
                            "tgtEl",
                        )?,
                        b"attrNameLst" => set_once(
                            &mut attribute_name,
                            parse_attribute_name(&raw, root.as_ref().expect("root"))?,
                            "attrNameLst",
                        )?,
                        _ => {}
                    }
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = behavior_child_rank(&local);
                    if rank != 0 && rank < boundary {
                        return Err(out_of_order(&String::from_utf8_lossy(&local)));
                    }
                    boundary = boundary.max(rank);
                    match local.as_slice() {
                        b"cTn" => set_once(
                            &mut common,
                            parse_common_time_node(
                                &capture_empty_element(&child)?,
                                root.as_ref().expect("root"),
                            )?,
                            "cTn",
                        )?,
                        b"tgtEl" => set_once(&mut target, TimingTarget::Unsupported, "tgtEl")?,
                        b"attrNameLst" => set_once(&mut attribute_name, None, "attrNameLst")?,
                        _ => {}
                    }
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"cBhvr" => {
                return Ok(ParsedBehavior {
                    common: common.ok_or_else(|| OxmlError::MissingElement("p:cTn".to_owned()))?,
                    target: target.unwrap_or(TimingTarget::Unsupported),
                    attribute_name: attribute_name.flatten(),
                });
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing p:cBhvr".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn behavior_child_rank(local: &[u8]) -> u8 {
    match local {
        b"cTn" => 1,
        b"tgtEl" => 2,
        b"attrNameLst" => 3,
        _ => 0,
    }
}

fn parse_container_common(xml: &[u8], inherited: &NamespaceBindings) -> Result<CommonTimeNode> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut common = None;
    let mut passed_common_slot = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let is_common = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn";
                let raw = capture_element(&mut reader, &child)?;
                if is_common {
                    if passed_common_slot {
                        return Err(out_of_order("cTn"));
                    }
                    set_once(
                        &mut common,
                        parse_common_time_node(&raw, root.as_ref().expect("root"))?,
                        "cTn",
                    )?;
                } else if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    passed_common_slot = true;
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn"
                {
                    if passed_common_slot {
                        return Err(out_of_order("cTn"));
                    }
                    set_once(
                        &mut common,
                        parse_common_time_node(
                            &capture_empty_element(&child)?,
                            root.as_ref().expect("root"),
                        )?,
                        "cTn",
                    )?;
                } else if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    passed_common_slot = true;
                }
            }
            Event::End(_) => {
                return common.ok_or_else(|| OxmlError::MissingElement("p:cTn".to_owned()));
            }
            Event::Eof => return Err(OxmlError::MissingElement("p:cTn".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_sequence(xml: &[u8], inherited: &NamespaceBindings) -> Result<TimingSequence> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut common = None;
    let mut previous_conditions = None;
    let mut next_conditions = None;
    let mut attributes = Vec::new();
    let mut boundary = 0u8;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                attributes = all_attributes(&start)?;
                let scope = inherited.with_start(&start)?;
                validate_element(&start, &scope, b"seq")?;
                root = Some(scope);
            }
            Event::Empty(start) if root.is_none() => {
                let scope = inherited.with_start(&start)?;
                validate_element(&start, &scope, b"seq")?;
                return Err(OxmlError::MissingElement("p:cTn".to_owned()));
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = sequence_child_rank(&local);
                    if rank != 0 && rank < boundary {
                        return Err(out_of_order(&String::from_utf8_lossy(&local)));
                    }
                    boundary = boundary.max(rank);
                    match local.as_slice() {
                        b"cTn" => set_once(
                            &mut common,
                            parse_common_time_node(&raw, root.as_ref().expect("root"))?,
                            "cTn",
                        )?,
                        b"prevCondLst" => set_once(
                            &mut previous_conditions,
                            parse_condition_list(
                                &raw,
                                root.as_ref().expect("root"),
                                b"prevCondLst",
                            )?,
                            "prevCondLst",
                        )?,
                        b"nextCondLst" => set_once(
                            &mut next_conditions,
                            parse_condition_list(
                                &raw,
                                root.as_ref().expect("root"),
                                b"nextCondLst",
                            )?,
                            "nextCondLst",
                        )?,
                        _ => {}
                    }
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = sequence_child_rank(&local);
                    if rank != 0 && rank < boundary {
                        return Err(out_of_order(&String::from_utf8_lossy(&local)));
                    }
                    boundary = boundary.max(rank);
                    match local.as_slice() {
                        b"cTn" => set_once(
                            &mut common,
                            parse_common_time_node(
                                &capture_empty_element(&child)?,
                                root.as_ref().expect("root"),
                            )?,
                            "cTn",
                        )?,
                        b"prevCondLst" => {
                            set_once(&mut previous_conditions, Vec::new(), "prevCondLst")?
                        }
                        b"nextCondLst" => {
                            set_once(&mut next_conditions, Vec::new(), "nextCondLst")?
                        }
                        _ => {}
                    }
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"seq" => {
                return Ok(TimingSequence {
                    common: common.ok_or_else(|| OxmlError::MissingElement("p:cTn".to_owned()))?,
                    concurrent: parse_optional_bool(
                        attribute(&attributes, "concurrent"),
                        "concurrent",
                    )?,
                    next_action: attribute(&attributes, "nextAc").map(str::to_owned),
                    previous_conditions: previous_conditions.unwrap_or_default(),
                    next_conditions: next_conditions.unwrap_or_default(),
                });
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing p:seq".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_set_value(xml: &[u8], inherited: &NamespaceBindings) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut value = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let is_to = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"to";
                let raw = capture_element(&mut reader, &child)?;
                if is_to {
                    set_once(
                        &mut value,
                        parse_animation_variant(&raw, root.as_ref().expect("root"))?,
                        "to",
                    )?;
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"to"
                {
                    set_once(&mut value, None, "to")?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(value.flatten()),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_animation_variant(xml: &[u8], inherited: &NamespaceBindings) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut value = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(variant) => {
                let scope = root.as_ref().expect("root").with_start(&variant)?;
                let modeled = scope.element_uri(variant.name().as_ref()) == Some(P_NS)
                    && is_animation_variant(local_name(variant.name().as_ref()));
                let attributes = all_attributes(&variant)?;
                let _ = capture_element(&mut reader, &variant)?;
                if modeled {
                    set_once(
                        &mut value,
                        attribute(&attributes, "val").map(str::to_owned),
                        "animation variant",
                    )?;
                }
            }
            Event::Empty(variant) => {
                let scope = root.as_ref().expect("root").with_start(&variant)?;
                if scope.element_uri(variant.name().as_ref()) != Some(P_NS)
                    || !is_animation_variant(local_name(variant.name().as_ref()))
                {
                    continue;
                }
                let attributes = all_attributes(&variant)?;
                set_once(
                    &mut value,
                    attribute(&attributes, "val").map(str::to_owned),
                    "animation variant",
                )?;
            }
            Event::End(_) | Event::Eof => return Ok(value.flatten()),
            _ => {}
        }
        buffer.clear();
    }
}

fn is_animation_variant(local: &[u8]) -> bool {
    matches!(
        local,
        b"boolVal" | b"clrVal" | b"fltVal" | b"intVal" | b"strVal"
    )
}

fn sequence_child_rank(local: &[u8]) -> u8 {
    match local {
        b"cTn" => 1,
        b"prevCondLst" => 2,
        b"nextCondLst" => 3,
        _ => 0,
    }
}

fn parse_common_time_node(xml: &[u8], inherited: &NamespaceBindings) -> Result<CommonTimeNode> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut value = None;
    let mut seen_children = [false; 7];
    let mut boundary = 0usize;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let scope = inherited.with_start(&start)?;
                validate_element(&start, &scope, b"cTn")?;
                value = Some(common_from_start(&start)?);
                root = Some(scope);
            }
            Event::Empty(start) if root.is_none() => return common_from_start(&start),
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = common_child_rank(&local);
                    check_sequence_child(&mut seen_children, &mut boundary, rank, &local)?;
                    if local == b"stCondLst" {
                        value.as_mut().expect("value").start_conditions =
                            parse_condition_list(&raw, root.as_ref().expect("root"), b"stCondLst")?;
                    } else if local == b"endCondLst" {
                        value.as_mut().expect("value").end_conditions = parse_condition_list(
                            &raw,
                            root.as_ref().expect("root"),
                            b"endCondLst",
                        )?;
                    } else if local == b"childTnLst" {
                        value.as_mut().expect("value").children =
                            parse_node_list(&raw, root.as_ref().expect("root"), b"childTnLst")?;
                    }
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    let rank = common_child_rank(&local);
                    check_sequence_child(&mut seen_children, &mut boundary, rank, &local)?;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"cTn" => {
                return value.ok_or_else(|| OxmlError::MissingElement("p:cTn".to_owned()));
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing p:cTn".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn common_from_start(start: &BytesStart<'_>) -> Result<CommonTimeNode> {
    let attrs = all_attributes(start)?;
    let id = attribute(&attrs, "id")
        .ok_or_else(|| OxmlError::MissingElement("p:cTn/@id".to_owned()))?
        .parse::<u32>()
        .map_err(|_| OxmlError::InvalidValue("invalid p:cTn/@id".to_owned()))?;
    Ok(CommonTimeNode {
        id,
        group_id: parse_optional_u32(attribute(&attrs, "grpId"), "grpId")?,
        duration: TimingDuration::parse(attribute(&attrs, "dur"))?,
        fill: attribute(&attrs, "fill").map(parse_fill),
        restart: attribute(&attrs, "restart").map(parse_restart),
        node_type: attribute(&attrs, "nodeType").map(parse_node_type),
        preset_id: parse_optional_u32(attribute(&attrs, "presetID"), "presetID")?,
        preset_class: attribute(&attrs, "presetClass").map(str::to_owned),
        preset_subtype: parse_optional_u32(attribute(&attrs, "presetSubtype"), "presetSubtype")?,
        start_conditions: Vec::new(),
        end_conditions: Vec::new(),
        children: Vec::new(),
    })
}

fn common_child_rank(local: &[u8]) -> usize {
    match local {
        b"stCondLst" => 1,
        b"endCondLst" => 2,
        b"endSync" => 3,
        b"iterate" => 4,
        b"childTnLst" => 5,
        b"subTnLst" => 6,
        _ => 0,
    }
}

fn check_sequence_child(
    seen: &mut [bool; 7],
    boundary: &mut usize,
    rank: usize,
    local: &[u8],
) -> Result<()> {
    if rank == 0 {
        return Ok(());
    }
    if seen[rank] {
        return Err(duplicate(&String::from_utf8_lossy(local)));
    }
    if rank < *boundary {
        return Err(out_of_order(&String::from_utf8_lossy(local)));
    }
    seen[rank] = true;
    *boundary = rank;
    Ok(())
}

fn parse_condition_list(
    xml: &[u8],
    inherited: &NamespaceBindings,
    expected: &[u8],
) -> Result<Vec<TimingCondition>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut conditions = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cond"
                {
                    conditions.push(parse_condition(&raw, root.as_ref().expect("root"))?);
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cond"
                {
                    conditions.push(condition_from_start(&child, TimingTarget::Unsupported)?);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == expected => {
                return Ok(conditions);
            }
            Event::Eof => return Ok(conditions),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_condition_target_presence(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<HashMap<(u32, bool, usize), bool>> {
    let mut presence = HashMap::new();
    collect_condition_target_presence_in_element(xml, inherited, &mut presence)?;
    Ok(presence)
}

fn collect_condition_target_presence_in_element(
    xml: &[u8],
    inherited: &NamespaceBindings,
    presence: &mut HashMap<(u32, bool, usize), bool>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut node_id = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let namespaces = inherited.with_start(&start)?;
                if namespaces.element_uri(start.name().as_ref()) == Some(P_NS)
                    && local_name(start.name().as_ref()) == b"cTn"
                {
                    let attrs = all_attributes(&start)?;
                    node_id = attribute(&attrs, "id").and_then(|value| value.parse().ok());
                }
                root = Some(namespaces);
            }
            Event::Start(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &child)?;
                if let Some(node_id) = node_id
                    && namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && matches!(local.as_slice(), b"stCondLst" | b"endCondLst")
                {
                    collect_condition_list_target_presence(
                        &raw,
                        root.as_ref().expect("root set"),
                        node_id,
                        local == b"endCondLst",
                        presence,
                    )?;
                } else {
                    collect_condition_target_presence_in_element(
                        &raw,
                        root.as_ref().expect("root set"),
                        presence,
                    )?;
                }
            }
            Event::Empty(_) => {}
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_condition_list_target_presence(
    xml: &[u8],
    inherited: &NamespaceBindings,
    node_id: u32,
    end_condition: bool,
    presence: &mut HashMap<(u32, bool, usize), bool>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut index = 0;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                let raw = capture_element(&mut reader, &child)?;
                if namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cond"
                {
                    presence.insert(
                        (node_id, end_condition, index),
                        condition_fragment_has_explicit_target(
                            &raw,
                            root.as_ref().expect("root set"),
                        )?,
                    );
                    index += 1;
                }
            }
            Event::Empty(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                if namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cond"
                {
                    presence.insert((node_id, end_condition, index), false);
                    index += 1;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn condition_fragment_has_explicit_target(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<bool> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                let explicit = namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && (local_name(child.name().as_ref()) == b"tgtEl"
                        || condition_direct_target_from_start(&child, &namespaces)?.is_some());
                let _ = capture_element(&mut reader, &child)?;
                if explicit {
                    return Ok(true);
                }
            }
            Event::Empty(child) => {
                let namespaces = root.as_ref().expect("root set").with_start(&child)?;
                if namespaces.element_uri(child.name().as_ref()) == Some(P_NS)
                    && (local_name(child.name().as_ref()) == b"tgtEl"
                        || condition_direct_target_from_start(&child, &namespaces)?.is_some())
                {
                    return Ok(true);
                }
            }
            Event::End(_) | Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_condition(xml: &[u8], inherited: &NamespaceBindings) -> Result<TimingCondition> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut attrs = None;
    let mut target = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                attrs = Some(start.to_owned());
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                let direct_target = condition_direct_target_from_start(&child, &scope)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) && local == b"tgtEl" {
                    set_once(
                        &mut target,
                        parse_target(&raw, root.as_ref().expect("root"))?,
                        "condition target",
                    )?;
                } else if let Some(direct_target) = direct_target {
                    set_once(&mut target, direct_target, "condition target")?;
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let local = local_name(child.name().as_ref()).to_vec();
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) && local == b"tgtEl" {
                    set_once(&mut target, TimingTarget::Unsupported, "condition target")?;
                } else if let Some(direct_target) =
                    condition_direct_target_from_start(&child, &scope)?
                {
                    set_once(&mut target, direct_target, "condition target")?;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"cond" => {
                return condition_from_start(
                    attrs.as_ref().expect("attrs"),
                    target.unwrap_or(TimingTarget::Unsupported),
                );
            }
            Event::Eof => return Err(OxmlError::MissingElement("closing p:cond".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn condition_from_start(start: &BytesStart<'_>, target: TimingTarget) -> Result<TimingCondition> {
    let attrs = all_attributes(start)?;
    Ok(TimingCondition {
        event: attribute(&attrs, "evt").map(parse_event),
        delay: TimingDuration::parse(attribute(&attrs, "delay"))?,
        target,
    })
}

fn parse_target(xml: &[u8], inherited: &NamespaceBindings) -> Result<TimingTarget> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut target = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let parsed = target_element_from_start(&child, &scope)?;
                let _ = capture_element(&mut reader, &child)?;
                if let Some(parsed) = parsed {
                    set_once(&mut target, parsed, "tgtEl target")?;
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if let Some(parsed) = target_element_from_start(&child, &scope)? {
                    set_once(&mut target, parsed, "tgtEl target")?;
                }
            }
            Event::End(_) => return Ok(target.unwrap_or(TimingTarget::Unsupported)),
            Event::Eof => return Ok(target.unwrap_or(TimingTarget::Unsupported)),
            _ => {}
        }
        buffer.clear();
    }
}

fn target_element_from_start(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<Option<TimingTarget>> {
    if namespaces.element_uri(start.name().as_ref()) != Some(P_NS) {
        return Ok(None);
    }
    let attrs = all_attributes(start)?;
    match local_name(start.name().as_ref()) {
        b"spTgt" => parse_required_u32(attribute(&attrs, "spid"), "spid")
            .map(TimingTarget::Shape)
            .map(Some),
        b"sldTgt" => Ok(Some(TimingTarget::Slide)),
        b"sndTgt" | b"inkTgt" => Ok(Some(TimingTarget::Unsupported)),
        _ => Ok(None),
    }
}

fn condition_direct_target_from_start(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<Option<TimingTarget>> {
    if namespaces.element_uri(start.name().as_ref()) != Some(P_NS) {
        return Ok(None);
    }
    match local_name(start.name().as_ref()) {
        b"tn" => {
            let attrs = all_attributes(start)?;
            parse_required_u32(attribute(&attrs, "val"), "val")
                .map(TimingTarget::TimeNode)
                .map(Some)
        }
        b"rtn" => Ok(Some(TimingTarget::Unsupported)),
        _ => Ok(None),
    }
}

fn parse_build_list(xml: &[u8], inherited: &NamespaceBindings) -> Result<Vec<TimingBuild>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut builds = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => root = Some(inherited.with_start(&start)?),
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    builds.push(build_from_start(&child)?);
                }
                let _ = capture_element(&mut reader, &child)?;
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS) {
                    builds.push(build_from_start(&child)?);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"bldLst" => return Ok(builds),
            Event::Eof => return Ok(builds),
            _ => {}
        }
        buffer.clear();
    }
}

fn build_from_start(start: &BytesStart<'_>) -> Result<TimingBuild> {
    let attrs = all_attributes(start)?;
    Ok(TimingBuild {
        kind: String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned(),
        shape_id: parse_optional_u32(attribute(&attrs, "spid"), "spid")?,
        group_id: parse_optional_u32(attribute(&attrs, "grpId"), "grpId")?,
        build_mode: attribute(&attrs, "build").map(str::to_owned),
        build_level: parse_optional_u32(attribute(&attrs, "bldLvl"), "bldLvl")?,
        reverse: parse_optional_bool(attribute(&attrs, "rev"), "rev")?,
        advance_automatically: attribute(&attrs, "advAuto")
            .map(|value| TimingDuration::parse(Some(value)))
            .transpose()?,
        animate_background: parse_optional_bool(attribute(&attrs, "animBg"), "animBg")?,
        auto_update_animated_background: parse_optional_bool(
            attribute(&attrs, "autoUpdateAnimBg"),
            "autoUpdateAnimBg",
        )?,
        ui_expand: parse_optional_bool(attribute(&attrs, "uiExpand"), "uiExpand")?,
    })
}

fn selected_transition_fragment(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<SelectedFragment>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut fallback = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let scope = inherited.with_start(&start)?;
                if scope.element_uri(start.name().as_ref()) != Some(MC_NS)
                    || local_name(start.name().as_ref()) != b"AlternateContent"
                {
                    return Err(OxmlError::UnexpectedElement(
                        String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    ));
                }
                root = Some(scope);
            }
            Event::Start(branch) => {
                let wrapper = root.as_ref().expect("root");
                let branch_scope = wrapper.with_start(&branch)?;
                let is_mc = branch_scope.element_uri(branch.name().as_ref()) == Some(MC_NS);
                let local = local_name(branch.name().as_ref()).to_vec();
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &branch)?;
                let range = start..reader.buffer_position() as usize;
                if is_mc && local == b"Choice" && choice_is_supported(&branch, &branch_scope)? {
                    return direct_transition_in_branch(&raw, wrapper)
                        .map(|selected| selected.map(|selected| offset_fragment(selected, range)));
                }
                if is_mc && local == b"Fallback" && fallback.is_none() {
                    fallback = Some((raw, wrapper.clone(), range));
                }
            }
            Event::Empty(branch) => {
                let wrapper = root.as_ref().expect("root");
                let branch_scope = wrapper.with_start(&branch)?;
                let is_mc = branch_scope.element_uri(branch.name().as_ref()) == Some(MC_NS);
                let local = local_name(branch.name().as_ref()).to_vec();
                if is_mc && local == b"Choice" && choice_is_supported(&branch, &branch_scope)? {
                    return Ok(None);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"AlternateContent" => {
                let Some((raw, wrapper, range)) = fallback else {
                    return Ok(None);
                };
                return direct_transition_in_branch(&raw, &wrapper)
                    .map(|selected| selected.map(|selected| offset_fragment(selected, range)));
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

fn choice_is_supported(start: &BytesStart<'_>, namespaces: &NamespaceBindings) -> Result<bool> {
    let attributes = all_attributes(start)?;
    let Some(requires) = attribute(&attributes, "Requires") else {
        return Ok(false);
    };
    let entries = namespaces.entries();
    let mut prefixes = requires.split_whitespace().peekable();
    if prefixes.peek().is_none() {
        return Ok(false);
    }
    Ok(prefixes.all(|prefix| {
        entries
            .iter()
            .find(|(candidate, _)| candidate == prefix)
            .is_some_and(|(_, uri)| matches!(uri.as_str(), P_NS | P14_NS | P159_NS))
    }))
}

fn direct_transition_in_branch(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<SelectedFragment>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut selected = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root").clone();
                let scope = parent.with_start(&child)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"transition"
                {
                    set_once(
                        &mut selected,
                        SelectedFragment {
                            xml: raw,
                            parent_namespaces: parent,
                            range: start..reader.buffer_position() as usize,
                        },
                        "transition in selected compatibility branch",
                    )?;
                }
            }
            Event::Empty(child) => {
                let parent = root.as_ref().expect("root").clone();
                let scope = parent.with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"transition"
                {
                    let range = start_tag_range(xml, reader.buffer_position() as usize)?;
                    set_once(
                        &mut selected,
                        SelectedFragment {
                            xml: capture_empty_element(&child)?,
                            parent_namespaces: parent,
                            range,
                        },
                        "transition in selected compatibility branch",
                    )?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(selected),
            _ => {}
        }
        buffer.clear();
    }
}

fn offset_fragment(
    mut selected: SelectedFragment,
    branch_range: std::ops::Range<usize>,
) -> SelectedFragment {
    selected.range =
        branch_range.start + selected.range.start..branch_range.start + selected.range.end;
    selected
}

fn rewrite_selected_transition_fragment(
    xml: &[u8],
    inherited: &NamespaceBindings,
    namespace_uri: &str,
    local: &[u8],
    attribute_name: &str,
    value: &str,
) -> Result<Vec<u8>> {
    if !root_is(xml, inherited, MC_NS, b"AlternateContent")? {
        if namespace_uri == P159_NS && local == b"morph" {
            return rewrite_selected_effect_attribute(
                xml,
                inherited,
                namespace_uri,
                local,
                attribute_name,
                value,
            );
        }
        return rewrite_first_namespaced_attribute(
            xml,
            inherited,
            namespace_uri,
            local,
            attribute_name,
            value,
        );
    }
    let selected = selected_transition_fragment(xml, inherited)?.ok_or_else(|| {
        OxmlError::MissingElement("selected p:transition in mc:AlternateContent".to_owned())
    })?;
    let replacement = if namespace_uri == P159_NS && local == b"morph" {
        rewrite_selected_effect_attribute(
            &selected.xml,
            &selected.parent_namespaces,
            namespace_uri,
            local,
            attribute_name,
            value,
        )?
    } else {
        rewrite_first_namespaced_attribute(
            &selected.xml,
            &selected.parent_namespaces,
            namespace_uri,
            local,
            attribute_name,
            value,
        )?
    };
    let mut output = Vec::with_capacity(xml.len() + replacement.len() - selected.range.len());
    output.extend_from_slice(&xml[..selected.range.start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[selected.range.end..]);
    Ok(output)
}

fn rewrite_selected_effect_attribute(
    xml: &[u8],
    inherited: &NamespaceBindings,
    namespace_uri: &str,
    local: &[u8],
    attribute_name: &str,
    value: &str,
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                let tag_range = start_tag_range(xml, reader.buffer_position() as usize)?;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(namespace_uri)
                    && local_name(child.name().as_ref()) == local
                {
                    return replace_tag_attribute(xml, tag_range, attribute_name, value);
                }
                if scope.element_uri(child.name().as_ref()) == Some(MC_NS)
                    && local_name(child.name().as_ref()) == b"AlternateContent"
                {
                    let selected = selected_effect_fragment(&raw, parent)?.ok_or_else(|| {
                        OxmlError::MissingElement("selected transition effect".to_owned())
                    })?;
                    let replacement = rewrite_first_namespaced_attribute(
                        &selected.xml,
                        &selected.parent_namespaces,
                        namespace_uri,
                        local,
                        attribute_name,
                        value,
                    )?;
                    let range = tag_range.start + selected.range.start
                        ..tag_range.start + selected.range.end;
                    let mut output = Vec::with_capacity(
                        xml.len() + replacement.len().saturating_sub(range.len()),
                    );
                    output.extend_from_slice(&xml[..range.start]);
                    output.extend_from_slice(&replacement);
                    output.extend_from_slice(&xml[range.end..]);
                    return Ok(output);
                }
            }
            Event::Empty(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(namespace_uri)
                    && local_name(child.name().as_ref()) == local
                {
                    return replace_tag_attribute(
                        xml,
                        start_tag_range(xml, reader.buffer_position() as usize)?,
                        attribute_name,
                        value,
                    );
                }
            }
            Event::End(_) | Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Err(OxmlError::MissingElement(format!(
        "{}:{}",
        namespace_uri,
        String::from_utf8_lossy(local)
    )))
}

fn root_is(
    xml: &[u8],
    inherited: &NamespaceBindings,
    namespace_uri: &str,
    expected: &[u8],
) -> Result<bool> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) | Event::Empty(start) => {
                let scope = inherited.with_start(&start)?;
                return Ok(
                    scope.element_uri(start.name().as_ref()) == Some(namespace_uri)
                        && local_name(start.name().as_ref()) == expected,
                );
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_transition_children(
    reader: &mut Reader<&[u8]>,
    namespaces: &NamespaceBindings,
) -> Result<(
    Option<TransitionEffect>,
    Vec<TransitionParameter>,
    Option<MorphMetadata>,
)> {
    let mut effect = None;
    let mut effect_parameters = Vec::new();
    let mut morph = None;
    let mut buffer = Vec::new();
    let mut seen_effect = false;
    let mut seen_sound_action = false;
    let mut seen_extension_list = false;
    let mut boundary = 0u8;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(child) => {
                let scope = namespaces.with_start(&child)?;
                let raw = capture_element(reader, &child)?;
                if let Some(projection) = transition_child_projection(&child, &scope, Some(&raw))? {
                    record_transition_child(
                        projection,
                        &mut boundary,
                        &mut seen_effect,
                        &mut seen_sound_action,
                        &mut seen_extension_list,
                        &mut effect,
                        &mut effect_parameters,
                        &mut morph,
                    )?;
                }
            }
            Event::Empty(child) => {
                let scope = namespaces.with_start(&child)?;
                if let Some(projection) = transition_child_projection(&child, &scope, None)? {
                    record_transition_child(
                        projection,
                        &mut boundary,
                        &mut seen_effect,
                        &mut seen_sound_action,
                        &mut seen_extension_list,
                        &mut effect,
                        &mut effect_parameters,
                        &mut morph,
                    )?;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"transition" => {
                return Ok((effect, effect_parameters, morph));
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("closing p:transition".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

enum TransitionChild {
    Effect(
        TransitionEffect,
        Vec<TransitionParameter>,
        Option<MorphMetadata>,
    ),
    UnsupportedEffect,
    SoundAction,
    ExtensionList,
}

fn transition_child_projection(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
    raw: Option<&[u8]>,
) -> Result<Option<TransitionChild>> {
    let local = local_name(start.name().as_ref()).to_vec();
    match namespaces.element_uri(start.name().as_ref()) {
        Some(P_NS) if local == b"sndAc" => Ok(Some(TransitionChild::SoundAction)),
        Some(P_NS) if local == b"extLst" => Ok(Some(TransitionChild::ExtensionList)),
        Some(P_NS) => {
            let attributes = all_attributes(start)?;
            Ok(Some(TransitionChild::Effect(
                parse_transition_effect(&local),
                transition_parameters(&attributes),
                None,
            )))
        }
        Some(P159_NS) if local == b"morph" => {
            let attributes = all_attributes(start)?;
            let metadata = MorphMetadata {
                option: attribute(&attributes, "option").map(str::to_owned),
            };
            Ok(Some(TransitionChild::Effect(
                TransitionEffect::Morph,
                Vec::new(),
                Some(metadata),
            )))
        }
        Some(MC_NS) if local == b"AlternateContent" => match raw {
            Some(raw) => selected_effect_from_alternate(raw, namespaces).map(Some),
            None => Ok(Some(TransitionChild::UnsupportedEffect)),
        },
        _ => Ok(None),
    }
}

#[allow(clippy::too_many_arguments)]
fn record_transition_child(
    child: TransitionChild,
    boundary: &mut u8,
    seen_effect: &mut bool,
    seen_sound_action: &mut bool,
    seen_extension_list: &mut bool,
    effect: &mut Option<TransitionEffect>,
    effect_parameters: &mut Vec<TransitionParameter>,
    morph: &mut Option<MorphMetadata>,
) -> Result<()> {
    let (rank, seen, name) = match &child {
        TransitionChild::Effect(..) | TransitionChild::UnsupportedEffect => {
            (1, seen_effect, "transition effect")
        }
        TransitionChild::SoundAction => (2, seen_sound_action, "sndAc"),
        TransitionChild::ExtensionList => (3, seen_extension_list, "extLst"),
    };
    if *seen {
        return Err(duplicate(name));
    }
    if rank < *boundary {
        return Err(out_of_order(name));
    }
    *seen = true;
    *boundary = rank;
    if let TransitionChild::Effect(parsed_effect, parameters, metadata) = child {
        *effect = Some(parsed_effect);
        *effect_parameters = parameters;
        *morph = metadata;
    }
    Ok(())
}

fn selected_effect_from_alternate(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<TransitionChild> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut fallback = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(branch) => {
                let wrapper = root.as_ref().expect("root");
                let scope = wrapper.with_start(&branch)?;
                let is_mc = scope.element_uri(branch.name().as_ref()) == Some(MC_NS);
                let local = local_name(branch.name().as_ref()).to_vec();
                let raw = capture_element(&mut reader, &branch)?;
                if is_mc && local == b"Choice" && choice_is_supported(&branch, &scope)? {
                    return effect_from_branch(&raw, wrapper);
                }
                if is_mc && local == b"Fallback" && fallback.is_none() {
                    fallback = Some((raw, wrapper.clone()));
                }
            }
            Event::Empty(branch) => {
                let wrapper = root.as_ref().expect("root");
                let scope = wrapper.with_start(&branch)?;
                if scope.element_uri(branch.name().as_ref()) == Some(MC_NS)
                    && local_name(branch.name().as_ref()) == b"Choice"
                    && choice_is_supported(&branch, &scope)?
                {
                    return Ok(TransitionChild::UnsupportedEffect);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"AlternateContent" => {
                return match fallback {
                    Some((raw, wrapper)) => effect_from_branch(&raw, &wrapper),
                    None => Ok(TransitionChild::UnsupportedEffect),
                };
            }
            Event::Eof => return Ok(TransitionChild::UnsupportedEffect),
            _ => {}
        }
        buffer.clear();
    }
}

fn selected_effect_fragment(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<SelectedFragment>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut fallback = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(branch) => {
                let wrapper = root.as_ref().expect("root");
                let scope = wrapper.with_start(&branch)?;
                let is_mc = scope.element_uri(branch.name().as_ref()) == Some(MC_NS);
                let local = local_name(branch.name().as_ref()).to_vec();
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &branch)?;
                let range = start..reader.buffer_position() as usize;
                if is_mc && local == b"Choice" && choice_is_supported(&branch, &scope)? {
                    return direct_effect_fragment(&raw, wrapper)
                        .map(|selected| selected.map(|selected| offset_fragment(selected, range)));
                }
                if is_mc && local == b"Fallback" && fallback.is_none() {
                    fallback = Some((raw, wrapper.clone(), range));
                }
            }
            Event::Empty(branch) => {
                let wrapper = root.as_ref().expect("root");
                let scope = wrapper.with_start(&branch)?;
                if scope.element_uri(branch.name().as_ref()) == Some(MC_NS)
                    && local_name(branch.name().as_ref()) == b"Choice"
                    && choice_is_supported(&branch, &scope)?
                {
                    return Ok(None);
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"AlternateContent" => {
                let Some((raw, wrapper, range)) = fallback else {
                    return Ok(None);
                };
                return direct_effect_fragment(&raw, &wrapper)
                    .map(|selected| selected.map(|selected| offset_fragment(selected, range)));
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn direct_effect_fragment(
    xml: &[u8],
    inherited: &NamespaceBindings,
) -> Result<Option<SelectedFragment>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut selected = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(effect) => {
                let parent = root.as_ref().expect("root").clone();
                let scope = parent.with_start(&effect)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &effect)?;
                if direct_effect_projection(&effect, &scope)?.is_some() {
                    set_once(
                        &mut selected,
                        SelectedFragment {
                            xml: raw,
                            parent_namespaces: parent,
                            range: start..reader.buffer_position() as usize,
                        },
                        "effect in selected compatibility branch",
                    )?;
                }
            }
            Event::Empty(effect) => {
                let parent = root.as_ref().expect("root").clone();
                let scope = parent.with_start(&effect)?;
                if direct_effect_projection(&effect, &scope)?.is_some() {
                    let range = start_tag_range(xml, reader.buffer_position() as usize)?;
                    set_once(
                        &mut selected,
                        SelectedFragment {
                            xml: capture_empty_element(&effect)?,
                            parent_namespaces: parent,
                            range,
                        },
                        "effect in selected compatibility branch",
                    )?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(selected),
            _ => {}
        }
        buffer.clear();
    }
}

fn effect_from_branch(xml: &[u8], inherited: &NamespaceBindings) -> Result<TransitionChild> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut selected = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(effect) => {
                let scope = root.as_ref().expect("root").with_start(&effect)?;
                let projected = direct_effect_projection(&effect, &scope)?;
                let _ = capture_element(&mut reader, &effect)?;
                if let Some(projected) = projected {
                    set_once(
                        &mut selected,
                        projected,
                        "effect in selected compatibility branch",
                    )?;
                }
            }
            Event::Empty(effect) => {
                let scope = root.as_ref().expect("root").with_start(&effect)?;
                if let Some(projected) = direct_effect_projection(&effect, &scope)? {
                    set_once(
                        &mut selected,
                        projected,
                        "effect in selected compatibility branch",
                    )?;
                }
            }
            Event::End(_) | Event::Eof => {
                return Ok(selected.unwrap_or(TransitionChild::UnsupportedEffect));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn direct_effect_projection(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
) -> Result<Option<TransitionChild>> {
    let local = local_name(start.name().as_ref()).to_vec();
    if namespaces.element_uri(start.name().as_ref()) == Some(P_NS)
        && !matches!(local.as_slice(), b"sndAc" | b"extLst")
    {
        let attributes = all_attributes(start)?;
        return Ok(Some(TransitionChild::Effect(
            parse_transition_effect(&local),
            transition_parameters(&attributes),
            None,
        )));
    }
    if namespaces.element_uri(start.name().as_ref()) == Some(P159_NS) && local == b"morph" {
        let attributes = all_attributes(start)?;
        return Ok(Some(TransitionChild::Effect(
            TransitionEffect::Morph,
            Vec::new(),
            Some(MorphMetadata {
                option: attribute(&attributes, "option").map(str::to_owned),
            }),
        )));
    }
    Ok(None)
}

fn transition_parameters(attributes: &[(String, String)]) -> Vec<TransitionParameter> {
    attributes
        .iter()
        .filter(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
        .map(|(name, value)| TransitionParameter {
            name: name.clone(),
            value: value.clone(),
        })
        .collect()
}

fn parse_attribute_name(xml: &[u8], inherited: &NamespaceBindings) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut value = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                let is_attribute_name = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"attrName";
                let raw = capture_element(&mut reader, &child)?;
                if is_attribute_name {
                    set_once(&mut value, direct_text(&raw)?, "attrName")?;
                }
            }
            Event::Empty(child) => {
                let scope = root.as_ref().expect("root").with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"attrName"
                {
                    set_once(&mut value, None, "attrName")?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(value.flatten()),
            _ => {}
        }
        buffer.clear();
    }
}

fn direct_text(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut saw_root = false;
    let mut text_value = String::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if !saw_root => {
                saw_root = true;
                let _ = start;
            }
            Event::Start(child) => {
                let _ = capture_element(&mut reader, &child)?;
            }
            Event::Text(text) => text_value.push_str(
                &text
                    .decode()
                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
            ),
            Event::CData(text) => text_value.push_str(
                &text
                    .decode()
                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
            ),
            Event::GeneralRef(reference) => text_value.push_str(&resolve_entity(&reference)),
            Event::End(_) | Event::Eof => {
                let trimmed = text_value.trim();
                return Ok((!trimmed.is_empty()).then(|| trimmed.to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn rewrite_modeled_node_duration(
    xml: &[u8],
    inherited: &NamespaceBindings,
    id: u32,
    value: &str,
) -> Result<Vec<u8>> {
    let mut matches = Vec::new();
    collect_timing_ranges(xml, inherited, 0, id, &mut matches)?;
    if matches.len() != 1 {
        return Err(OxmlError::InvalidValue(format!(
            "timing node id {id} matched {} nodes",
            matches.len()
        )));
    }
    replace_tag_attribute(xml, matches.pop().expect("one range"), "dur", value)
}

fn collect_timing_ranges(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    id: u32,
    matches: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"tnLst"
                {
                    collect_node_list_ranges(&raw, parent, base + start, id, matches)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            Event::Empty(_) => {}
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_node_list_ranges(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    id: u32,
    matches: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                let modeled = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && is_modeled_timing_node(local_name(child.name().as_ref()));
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if modeled {
                    collect_modeled_node_ranges(&raw, parent, base + start, id, matches)?;
                }
            }
            Event::Eof => break,
            Event::End(_) => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn is_modeled_timing_node(local: &[u8]) -> bool {
    matches!(
        local,
        b"par" | b"seq" | b"set" | b"anim" | b"animEffect" | b"animMotion"
    )
}

fn collect_modeled_node_ranges(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    id: u32,
    matches: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut container = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let name = start.name();
                let local = local_name(name.as_ref());
                container = matches!(local, b"par" | b"seq");
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                let name = child.name();
                let local = local_name(name.as_ref());
                let owns_common = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && ((container && local == b"cTn") || (!container && local == b"cBhvr"));
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if owns_common {
                    if container {
                        collect_common_time_node_ranges(&raw, parent, base + start, id, matches)?;
                    } else {
                        collect_behavior_ranges(&raw, parent, base + start, id, matches)?;
                    }
                }
            }
            Event::Empty(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                if container
                    && scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn"
                {
                    let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                    let raw = capture_empty_element(&child)?;
                    collect_common_time_node_ranges(&raw, parent, base + start, id, matches)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_behavior_ranges(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    id: u32,
    matches: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                root = Some(inherited.with_start(&start)?);
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                let is_common = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn";
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if is_common {
                    collect_common_time_node_ranges(&raw, parent, base + start, id, matches)?;
                }
            }
            Event::Empty(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                if scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"cTn"
                {
                    let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                    let raw = capture_empty_element(&child)?;
                    collect_common_time_node_ranges(&raw, parent, base + start, id, matches)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_common_time_node_ranges(
    xml: &[u8],
    inherited: &NamespaceBindings,
    base: usize,
    id: u32,
    matches: &mut Vec<std::ops::Range<usize>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut root = None;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if root.is_none() => {
                let scope = inherited.with_start(&start)?;
                if attribute(&all_attributes(&start)?, "id")
                    .and_then(|value| value.parse::<u32>().ok())
                    == Some(id)
                {
                    let range = start_tag_range(xml, reader.buffer_position() as usize)?;
                    matches.push(base + range.start..base + range.end);
                }
                root = Some(scope);
            }
            Event::Empty(start) if root.is_none() => {
                if attribute(&all_attributes(&start)?, "id")
                    .and_then(|value| value.parse::<u32>().ok())
                    == Some(id)
                {
                    let range = start_tag_range(xml, reader.buffer_position() as usize)?;
                    matches.push(base + range.start..base + range.end);
                }
                return Ok(());
            }
            Event::Start(child) => {
                let parent = root.as_ref().expect("root");
                let scope = parent.with_start(&child)?;
                let is_child_list = scope.element_uri(child.name().as_ref()) == Some(P_NS)
                    && local_name(child.name().as_ref()) == b"childTnLst";
                let start = start_tag_range(xml, reader.buffer_position() as usize)?.start;
                let raw = capture_element(&mut reader, &child)?;
                if is_child_list {
                    collect_node_list_ranges(&raw, parent, base + start, id, matches)?;
                }
            }
            Event::End(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn rewrite_first_namespaced_attribute(
    xml: &[u8],
    inherited: &NamespaceBindings,
    namespace_uri: &str,
    local: &[u8],
    attribute_name: &str,
    value: &str,
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut scopes = vec![inherited.clone()];
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let scope = scopes.last().expect("scope").with_start(&start)?;
                if scope.element_uri(start.name().as_ref()) == Some(namespace_uri)
                    && local_name(start.name().as_ref()) == local
                {
                    return replace_tag_attribute(
                        xml,
                        start_tag_range(xml, reader.buffer_position() as usize)?,
                        attribute_name,
                        value,
                    );
                }
                scopes.push(scope);
            }
            Event::Empty(start) => {
                let scope = scopes.last().expect("scope").with_start(&start)?;
                if scope.element_uri(start.name().as_ref()) == Some(namespace_uri)
                    && local_name(start.name().as_ref()) == local
                {
                    return replace_tag_attribute(
                        xml,
                        start_tag_range(xml, reader.buffer_position() as usize)?,
                        attribute_name,
                        value,
                    );
                }
            }
            Event::End(_) => {
                if scopes.len() > 1 {
                    scopes.pop();
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Err(OxmlError::MissingElement(format!(
        "{}:{}",
        namespace_uri,
        String::from_utf8_lossy(local)
    )))
}

fn start_tag_range(xml: &[u8], end: usize) -> Result<std::ops::Range<usize>> {
    let start = xml[..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .ok_or_else(|| OxmlError::InvalidValue("missing start-tag boundary".to_owned()))?;
    Ok(start..end)
}

fn replace_tag_attribute(
    xml: &[u8],
    range: std::ops::Range<usize>,
    attribute_name: &str,
    value: &str,
) -> Result<Vec<u8>> {
    let tag = &xml[range.clone()];
    let mut index = 1usize;
    while index < tag.len()
        && !tag[index].is_ascii_whitespace()
        && !matches!(tag[index], b'/' | b'>')
    {
        index += 1;
    }
    let mut insertion = None;
    while index < tag.len() {
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= tag.len() || matches!(tag[index], b'/' | b'>') {
            insertion = Some(index);
            break;
        }
        let name_start = index;
        while index < tag.len()
            && !tag[index].is_ascii_whitespace()
            && !matches!(tag[index], b'=' | b'/' | b'>')
        {
            index += 1;
        }
        let name_end = index;
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if tag.get(index) != Some(&b'=') {
            return Err(OxmlError::InvalidValue(
                "malformed start-tag attribute".to_owned(),
            ));
        }
        index += 1;
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *tag
            .get(index)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| OxmlError::InvalidValue("unquoted start-tag attribute".to_owned()))?;
        index += 1;
        let value_start = index;
        while index < tag.len() && tag[index] != quote {
            index += 1;
        }
        if index >= tag.len() {
            return Err(OxmlError::InvalidValue(
                "unterminated start-tag attribute".to_owned(),
            ));
        }
        if &tag[name_start..name_end] == attribute_name.as_bytes() {
            let mut output = Vec::with_capacity(xml.len() + value.len());
            output.extend_from_slice(&xml[..range.start + value_start]);
            output.extend_from_slice(escaped_attribute(value, quote).as_bytes());
            output.extend_from_slice(&xml[range.start + index..]);
            return Ok(output);
        }
        index += 1;
    }
    let insertion = insertion
        .ok_or_else(|| OxmlError::InvalidValue("missing start-tag terminator".to_owned()))?;
    let mut output = Vec::with_capacity(xml.len() + attribute_name.len() + value.len() + 4);
    output.extend_from_slice(&xml[..range.start + insertion]);
    output.extend_from_slice(b" ");
    output.extend_from_slice(attribute_name.as_bytes());
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(escaped_attribute(value, b'"').as_bytes());
    output.extend_from_slice(b"\"");
    output.extend_from_slice(&xml[range.start + insertion..]);
    Ok(output)
}

fn escaped_attribute(value: &str, quote: u8) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '"' if quote == b'"' => escaped.push_str("&quot;"),
            '\'' if quote == b'\'' => escaped.push_str("&apos;"),
            '\r' => escaped.push_str("&#xD;"),
            '\n' => escaped.push_str("&#xA;"),
            '\t' => escaped.push_str("&#x9;"),
            character => escaped.push(character),
        }
    }
    escaped
}

fn validate_element(
    start: &BytesStart<'_>,
    namespaces: &NamespaceBindings,
    expected: &[u8],
) -> Result<()> {
    if local_name(start.name().as_ref()) != expected
        || namespaces.element_uri(start.name().as_ref()) != Some(P_NS)
    {
        return Err(OxmlError::UnexpectedElement(
            String::from_utf8_lossy(start.name().as_ref()).into_owned(),
        ));
    }
    Ok(())
}

fn attribute<'a>(attributes: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|(attribute, _)| attribute == name)
        .map(|(_, value)| value.as_str())
}

fn namespaced_attribute<'a>(
    attributes: &'a [(String, String)],
    namespaces: &NamespaceBindings,
    namespace_uri: &str,
    name: &str,
) -> Option<&'a str> {
    attributes
        .iter()
        .find(|(attribute, _)| {
            local_name(attribute.as_bytes()) == name.as_bytes()
                && namespaces.attribute_uri(attribute.as_bytes()) == Some(namespace_uri)
        })
        .map(|(_, value)| value.as_str())
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(duplicate(name));
    }
    *slot = Some(value);
    Ok(())
}

fn parse_optional_u64(value: Option<&str>, name: &str) -> Result<Option<u64>> {
    value
        .map(|value| {
            value
                .parse::<u64>()
                .map_err(|_| OxmlError::InvalidValue(format!("invalid @{name} value {value}")))
        })
        .transpose()
}

fn parse_optional_u32(value: Option<&str>, name: &str) -> Result<Option<u32>> {
    value
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| OxmlError::InvalidValue(format!("invalid @{name} value {value}")))
        })
        .transpose()
}

fn parse_required_u32(value: Option<&str>, name: &str) -> Result<u32> {
    parse_optional_u32(value, name)?.ok_or_else(|| OxmlError::MissingElement(format!("@{name}")))
}

fn parse_optional_bool(value: Option<&str>, name: &str) -> Result<Option<bool>> {
    value
        .map(|value| match value {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(OxmlError::InvalidValue(format!(
                "invalid @{name} value {value}"
            ))),
        })
        .transpose()
}

fn parse_fill(value: &str) -> TimingFill {
    match value {
        "hold" => TimingFill::Hold,
        "remove" => TimingFill::Remove,
        "freeze" => TimingFill::Freeze,
        "transition" => TimingFill::Transition,
        value => TimingFill::Other(value.to_owned()),
    }
}

fn parse_restart(value: &str) -> TimingRestart {
    match value {
        "always" => TimingRestart::Always,
        "whenNotActive" => TimingRestart::WhenNotActive,
        "never" => TimingRestart::Never,
        value => TimingRestart::Other(value.to_owned()),
    }
}

fn parse_node_type(value: &str) -> TimingNodeType {
    match value {
        "tmRoot" => TimingNodeType::TimingRoot,
        "mainSeq" => TimingNodeType::MainSequence,
        "clickEffect" => TimingNodeType::ClickEffect,
        "withEffect" => TimingNodeType::WithEffect,
        "afterEffect" => TimingNodeType::AfterEffect,
        value => TimingNodeType::Other(value.to_owned()),
    }
}

fn parse_event(value: &str) -> TimingEvent {
    match value {
        "onBegin" => TimingEvent::OnBegin,
        "onEnd" => TimingEvent::OnEnd,
        "onClick" => TimingEvent::OnClick,
        "onNext" => TimingEvent::OnNext,
        "onPrev" => TimingEvent::OnPrevious,
        "onStopAudio" => TimingEvent::OnStopAudio,
        value => TimingEvent::Other(value.to_owned()),
    }
}

fn parse_speed(value: &str) -> TransitionSpeed {
    match value {
        "slow" => TransitionSpeed::Slow,
        "med" => TransitionSpeed::Medium,
        "fast" => TransitionSpeed::Fast,
        value => TransitionSpeed::Other(value.to_owned()),
    }
}

fn parse_transition_effect(value: &[u8]) -> TransitionEffect {
    match value {
        b"cut" => TransitionEffect::Cut,
        b"fade" => TransitionEffect::Fade,
        b"wipe" => TransitionEffect::Wipe,
        b"push" => TransitionEffect::Push,
        b"zoom" => TransitionEffect::Zoom,
        value => TransitionEffect::Other(String::from_utf8_lossy(value).into_owned()),
    }
}

fn unsupported(local: &[u8], xml: &[u8]) -> TimingNode {
    TimingNode::Unsupported(TimingUnsupported {
        local_name: String::from_utf8_lossy(local).into_owned(),
        raw_xml: xml.to_vec(),
    })
}

fn duplicate(name: &str) -> OxmlError {
    OxmlError::InvalidValue(format!("duplicate p:{name}"))
}

fn out_of_order(name: &str) -> OxmlError {
    OxmlError::InvalidValue(format!("p:{name} is out of schema order"))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}
