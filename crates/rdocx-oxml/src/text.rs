//! Text content elements: `CT_P` (paragraph), `CT_R` (run), `CT_Text`.

use std::sync::atomic::{AtomicU64, Ordering};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::{NsReader, Reader, Writer, XmlVersion};

use crate::content_control::CT_Sdt;
use crate::drawing::CT_Drawing;
use crate::error::{OxmlError, Result};
use crate::math::OfficeMath;
use crate::namespace::{R_NS, matches_local_name};
use crate::numbering::{namespace_bindings, parse_scoped_ppr, word_prefixes_at};
use crate::properties::{CT_PPr, CT_RPr, is_word_attribute, is_word_element};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::CT_Revision;

static NEXT_FIELD_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

/// `CT_Text` — The text content of a run, with optional xml:space="preserve".
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Text {
    pub text: String,
    pub preserve_space: bool,
}

impl CT_Text {
    pub fn new(text: &str) -> Self {
        CT_Text {
            text: text.to_string(),
            preserve_space: text.starts_with(' ') || text.ends_with(' '),
        }
    }
}

/// A parsed Word field with its stored result and update marker.
#[derive(Debug, Clone)]
pub struct Field {
    pub instruction: FieldInstruction,
    pub cached_result: String,
    pub dirty: Option<bool>,
    nested_order: Vec<NestedFieldPosition>,
    source: FieldSource,
}

#[derive(Debug, Clone, Copy)]
enum NestedFieldPosition {
    Argument(usize),
    Switch(usize),
}

impl Field {
    /// Construct a field that will be written in the simple field form.
    pub fn new(instruction: &str, cached_result: &str) -> Self {
        let instruction = parse_field_instruction(instruction);
        Self {
            source: FieldSource::New {
                original_instruction: instruction.clone(),
            },
            instruction,
            cached_result: cached_result.to_owned(),
            dirty: None,
            nested_order: Vec::new(),
        }
    }

    fn parsed(
        instruction: FieldInstruction,
        cached_result: String,
        cached_segments: Vec<CachedDisplaySegment>,
        dirty: Option<bool>,
        form: FieldForm,
        raw_xml: Vec<u8>,
        word_prefixes: Vec<String>,
    ) -> Self {
        let original_instruction = instruction.clone();
        let original_cached_result = cached_result.clone();
        Self {
            instruction,
            cached_result,
            dirty,
            nested_order: Vec::new(),
            source: FieldSource::Parsed {
                source_id: NEXT_FIELD_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
                form,
                raw_xml,
                original_instruction,
                original_cached_result,
                cached_segments,
                original_dirty: dirty,
                word_prefixes,
            },
        }
    }

    /// Return nested fields in their original instruction order.
    #[doc(hidden)]
    pub fn nested_fields_in_source_order(&self) -> Vec<&Field> {
        let original = self.original_instruction();
        let structured_unchanged = instruction_structure_eq(&self.instruction, original);
        if structured_unchanged && self.instruction.raw != original.raw {
            return Vec::new();
        }
        self.nested_fields_from_instruction(&self.instruction, structured_unchanged)
    }

    /// Return the instruction selected by the public raw-versus-structured edit rules.
    #[doc(hidden)]
    pub fn effective_instruction(&self) -> FieldInstruction {
        instruction_for_write(self)
    }

    /// Return nested fields from an effective instruction in evaluation order.
    #[doc(hidden)]
    pub fn effective_nested_fields_in_source_order<'a>(
        &self,
        instruction: &'a FieldInstruction,
    ) -> Vec<&'a Field> {
        self.nested_fields_from_instruction(
            instruction,
            instruction_structure_eq(&self.instruction, self.original_instruction()),
        )
    }

    fn original_instruction(&self) -> &FieldInstruction {
        match &self.source {
            FieldSource::New {
                original_instruction,
            }
            | FieldSource::Parsed {
                original_instruction,
                ..
            } => original_instruction,
        }
    }

    fn nested_fields_from_instruction<'a>(
        &self,
        instruction: &'a FieldInstruction,
        preserve_source_order: bool,
    ) -> Vec<&'a Field> {
        let mut fields = Vec::new();
        if preserve_source_order {
            for position in &self.nested_order {
                let argument = match *position {
                    NestedFieldPosition::Argument(index) => instruction.arguments.get(index),
                    NestedFieldPosition::Switch(index) => instruction
                        .switches
                        .get(index)
                        .and_then(|switch| switch.argument.as_ref()),
                };
                if let Some(FieldArgument::Nested(field)) = argument {
                    fields.push(field.as_ref());
                }
            }
        }
        for argument in instruction.arguments.iter().chain(
            instruction
                .switches
                .iter()
                .filter_map(|switch| switch.argument.as_ref()),
        ) {
            let FieldArgument::Nested(field) = argument else {
                continue;
            };
            if !fields
                .iter()
                .any(|ordered| std::ptr::eq(*ordered, field.as_ref()))
            {
                fields.push(field.as_ref());
            }
        }
        fields
    }

    fn is_unchanged(&self) -> bool {
        match &self.source {
            FieldSource::New { .. } => false,
            FieldSource::Parsed {
                original_instruction,
                original_cached_result,
                original_dirty,
                ..
            } => {
                self.instruction == *original_instruction
                    && self.cached_result == *original_cached_result
                    && self.dirty == *original_dirty
            }
        }
    }

    fn is_parsed_complex(&self) -> bool {
        matches!(
            self.source,
            FieldSource::Parsed {
                form: FieldForm::Complex,
                ..
            }
        )
    }

    /// Whether this field was parsed from a complex `w:fldChar` sequence.
    #[doc(hidden)]
    pub fn is_complex(&self) -> bool {
        self.is_parsed_complex()
    }

    /// Return the text this field contributes to [`CT_R::text`].
    #[doc(hidden)]
    pub fn projected_text(&self) -> Option<&str> {
        self.is_parsed_complex()
            .then_some(self.cached_result.as_str())
    }

    /// Whether the retained field source carries semantic attributes outside
    /// the modeled instruction, type, and dirty-state projection.
    pub fn has_unmodeled_semantic_attributes(&self) -> bool {
        let FieldSource::Parsed {
            form,
            raw_xml,
            word_prefixes,
            ..
        } = &self.source
        else {
            return false;
        };

        field_source_has_unmodeled_semantic_attributes(raw_xml, *form, word_prefixes)
            .unwrap_or(true)
    }

    /// Return the stored display split by its original result-run formatting.
    #[doc(hidden)]
    pub fn cached_display_segments(&self) -> Vec<(&str, Option<&CT_RPr>)> {
        if let FieldSource::Parsed {
            original_cached_result,
            cached_segments,
            ..
        } = &self.source
            && !cached_segments.is_empty()
        {
            if self.cached_result == *original_cached_result {
                return cached_segments
                    .iter()
                    .map(|segment| (segment.text.as_str(), segment.properties.as_ref()))
                    .collect();
            }
            return vec![(
                self.cached_result.as_str(),
                cached_segments
                    .first()
                    .and_then(|segment| segment.properties.as_ref()),
            )];
        }
        vec![(self.cached_result.as_str(), None)]
    }

    /// Return the parsed source fragment and its cache-aware replacement.
    #[doc(hidden)]
    pub fn source_replacement(&self) -> Result<Option<(&[u8], Vec<u8>)>> {
        let FieldSource::Parsed { raw_xml, .. } = &self.source else {
            return Ok(None);
        };
        let mut writer = Writer::new(Vec::new());
        write_field(&mut writer, self, None)?;
        Ok(Some((raw_xml, writer.into_inner())))
    }
}

impl PartialEq for Field {
    fn eq(&self, other: &Self) -> bool {
        self.instruction == other.instruction
            && self.cached_result == other.cached_result
            && self.dirty == other.dirty
    }
}

/// The shared grammar for simple and complex field instructions.
#[derive(Debug, Clone, PartialEq)]
pub struct FieldInstruction {
    pub raw: String,
    pub name: String,
    pub arguments: Vec<FieldArgument>,
    pub switches: Vec<FieldSwitch>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FieldArgument {
    Text(String),
    Nested(Box<Field>),
}

#[derive(Debug, Clone, PartialEq)]
pub struct FieldSwitch {
    pub name: String,
    pub argument: Option<FieldArgument>,
}

#[derive(Debug, Clone)]
enum FieldSource {
    New {
        original_instruction: FieldInstruction,
    },
    Parsed {
        source_id: u64,
        form: FieldForm,
        raw_xml: Vec<u8>,
        original_instruction: FieldInstruction,
        original_cached_result: String,
        cached_segments: Vec<CachedDisplaySegment>,
        original_dirty: Option<bool>,
        word_prefixes: Vec<String>,
    },
}

#[derive(Debug, Clone)]
struct CachedDisplaySegment {
    run_index: usize,
    text: String,
    properties: Option<CT_RPr>,
}

#[derive(Debug, Clone, Copy)]
enum FieldForm {
    Simple,
    Complex,
}

/// Content that can appear inside a run.
#[derive(Debug, Clone, PartialEq)]
pub enum RunContent {
    Text(CT_Text),
    /// Deleted text projected from `w:delText` inside a deletion wrapper.
    DeletedText(CT_Text),
    Tab,
    Break(BreakType),
    Drawing(CT_Drawing),
    /// A simple or complex Word field.
    Field(Field),
    /// A footnote reference (`<w:footnoteReference w:id="..."/>`).
    FootnoteRef {
        id: i32,
    },
    /// An endnote reference (`<w:endnoteReference w:id="..."/>`).
    EndnoteRef {
        id: i32,
    },
    /// A comment reference (`<w:commentReference w:id="..."/>`).
    CommentReference {
        id: i32,
        /// Number of raw run children that precede this reference.
        raw_before: usize,
    },
}

/// A typed comment range boundary at a run insertion point.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CommentRangeMarker {
    Start {
        id: i32,
        run_index: usize,
        /// Number of raw children at this run boundary that precede the marker.
        raw_before: usize,
    },
    End {
        id: i32,
        run_index: usize,
        /// Number of raw children at this run boundary that precede the marker.
        raw_before: usize,
    },
}

/// Read projection of a bookmark marker retained at a run boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BookmarkMarker {
    start: bool,
    id: Option<i32>,
    name: Option<String>,
    run_index: usize,
    raw_before: usize,
}

impl BookmarkMarker {
    fn new(
        start: bool,
        id: i32,
        name: Option<String>,
        run_index: usize,
        raw_before: usize,
    ) -> Self {
        Self {
            start,
            id: Some(id),
            name,
            run_index,
            raw_before,
        }
    }

    pub fn is_start(&self) -> bool {
        self.start
    }

    pub fn id(&self) -> Option<i32> {
        self.id
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn run_index(&self) -> usize {
        self.run_index
    }

    /// Number of preserved raw children before this marker at its run boundary.
    #[doc(hidden)]
    pub fn raw_before(&self) -> usize {
        self.raw_before
    }
}

impl CommentRangeMarker {
    fn run_index(&self) -> usize {
        match self {
            Self::Start { run_index, .. } | Self::End { run_index, .. } => *run_index,
        }
    }

    fn raw_before(&self) -> usize {
        match self {
            Self::Start { raw_before, .. } | Self::End { raw_before, .. } => *raw_before,
        }
    }
}

/// Types of breaks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakType {
    Line,
    Page,
    Column,
}

/// `CT_R` — A run of text with uniform formatting.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_R {
    pub properties: Option<CT_RPr>,
    pub content: Vec<RunContent>,
    /// Unknown child elements captured as raw XML.
    pub extra_xml: Vec<Vec<u8>>,
    /// Encoded raw-child positions among properties and typed run content.
    #[doc(hidden)]
    pub extra_xml_positions: Vec<usize>,
    /// Drawings read out of an `mc:AlternateContent` block, for layout only.
    ///
    /// Never serialised. The verbatim copy in `extra_xml` is what gets
    /// written, so emitting these as well would duplicate the element.
    pub alt_drawings: Vec<CT_Drawing>,
}

const RAW_LEGACY_HORIZONTAL_RULE_FLAG: usize = 1usize << (usize::BITS - 1);
const RAW_CHILD_POSITION_MASK: usize = !RAW_LEGACY_HORIZONTAL_RULE_FLAG;
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";
const OFFICE_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:office:office";

fn resolved_name_matches(
    namespace: &ResolveResult<'_>,
    local_name: &[u8],
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    local_name == expected_local_name
        && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected_namespace)
}

fn rect_has_enabled_horizontal_rule(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> bool {
    let mut enabled = None;
    for attribute in element.attributes() {
        let Ok(attribute) = attribute else {
            return false;
        };
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_name_matches(&namespace, local_name.as_ref(), OFFICE_NAMESPACE, b"hr") {
            if enabled.is_some() {
                return false;
            }
            let Ok(value) =
                attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            else {
                return false;
            };
            enabled = Some(matches!(value.as_bytes(), b"t" | b"true"));
        }
    }
    enabled == Some(true)
}

fn is_legacy_horizontal_rule(raw_xml: &[u8], inherited_namespaces: &[(String, String)]) -> bool {
    let scoped_xml = if inherited_namespaces.is_empty() {
        None
    } else {
        let mut wrapper = BytesStart::new("rdocx-scope");
        let names = inherited_namespaces
            .iter()
            .map(|(prefix, _)| {
                if prefix.is_empty() {
                    "xmlns".to_owned()
                } else {
                    format!("xmlns:{prefix}")
                }
            })
            .collect::<Vec<_>>();
        for ((_, namespace), name) in inherited_namespaces.iter().zip(&names) {
            wrapper.push_attribute((name.as_str(), namespace.as_str()));
        }
        let mut writer = Writer::new(Vec::new());
        if writer.write_event(Event::Start(wrapper)).is_err() {
            return false;
        }
        writer.get_mut().extend_from_slice(raw_xml);
        if writer
            .write_event(Event::End(BytesEnd::new("rdocx-scope")))
            .is_err()
        {
            return false;
        }
        Some(writer.into_inner())
    };
    let mut reader = NsReader::from_reader(scoped_xml.as_deref().unwrap_or(raw_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut state = u8::from(scoped_xml.is_some()) * 4;
    let mut found_rect = false;

    loop {
        let Ok((namespace, event)) = reader.read_resolved_event_into(&mut buffer) else {
            return false;
        };
        match event {
            Event::Start(element) if state == 4 && element.name().as_ref() == b"rdocx-scope" => {
                state = 0;
            }
            Event::Start(element) if state == 0 => {
                if !resolved_name_matches(
                    &namespace,
                    element.local_name().as_ref(),
                    crate::namespace::W_NS.as_bytes(),
                    b"pict",
                ) {
                    return false;
                }
                state = 1;
            }
            Event::Start(element) if state == 1 && !found_rect => {
                if !resolved_name_matches(
                    &namespace,
                    element.local_name().as_ref(),
                    VML_NAMESPACE,
                    b"rect",
                ) || !rect_has_enabled_horizontal_rule(&reader, &element)
                {
                    return false;
                }
                found_rect = true;
                state = 2;
            }
            Event::Empty(element) if state == 1 && !found_rect => {
                if !resolved_name_matches(
                    &namespace,
                    element.local_name().as_ref(),
                    VML_NAMESPACE,
                    b"rect",
                ) || !rect_has_enabled_horizontal_rule(&reader, &element)
                {
                    return false;
                }
                found_rect = true;
            }
            Event::End(element) if state == 2 => {
                if !resolved_name_matches(
                    &namespace,
                    element.local_name().as_ref(),
                    VML_NAMESPACE,
                    b"rect",
                ) {
                    return false;
                }
                state = 1;
            }
            Event::End(element) if state == 1 && found_rect => {
                if !resolved_name_matches(
                    &namespace,
                    element.local_name().as_ref(),
                    crate::namespace::W_NS.as_bytes(),
                    b"pict",
                ) {
                    return false;
                }
                state = 3;
            }
            Event::End(element)
                if state == 3
                    && scoped_xml.is_some()
                    && element.name().as_ref() == b"rdocx-scope" =>
            {
                state = 4;
            }
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Event::Eof => {
                let finished = if scoped_xml.is_some() {
                    state == 4
                } else {
                    state == 3
                };
                return finished && found_rect;
            }
            _ => return false,
        }
        buffer.clear();
    }
}

#[allow(non_snake_case)]
impl CT_R {
    pub fn new(text: &str) -> Self {
        CT_R {
            properties: None,
            content: vec![RunContent::Text(CT_Text::new(text))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        }
    }

    /// Decode a raw-child boundary stored in `extra_xml_positions`.
    #[doc(hidden)]
    pub fn raw_child_position(encoded: usize) -> usize {
        encoded & RAW_CHILD_POSITION_MASK
    }

    /// Report whether a raw child was parsed as a legacy horizontal rule.
    #[doc(hidden)]
    pub fn raw_child_is_legacy_horizontal_rule(encoded: usize) -> bool {
        encoded & RAW_LEGACY_HORIZONTAL_RULE_FLAG != 0
    }

    /// Replace a raw-child boundary without changing its parsed classification.
    #[doc(hidden)]
    pub fn set_raw_child_position(encoded: &mut usize, position: usize) {
        debug_assert_eq!(position & RAW_LEGACY_HORIZONTAL_RULE_FLAG, 0);
        *encoded =
            (position & RAW_CHILD_POSITION_MASK) | (*encoded & RAW_LEGACY_HORIZONTAL_RULE_FLAG);
    }

    fn encode_raw_child_position(position: usize, legacy_horizontal_rule: bool) -> usize {
        debug_assert_eq!(position & RAW_LEGACY_HORIZONTAL_RULE_FLAG, 0);
        position
            | if legacy_horizontal_rule {
                RAW_LEGACY_HORIZONTAL_RULE_FLAG
            } else {
                0
            }
    }

    /// Get the combined text of all text content in this run.
    pub fn text(&self) -> String {
        let mut result = String::new();
        for item in &self.content {
            match item {
                RunContent::Text(t) | RunContent::DeletedText(t) => result.push_str(&t.text),
                RunContent::Tab => result.push('\t'),
                RunContent::Break(_) => result.push('\n'),
                RunContent::Drawing(_) => {} // Drawings have no text content
                RunContent::Field(field) => {
                    if let Some(text) = field.projected_text() {
                        result.push_str(text);
                    }
                }
                RunContent::FootnoteRef { .. }
                | RunContent::EndnoteRef { .. }
                | RunContent::CommentReference { .. } => {}
            }
        }
        result
    }

    /// Replace typed run content while retaining every raw child boundary.
    #[doc(hidden)]
    pub fn replace_content(&mut self, content: Vec<RunContent>) {
        let property_boundary = usize::from(self.properties.is_some());
        let replacement_end = property_boundary + content.len();
        for position in &mut self.extra_xml_positions {
            if Self::raw_child_position(*position) > property_boundary {
                Self::set_raw_child_position(position, replacement_end);
            }
        }
        self.content = content;
    }

    /// Append typed content without moving raw children after existing content.
    #[doc(hidden)]
    pub fn append_content(&mut self, content: RunContent) {
        self.content.push(content);
    }

    /// Materialize run properties in their required first-child position.
    #[doc(hidden)]
    pub fn ensure_properties(&mut self) -> &mut CT_RPr {
        if self.properties.is_none() {
            for position in &mut self.extra_xml_positions {
                let shifted = Self::raw_child_position(*position) + 1;
                Self::set_raw_child_position(position, shifted);
            }
            self.properties = Some(CT_RPr::default());
        }
        self.properties.as_mut().expect("properties were inserted")
    }

    /// Remap raw boundaries after selected typed content children are removed.
    #[doc(hidden)]
    pub fn remap_removed_content(&mut self, removed: &[bool]) {
        let property_boundary = usize::from(self.properties.is_some());
        for position in &mut self.extra_xml_positions {
            let decoded = Self::raw_child_position(*position);
            let content_boundary = decoded.saturating_sub(property_boundary);
            let remapped = decoded.saturating_sub(
                removed
                    .iter()
                    .take(content_boundary.min(removed.len()))
                    .filter(|remove| **remove)
                    .count(),
            );
            Self::set_raw_child_position(position, remapped);
        }
    }

    /// Remove selected comment-reference content and retain surrounding raw XML.
    #[doc(hidden)]
    pub fn remove_comment_references(&mut self, ids: &[i32]) -> bool {
        let removed = self
            .content
            .iter()
            .map(|content| {
                matches!(content, RunContent::CommentReference { id, .. } if ids.contains(id))
            })
            .collect::<Vec<_>>();
        let removed_any = removed.iter().any(|remove| *remove);
        if removed_any {
            self.remap_removed_content(&removed);
            self.content = self
                .content
                .drain(..)
                .zip(removed)
                .filter_map(|(content, remove)| (!remove).then_some(content))
                .collect();
        }
        removed_any
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(
            reader,
            &["w".to_owned(), format!("\0mc\0{}", crate::namespace::MC_NS)],
        )
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut properties = None;
        let mut content = Vec::new();
        let mut extra_xml = Vec::new();
        let mut extra_xml_positions = Vec::new();
        let mut alt_drawings = Vec::new();
        let mut modeled_children = 0usize;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"rPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        properties = Some(crate::numbering::parse_scoped_rpr(&raw, word_prefixes)?);
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"t", &prefixes) {
                        let preserve = e.attributes().any(|a| {
                            a.ok()
                                .map(|a| {
                                    a.key.as_ref() == b"xml:space"
                                        && a.value.as_ref() == b"preserve"
                                })
                                .unwrap_or(false)
                        });
                        // `read_text` returns the raw markup span, so entity
                        // references in it still need resolving.
                        let encoded = reader.read_text(name)?;
                        encoded
                            .decode()
                            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
                        let text = crate::xml_text::decode_escaped(&encoded);
                        content.push(RunContent::Text(CT_Text {
                            text,
                            preserve_space: preserve,
                        }));
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"delText", &prefixes) {
                        let preserve = e.attributes().any(|a| {
                            a.ok().is_some_and(|a| {
                                a.key.as_ref() == b"xml:space" && a.value.as_ref() == b"preserve"
                            })
                        });
                        let encoded = reader.read_text(name)?;
                        encoded
                            .decode()
                            .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
                        let text = crate::xml_text::decode_escaped(&encoded);
                        content.push(RunContent::DeletedText(CT_Text {
                            text,
                            preserve_space: preserve,
                        }));
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"drawing", &prefixes) {
                        content.push(RunContent::Drawing(CT_Drawing::from_xml_with_prefixes(
                            reader, &prefixes,
                        )?));
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"commentReference", &prefixes) {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        content.push(RunContent::CommentReference {
                            id,
                            raw_before: extra_xml.len(),
                        });
                        modeled_children += 1;
                    } else if is_element_in_namespace(
                        name.as_ref(),
                        b"AlternateContent",
                        crate::namespace::MC_NS,
                        &prefixes,
                    ) {
                        // Keep the block verbatim so the VML fallback survives
                        // a write, and separately read the DrawingML out of it
                        // so layout can see the shape. alt_drawings is never
                        // serialised, the raw copy below is what gets written.
                        let raw = capture_element(reader, e)?;
                        if let Some(drawing) = crate::drawing::parse_alternate_content(&raw) {
                            alt_drawings.push(drawing);
                        }
                        extra_xml.push(raw);
                        extra_xml_positions.push(modeled_children);
                    } else {
                        // Capture unknown child elements as raw XML
                        let is_word_pict = is_word_element(name.as_ref(), b"pict", &prefixes);
                        let raw = capture_element(reader, e)?;
                        let legacy_horizontal_rule = is_word_pict
                            && is_legacy_horizontal_rule(&raw, &namespace_bindings(word_prefixes));
                        extra_xml.push(raw);
                        extra_xml_positions.push(Self::encode_raw_child_position(
                            modeled_children,
                            legacy_horizontal_rule,
                        ));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"tab", &prefixes) {
                        content.push(RunContent::Tab);
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"br", &prefixes) {
                        let break_type = optional_word_attribute(e, b"type", &prefixes)
                            .map(|value| match value.as_bytes() {
                                b"page" => BreakType::Page,
                                b"column" => BreakType::Column,
                                _ => BreakType::Line,
                            })
                            .unwrap_or(BreakType::Line);
                        content.push(RunContent::Break(break_type));
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"footnoteReference", &prefixes) {
                        let id = optional_word_attribute(e, b"id", &prefixes)
                            .and_then(|value| value.parse::<i32>().ok())
                            .unwrap_or(0);
                        content.push(RunContent::FootnoteRef { id });
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"endnoteReference", &prefixes) {
                        let id = optional_word_attribute(e, b"id", &prefixes)
                            .and_then(|value| value.parse::<i32>().ok())
                            .unwrap_or(0);
                        content.push(RunContent::EndnoteRef { id });
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"commentReference", &prefixes) {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        content.push(RunContent::CommentReference {
                            id,
                            raw_before: extra_xml.len(),
                        });
                        modeled_children += 1;
                    } else if !is_word_element(name.as_ref(), b"rPr", &prefixes) {
                        // Capture unknown empty child elements (e.g.
                        // w:commentReference) as raw XML, mirroring the
                        // Event::Start fallback above.
                        //
                        // A self-closing <w:rPr/> is deliberately skipped.
                        // extra_xml is re-emitted after the run content, but
                        // CT_R requires w:rPr to be the first child, so
                        // capturing it here would move it past <w:t> and
                        // produce schema-invalid output. An empty rPr carries
                        // no formatting, so dropping it loses nothing.
                        extra_xml.push(capture_empty_element(e)?);
                        extra_xml_positions.push(modeled_children);
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"r") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_R {
            properties,
            content,
            extra_xml,
            extra_xml_positions,
            alt_drawings,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.to_xml_with_word_override(writer, None)
    }

    pub(crate) fn to_xml_with_word_override<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        foreign_word_namespace: Option<&str>,
    ) -> Result<()> {
        let mut run = BytesStart::new("w:r");
        if foreign_word_namespace.is_some() {
            run.push_attribute(("xmlns:w", crate::namespace::W_NS));
        }
        writer.write_event(Event::Start(run))?;

        let ordered_raw = self.extra_xml_positions.len() == self.extra_xml.len();
        let mut typed_boundary = 0usize;
        if ordered_raw {
            write_run_raw_boundary(writer, self, typed_boundary, foreign_word_namespace)?;
        }
        if let Some(ref props) = self.properties {
            props.to_xml_with_word_override(writer, foreign_word_namespace)?;
            typed_boundary += 1;
            if ordered_raw {
                write_run_raw_boundary(writer, self, typed_boundary, foreign_word_namespace)?;
            }
        }

        let mut raw_written = 0;
        for item in &self.content {
            match item {
                RunContent::Text(t) => {
                    let mut e = BytesStart::new("w:t");
                    if t.preserve_space {
                        e.push_attribute(("xml:space", "preserve"));
                    }
                    writer.write_event(Event::Start(e))?;
                    writer.write_event(Event::Text(BytesText::new(&t.text)))?;
                    writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                }
                RunContent::DeletedText(t) => {
                    let mut e = BytesStart::new("w:delText");
                    if t.preserve_space {
                        e.push_attribute(("xml:space", "preserve"));
                    }
                    writer.write_event(Event::Start(e))?;
                    writer.write_event(Event::Text(BytesText::new(&t.text)))?;
                    writer.write_event(Event::End(BytesEnd::new("w:delText")))?;
                }
                RunContent::Tab => {
                    writer.write_event(Event::Empty(BytesStart::new("w:tab")))?;
                }
                RunContent::Break(bt) => {
                    let mut e = BytesStart::new("w:br");
                    match bt {
                        BreakType::Page => e.push_attribute(("w:type", "page")),
                        BreakType::Column => e.push_attribute(("w:type", "column")),
                        BreakType::Line => {}
                    }
                    writer.write_event(Event::Empty(e))?;
                }
                RunContent::Drawing(d) => {
                    d.to_xml(writer)?;
                }
                RunContent::Field(_) => {
                    // Field runs are serialized at the paragraph level as <w:fldSimple>
                }
                RunContent::FootnoteRef { id } => {
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:footnoteReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    writer.write_event(Event::Empty(e))?;
                }
                RunContent::EndnoteRef { id } => {
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:endnoteReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    writer.write_event(Event::Empty(e))?;
                }
                RunContent::CommentReference { id, raw_before } => {
                    if !ordered_raw {
                        for raw in self
                            .extra_xml
                            .iter()
                            .take((*raw_before).min(self.extra_xml.len()))
                            .skip(raw_written)
                        {
                            write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
                            raw_written += 1;
                        }
                    }
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:commentReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    writer.write_event(Event::Empty(e))?;
                }
            }
            typed_boundary += 1;
            if ordered_raw {
                write_run_raw_boundary(writer, self, typed_boundary, foreign_word_namespace)?;
            }
        }

        // Write captured unknown child elements
        if !ordered_raw {
            for raw in self.extra_xml.iter().skip(raw_written) {
                write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        Ok(())
    }
}

fn write_run_raw_boundary<W: std::io::Write>(
    writer: &mut Writer<W>,
    run: &CT_R,
    boundary: usize,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    for (position, raw) in run.extra_xml_positions.iter().zip(&run.extra_xml) {
        if CT_R::raw_child_position(*position) == boundary {
            write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
        }
    }
    Ok(())
}

/// A hyperlink span that wraps a range of runs.
#[derive(Debug, Clone, PartialEq)]
pub struct HyperlinkSpan {
    /// The relationship ID for the hyperlink target.
    pub rel_id: Option<String>,
    /// Optional anchor within the document (for internal links).
    pub anchor: Option<String>,
    /// Optional user-facing hover text.
    pub tooltip: Option<String>,
    /// Optional location in the hyperlink target document.
    pub doc_location: Option<String>,
    /// Index of the first run in the hyperlink (inclusive).
    pub run_start: usize,
    /// Index of the last run in the hyperlink (exclusive).
    pub run_end: usize,
    /// Unmodeled owner attributes retained when the hyperlink came from XML.
    #[doc(hidden)]
    pub extra_attributes: Vec<(String, String)>,
    /// Raw children at `(relative run boundary, typed revisions before, XML)`.
    #[doc(hidden)]
    pub extra_xml: Vec<(usize, usize, Vec<u8>)>,
    /// Parent raw slot for a revision-only hyperlink preserved as one subtree.
    #[doc(hidden)]
    pub preserved_raw_before: Option<usize>,
}

const HYPERLINK_REVISION_FLAG: usize = 1usize << (usize::BITS - 1);

struct ParsedHyperlinkChildren {
    runs: Vec<CT_R>,
    run_sources: Vec<Option<Vec<u8>>>,
    revisions: Vec<(usize, CT_Revision)>,
    extra_xml: Vec<(usize, usize, Vec<u8>)>,
}

/// A hyperlink represented by a complex field sequence rather than `w:hyperlink`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComplexFieldHyperlink {
    pub run_start: usize,
    pub run_end: usize,
    pub target: String,
}

type ParsedHyperlinkAttributes = (
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Vec<(String, String)>,
);

pub(crate) fn hyperlink_revision_slot(hyperlink_index: usize) -> usize {
    HYPERLINK_REVISION_FLAG | hyperlink_index
}

#[doc(hidden)]
pub fn hyperlink_revision_index(slot: usize) -> Option<usize> {
    (slot & HYPERLINK_REVISION_FLAG != 0).then_some(slot & !HYPERLINK_REVISION_FLAG)
}

fn field_run(field: Field, properties: Option<CT_RPr>) -> CT_R {
    CT_R {
        properties,
        content: vec![RunContent::Field(field)],
        extra_xml: Vec::new(),
        extra_xml_positions: Vec::new(),
        alt_drawings: Vec::new(),
    }
}

fn parse_run_raw(raw: &[u8], word_prefixes: &[String]) -> Result<CT_R> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"r") => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                return CT_R::from_xml_with_prefixes(&mut reader, &prefixes);
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("w:r".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_simple_field(raw: &[u8], word_prefixes: &[String]) -> Result<Option<Field>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (instruction, dirty, prefixes) = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"fldSimple") => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                if !is_word_element(start.name().as_ref(), b"fldSimple", &prefixes) {
                    return Ok(None);
                }
                let Some(instruction) = optional_word_attribute(&start, b"instr", &prefixes) else {
                    return Ok(None);
                };
                let dirty = optional_word_attribute(&start, b"dirty", &prefixes)
                    .and_then(|value| parse_field_bool(&value));
                break (instruction, dirty, prefixes);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    };

    let mut cached_result = String::new();
    let mut cached_segments = Vec::new();
    let mut result_run_index = 0usize;
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let child_prefixes = word_prefixes_at(&start, &prefixes)?;
                if is_word_element(start.name().as_ref(), b"r", &child_prefixes) {
                    let run_raw = capture_element(&mut reader, &start)?;
                    let run = parse_run_raw(&run_raw, &child_prefixes)?;
                    if let Some(text) = simple_result_run_display(&run_raw, &child_prefixes)? {
                        cached_result.push_str(&text);
                        cached_segments.push(CachedDisplaySegment {
                            run_index: result_run_index,
                            text,
                            properties: run.properties,
                        });
                    }
                    result_run_index += 1;
                } else {
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                }
            }
            Event::End(end) if matches_local_name(end.name().as_ref(), b"fldSimple") => break,
            Event::Eof => break,
            _ => {}
        }
    }

    let instruction = parse_field_instruction(&instruction);
    if instruction.name.is_empty() {
        return Ok(None);
    }
    Ok(Some(Field::parsed(
        instruction,
        cached_result,
        cached_segments,
        dirty,
        FieldForm::Simple,
        raw.to_vec(),
        prefixes,
    )))
}

fn simple_result_run_display(raw: &[u8], word_prefixes: &[String]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let run_prefixes = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"r") => {
                break word_prefixes_at(&start, word_prefixes)?;
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    };
    let mut display = None::<String>;
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, &run_prefixes)?;
                if is_word_element(start.name().as_ref(), b"t", &prefixes)
                    || is_word_element(start.name().as_ref(), b"delText", &prefixes)
                {
                    let text = reader
                        .read_text(start.name())
                        .map(|text| crate::xml_text::decode_escaped(&text))
                        .unwrap_or_default();
                    display.get_or_insert_default().push_str(&text);
                } else if is_word_element(start.name().as_ref(), b"tab", &prefixes) {
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                    display.get_or_insert_default().push('\t');
                } else if is_word_element(start.name().as_ref(), b"br", &prefixes) {
                    let marker = match field_break_type(&start, &prefixes) {
                        BreakType::Line => '\n',
                        BreakType::Page => '\u{000c}',
                        BreakType::Column => '\u{000b}',
                    };
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                    display.get_or_insert_default().push(marker);
                } else {
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                }
            }
            Event::Empty(element) => {
                let prefixes = word_prefixes_at(&element, &run_prefixes)?;
                if is_word_element(element.name().as_ref(), b"t", &prefixes)
                    || is_word_element(element.name().as_ref(), b"delText", &prefixes)
                {
                    display.get_or_insert_default();
                } else if is_word_element(element.name().as_ref(), b"tab", &prefixes) {
                    display.get_or_insert_default().push('\t');
                } else if is_word_element(element.name().as_ref(), b"br", &prefixes) {
                    let marker = match field_break_type(&element, &prefixes) {
                        BreakType::Line => '\n',
                        BreakType::Page => '\u{000c}',
                        BreakType::Column => '\u{000b}',
                    };
                    display.get_or_insert_default().push(marker);
                }
            }
            Event::End(end) if matches_local_name(end.name().as_ref(), b"r") => {
                return Ok(display);
            }
            Event::Eof => return Ok(display),
            _ => {}
        }
    }
}

fn parse_field_bool(value: &str) -> Option<bool> {
    match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

#[derive(Debug)]
enum ComplexFieldEvent {
    Begin(Option<bool>),
    Separate(Option<bool>),
    End(Option<bool>),
    Instruction(String),
    Result(String),
    Tab,
    Break(BreakType),
}

fn complex_field_events(raw: &[u8], word_prefixes: &[String]) -> Result<Vec<ComplexFieldEvent>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut events = Vec::new();
    let mut buffer = Vec::new();
    let run_prefixes = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"r") => {
                break word_prefixes_at(&start, word_prefixes)?;
            }
            Event::Eof => return Ok(events),
            _ => {}
        }
        buffer.clear();
    };
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, &run_prefixes)?;
                if is_word_element(start.name().as_ref(), b"fldChar", &prefixes) {
                    push_field_char_event(&mut events, &start, &prefixes);
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                } else if is_word_element(start.name().as_ref(), b"instrText", &prefixes) {
                    events.push(ComplexFieldEvent::Instruction(
                        crate::xml_text::read_element_text(&mut reader, start.name()),
                    ));
                } else if is_word_element(start.name().as_ref(), b"t", &prefixes) {
                    events.push(ComplexFieldEvent::Result(
                        crate::xml_text::read_element_text(&mut reader, start.name()),
                    ));
                } else if is_word_element(start.name().as_ref(), b"tab", &prefixes) {
                    events.push(ComplexFieldEvent::Tab);
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                } else if is_word_element(start.name().as_ref(), b"br", &prefixes) {
                    events.push(ComplexFieldEvent::Break(field_break_type(
                        &start, &prefixes,
                    )));
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                } else {
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                }
            }
            Event::Empty(start) => {
                let prefixes = word_prefixes_at(&start, &run_prefixes)?;
                if is_word_element(start.name().as_ref(), b"fldChar", &prefixes) {
                    push_field_char_event(&mut events, &start, &prefixes);
                } else if is_word_element(start.name().as_ref(), b"tab", &prefixes) {
                    events.push(ComplexFieldEvent::Tab);
                } else if is_word_element(start.name().as_ref(), b"br", &prefixes) {
                    events.push(ComplexFieldEvent::Break(field_break_type(
                        &start, &prefixes,
                    )));
                }
            }
            Event::End(end) if matches_local_name(end.name().as_ref(), b"r") => break,
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(events)
}

fn field_break_type(element: &BytesStart<'_>, word_prefixes: &[String]) -> BreakType {
    optional_word_attribute(element, b"type", word_prefixes)
        .map(|value| match value.as_str() {
            "page" => BreakType::Page,
            "column" => BreakType::Column,
            _ => BreakType::Line,
        })
        .unwrap_or(BreakType::Line)
}

fn push_field_char_event(
    events: &mut Vec<ComplexFieldEvent>,
    element: &BytesStart<'_>,
    word_prefixes: &[String],
) {
    match optional_word_attribute(element, b"fldCharType", word_prefixes).as_deref() {
        Some("begin") => events.push(ComplexFieldEvent::Begin(
            optional_word_attribute(element, b"dirty", word_prefixes)
                .and_then(|value| parse_field_bool(&value)),
        )),
        Some("separate") => events.push(ComplexFieldEvent::Separate(
            optional_word_attribute(element, b"dirty", word_prefixes)
                .and_then(|value| parse_field_bool(&value)),
        )),
        Some("end") => events.push(ComplexFieldEvent::End(
            optional_word_attribute(element, b"dirty", word_prefixes)
                .and_then(|value| parse_field_bool(&value)),
        )),
        _ => {}
    }
}

struct ComplexFieldBuilder {
    start_run: usize,
    separate_run: Option<usize>,
    dirty: Option<bool>,
    instruction: Vec<InstructionPart>,
    cached_result: String,
    cached_segments: Vec<CachedDisplaySegment>,
    valid: bool,
}

struct ComplexFieldProjection<'a> {
    runs: &'a mut Vec<CT_R>,
    run_sources: &'a mut Vec<Option<Vec<u8>>>,
    extra_xml: &'a mut Vec<(usize, Vec<u8>)>,
    comment_ranges: &'a mut [CommentRangeMarker],
    bookmark_markers: &'a mut [BookmarkMarker],
    content_controls: &'a mut [(usize, usize, usize, CT_Sdt)],
    revisions: &'a mut [(usize, usize, CT_Revision)],
    hyperlinks: &'a mut [HyperlinkSpan],
    word_prefixes: &'a [String],
}

fn project_complex_fields(projection: ComplexFieldProjection<'_>) -> Result<()> {
    let ComplexFieldProjection {
        runs,
        run_sources,
        extra_xml,
        comment_ranges,
        bookmark_markers,
        content_controls,
        revisions,
        hyperlinks,
        word_prefixes,
    } = projection;
    if runs.len() != run_sources.len() {
        return Ok(());
    }
    let mut stack = Vec::<ComplexFieldBuilder>::new();
    let mut completed = Vec::<(usize, usize, Field, Option<CT_RPr>)>::new();
    for run_index in 0..runs.len() {
        let Some(raw) = run_sources[run_index].as_deref() else {
            continue;
        };
        for event in complex_field_events(raw, word_prefixes)? {
            match event {
                ComplexFieldEvent::Begin(dirty) => stack.push(ComplexFieldBuilder {
                    start_run: run_index,
                    separate_run: None,
                    dirty,
                    instruction: Vec::new(),
                    cached_result: String::new(),
                    cached_segments: Vec::new(),
                    valid: true,
                }),
                ComplexFieldEvent::Instruction(text) => {
                    if let Some(field) = stack.last_mut()
                        && field.separate_run.is_none()
                    {
                        field.instruction.push(InstructionPart::Text(text));
                    }
                }
                ComplexFieldEvent::Result(text) => {
                    for field in &mut stack {
                        if field.separate_run.is_some() {
                            field.cached_result.push_str(&text);
                            push_cached_display_segment(
                                &mut field.cached_segments,
                                run_index,
                                &text,
                                runs[run_index].properties.as_ref(),
                            );
                        }
                    }
                }
                ComplexFieldEvent::Tab => {
                    for field in &mut stack {
                        if field.separate_run.is_some() {
                            field.cached_result.push('\t');
                            push_cached_display_segment(
                                &mut field.cached_segments,
                                run_index,
                                "\t",
                                runs[run_index].properties.as_ref(),
                            );
                        }
                    }
                }
                ComplexFieldEvent::Break(break_type) => {
                    let marker = match break_type {
                        BreakType::Line => '\n',
                        BreakType::Page => '\u{000c}',
                        BreakType::Column => '\u{000b}',
                    };
                    for field in &mut stack {
                        if field.separate_run.is_some() {
                            field.cached_result.push(marker);
                            let mut text = String::new();
                            text.push(marker);
                            push_cached_display_segment(
                                &mut field.cached_segments,
                                run_index,
                                &text,
                                runs[run_index].properties.as_ref(),
                            );
                        }
                    }
                }
                ComplexFieldEvent::Separate(dirty) => {
                    if let Some(field) = stack.last_mut() {
                        merge_field_dirty(&mut field.dirty, dirty);
                        if field.separate_run.replace(run_index).is_some() {
                            field.valid = false;
                        }
                    }
                }
                ComplexFieldEvent::End(dirty) => {
                    let Some(mut field) = stack.pop() else {
                        continue;
                    };
                    merge_field_dirty(&mut field.dirty, dirty);
                    let (instruction, nested_order) =
                        parse_field_instruction_parts_with_order(field.instruction);
                    let source = complex_field_source(
                        field.start_run,
                        run_index,
                        run_sources,
                        extra_xml,
                        hyperlinks,
                    );
                    let mut parsed = Field::parsed(
                        instruction,
                        field.cached_result,
                        field.cached_segments,
                        field.dirty,
                        FieldForm::Complex,
                        source,
                        word_prefixes.to_vec(),
                    );
                    parsed.nested_order = nested_order;
                    let valid = field.valid
                        && !parsed.instruction.name.is_empty()
                        && run_sources[field.start_run..=run_index]
                            .iter()
                            .all(Option::is_some);
                    if let Some(parent) = stack.last_mut() {
                        if !valid {
                            parent.valid = false;
                        } else if parent.separate_run.is_none() {
                            parent.instruction.push(InstructionPart::Nested(parsed));
                        }
                    } else if valid {
                        completed.push((field.start_run, run_index, parsed, None));
                    }
                }
            }
        }
    }

    for (start, end, mut field, properties) in completed.into_iter().rev() {
        if has_typed_boundary_inside(
            start,
            end,
            comment_ranges,
            bookmark_markers,
            content_controls,
            revisions,
            hyperlinks,
        ) {
            continue;
        }
        let raw_xml = complex_field_source(start, end, run_sources, extra_xml, hyperlinks);
        if let FieldSource::Parsed {
            raw_xml: source, ..
        } = &mut field.source
        {
            *source = raw_xml;
        }
        extra_xml.retain(|(at, _)| !(*at > start && *at <= end));
        for hyperlink in hyperlinks.iter_mut() {
            let hyperlink_start = hyperlink.run_start;
            hyperlink.extra_xml.retain(|(boundary, _, _)| {
                let absolute = hyperlink_start + *boundary;
                !(absolute > start && absolute <= end)
            });
        }

        runs.splice(start..=end, [field_run(field, properties)]);
        run_sources.splice(start..=end, [None]);
        remap_complex_field_boundaries(
            start,
            end,
            ComplexFieldBoundariesMut {
                extra_xml,
                comment_ranges,
                bookmark_markers,
                content_controls,
                revisions,
                hyperlinks,
            },
        );
    }
    Ok(())
}

fn merge_field_dirty(current: &mut Option<bool>, next: Option<bool>) {
    if next == Some(true) || current.is_none() {
        *current = next;
    }
}

fn push_cached_display_segment(
    segments: &mut Vec<CachedDisplaySegment>,
    run_index: usize,
    text: &str,
    properties: Option<&CT_RPr>,
) {
    if let Some(segment) = segments.last_mut()
        && segment.run_index == run_index
    {
        segment.text.push_str(text);
        return;
    }
    segments.push(CachedDisplaySegment {
        run_index,
        text: text.to_owned(),
        properties: properties.cloned(),
    });
}

fn complex_field_source(
    start: usize,
    end: usize,
    run_sources: &[Option<Vec<u8>>],
    extra_xml: &[(usize, Vec<u8>)],
    hyperlinks: &[HyperlinkSpan],
) -> Vec<u8> {
    let mut source = Vec::new();
    for (run_index, raw_source) in run_sources.iter().enumerate().take(end + 1).skip(start) {
        if let Some(raw) = raw_source.as_deref() {
            source.extend_from_slice(raw);
        }
        if run_index < end {
            for (_, raw) in extra_xml.iter().filter(|(at, _)| *at == run_index + 1) {
                source.extend_from_slice(raw);
            }
            for hyperlink in hyperlinks
                .iter()
                .filter(|hyperlink| hyperlink.run_start <= start && hyperlink.run_end > end)
            {
                let relative = run_index + 1 - hyperlink.run_start;
                for (_, _, raw) in hyperlink
                    .extra_xml
                    .iter()
                    .filter(|(boundary, _, _)| *boundary == relative)
                {
                    source.extend_from_slice(raw);
                }
            }
        }
    }
    source
}

fn has_typed_boundary_inside(
    start: usize,
    end: usize,
    comment_ranges: &[CommentRangeMarker],
    bookmark_markers: &[BookmarkMarker],
    content_controls: &[(usize, usize, usize, CT_Sdt)],
    revisions: &[(usize, usize, CT_Revision)],
    hyperlinks: &[HyperlinkSpan],
) -> bool {
    let inside = |at: usize| at > start && at <= end;
    comment_ranges
        .iter()
        .any(|marker| inside(marker.run_index()))
        || bookmark_markers
            .iter()
            .any(|marker| inside(marker.run_index))
        || content_controls.iter().any(|(at, _, _, _)| inside(*at))
        || revisions.iter().any(|(at, _, _)| inside(*at))
        || hyperlinks.iter().any(|hyperlink| {
            let overlaps = hyperlink.run_start <= end && hyperlink.run_end > start;
            let contains_field = hyperlink.run_start <= start && hyperlink.run_end > end;
            overlaps && !contains_field
        })
}

struct ComplexFieldBoundariesMut<'a> {
    extra_xml: &'a mut [(usize, Vec<u8>)],
    comment_ranges: &'a mut [CommentRangeMarker],
    bookmark_markers: &'a mut [BookmarkMarker],
    content_controls: &'a mut [(usize, usize, usize, CT_Sdt)],
    revisions: &'a mut [(usize, usize, CT_Revision)],
    hyperlinks: &'a mut [HyperlinkSpan],
}

fn remap_complex_field_boundaries(
    start: usize,
    end: usize,
    boundaries: ComplexFieldBoundariesMut<'_>,
) {
    let ComplexFieldBoundariesMut {
        extra_xml,
        comment_ranges,
        bookmark_markers,
        content_controls,
        revisions,
        hyperlinks,
    } = boundaries;
    let removed = end - start;
    let remap = |at: &mut usize| {
        if *at > end {
            *at -= removed;
        } else if *at > start {
            *at = start + 1;
        }
    };
    for (at, _) in extra_xml {
        remap(at);
    }
    for marker in comment_ranges {
        match marker {
            CommentRangeMarker::Start { run_index, .. }
            | CommentRangeMarker::End { run_index, .. } => remap(run_index),
        }
    }
    for marker in bookmark_markers {
        remap(&mut marker.run_index);
    }
    for (at, _, _, _) in content_controls {
        remap(at);
    }
    for (at, _, _) in revisions {
        remap(at);
    }
    for hyperlink in hyperlinks {
        let old_start = hyperlink.run_start;
        let old_boundaries = hyperlink
            .extra_xml
            .iter()
            .map(|(boundary, _, _)| old_start + *boundary)
            .collect::<Vec<_>>();
        remap(&mut hyperlink.run_start);
        remap(&mut hyperlink.run_end);
        for ((boundary, _, _), mut absolute) in hyperlink.extra_xml.iter_mut().zip(old_boundaries) {
            remap(&mut absolute);
            *boundary = absolute.saturating_sub(hyperlink.run_start);
        }
    }
}

/// `CT_P` — A paragraph element containing runs and properties.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_P {
    pub properties: Option<CT_PPr>,
    pub runs: Vec<CT_R>,
    /// Hyperlink spans referencing ranges of runs.
    pub hyperlinks: Vec<HyperlinkSpan>,
    /// Typed comment range boundaries at run insertion points.
    pub comment_ranges: Vec<CommentRangeMarker>,
    /// Typed projections of preserved bookmark markers.
    pub bookmark_markers: Vec<BookmarkMarker>,
    /// Unknown child elements captured as raw XML with their insertion position (run index).
    pub extra_xml: Vec<(usize, Vec<u8>)>,
    /// Typed run controls at
    /// `(run index, raw children before, comment markers before, control)`.
    pub content_controls: Vec<(usize, usize, usize, CT_Sdt)>,
    /// Read projections of revision wrappers retained at paragraph or hyperlink boundaries.
    pub revisions: Vec<(usize, usize, CT_Revision)>,
    /// Typed OfficeMath projections keyed by `(run boundary, raw child slot)`.
    pub equations: Vec<(usize, usize, OfficeMath)>,
}

#[allow(non_snake_case)]
impl CT_P {
    pub fn new() -> Self {
        CT_P {
            properties: None,
            runs: Vec::new(),
            hyperlinks: Vec::new(),
            comment_ranges: Vec::new(),
            bookmark_markers: Vec::new(),
            extra_xml: Vec::new(),
            content_controls: Vec::new(),
            revisions: Vec::new(),
            equations: Vec::new(),
        }
    }

    /// Get the combined text of all runs in this paragraph.
    pub fn text(&self) -> String {
        self.runs().iter().map(|run| run.text()).collect()
    }

    /// Return valid complex `HYPERLINK` fields projected into synthetic runs.
    pub fn complex_field_hyperlinks(&self) -> Vec<ComplexFieldHyperlink> {
        self.runs
            .iter()
            .enumerate()
            .filter_map(|(run_index, run)| {
                let [RunContent::Field(field)] = run.content.as_slice() else {
                    return None;
                };
                if !field.is_parsed_complex()
                    || field.dirty == Some(true)
                    || field.cached_result.is_empty()
                    || field.instruction.name != "HYPERLINK"
                {
                    return None;
                }
                let target = field.instruction.arguments.iter().find_map(|argument| {
                    let FieldArgument::Text(target) = argument else {
                        return None;
                    };
                    (!target.is_empty()).then(|| target.clone())
                })?;
                Some(ComplexFieldHyperlink {
                    run_start: run_index,
                    run_end: run_index + 1,
                    target,
                })
            })
            .collect()
    }

    /// Return direct and content-control-wrapped runs in document order.
    pub fn runs(&self) -> Vec<&CT_R> {
        let mut runs = Vec::new();
        self.collect_runs(&mut runs);
        runs
    }

    /// Add a run with the given text.
    pub fn add_run(&mut self, text: &str) -> &mut CT_R {
        self.runs.push(CT_R::new(text));
        self.runs.last_mut().unwrap()
    }

    /// Insert a direct run while keeping every paragraph boundary projection aligned.
    #[doc(hidden)]
    pub fn insert_unwrapped_run(&mut self, run_index: usize, run: CT_R) -> bool {
        if run_index > self.runs.len() {
            return false;
        }
        for marker in &mut self.comment_ranges {
            match marker {
                CommentRangeMarker::Start { run_index: at, .. }
                | CommentRangeMarker::End { run_index: at, .. }
                    if *at >= run_index =>
                {
                    *at += 1;
                }
                CommentRangeMarker::Start { .. } | CommentRangeMarker::End { .. } => {}
            }
        }
        for marker in &mut self.bookmark_markers {
            if marker.run_index >= run_index {
                marker.run_index += 1;
            }
        }
        for (at, _, _, _) in &mut self.content_controls {
            if *at >= run_index {
                *at += 1;
            }
        }
        for (at, _, _) in &mut self.revisions {
            if *at >= run_index {
                *at += 1;
            }
        }

        let mut hyperlinks = Vec::with_capacity(self.hyperlinks.len() + 1);
        let mut hyperlink_map = Vec::with_capacity(self.hyperlinks.len());
        for mut hyperlink in self.hyperlinks.drain(..) {
            if hyperlink.run_start >= run_index {
                hyperlink.run_start += 1;
                hyperlink.run_end += 1;
                hyperlink_map.push((hyperlinks.len(), None, false));
                hyperlinks.push(hyperlink);
            } else if hyperlink.run_end > run_index {
                let split_at = run_index - hyperlink.run_start;
                let suffix = HyperlinkSpan {
                    rel_id: hyperlink.rel_id.clone(),
                    anchor: hyperlink.anchor.clone(),
                    tooltip: hyperlink.tooltip.clone(),
                    doc_location: hyperlink.doc_location.clone(),
                    run_start: run_index + 1,
                    run_end: hyperlink.run_end + 1,
                    extra_attributes: hyperlink.extra_attributes.clone(),
                    extra_xml: hyperlink
                        .extra_xml
                        .iter()
                        .filter(|(boundary, _, _)| *boundary >= split_at)
                        .map(|(boundary, before, raw)| (boundary - split_at, *before, raw.clone()))
                        .collect(),
                    preserved_raw_before: None,
                };
                hyperlink
                    .extra_xml
                    .retain(|(boundary, _, _)| *boundary < split_at);
                hyperlink.run_end = run_index;
                let prefix_index = hyperlinks.len();
                hyperlinks.push(hyperlink);
                let suffix_index = hyperlinks.len();
                hyperlinks.push(suffix);
                hyperlink_map.push((prefix_index, Some(suffix_index), false));
            } else {
                hyperlink_map.push((hyperlinks.len(), None, hyperlink.run_end == run_index));
                hyperlinks.push(hyperlink);
            }
        }
        self.hyperlinks = hyperlinks;
        for (at, slot, _) in &mut self.revisions {
            let Some(old_index) = hyperlink_revision_index(*slot) else {
                continue;
            };
            let Some((prefix, suffix, ended_at_insertion)) = hyperlink_map.get(old_index) else {
                continue;
            };
            if *ended_at_insertion && *at == run_index + 1 {
                *at = run_index;
            }
            let new_index = suffix.filter(|_| *at > run_index).unwrap_or(*prefix);
            *slot = hyperlink_revision_slot(new_index);
        }
        for (position, _) in &mut self.extra_xml {
            if *position >= run_index {
                *position += 1;
            }
        }
        for (position, _, _) in &mut self.equations {
            if *position >= run_index {
                *position += 1;
            }
        }
        self.runs.insert(run_index, run);
        true
    }

    /// Remove selected comment anchors and remap every collapsed run boundary.
    #[doc(hidden)]
    pub fn remove_comment_anchors(&mut self, ids: &[i32]) {
        for (run_index, raw_before, markers_before, _) in &mut self.content_controls {
            let preceding_removed =
                self.comment_ranges
                    .iter()
                    .filter(|marker| {
                        marker.run_index() == *run_index && marker.raw_before() == *raw_before
                    })
                    .take(*markers_before)
                    .filter(|marker| match marker {
                        CommentRangeMarker::Start { id, .. }
                        | CommentRangeMarker::End { id, .. } => ids.contains(id),
                    })
                    .count();
            *markers_before = markers_before.saturating_sub(preceding_removed);
        }
        self.comment_ranges.retain(|marker| match marker {
            CommentRangeMarker::Start { id, .. } | CommentRangeMarker::End { id, .. } => {
                !ids.contains(id)
            }
        });

        let mut removed = Vec::with_capacity(self.runs.len());
        for run in &mut self.runs {
            let removed_reference = run.remove_comment_references(ids);
            removed.push(
                removed_reference
                    && run.content.is_empty()
                    && run.extra_xml.is_empty()
                    && run.alt_drawings.is_empty(),
            );
        }
        if removed.iter().all(|remove| !remove) {
            return;
        }

        let old_run_count = self.runs.len();
        let boundary_map = (0..=old_run_count)
            .map(|boundary| {
                boundary
                    - removed
                        .iter()
                        .take(boundary)
                        .filter(|remove| **remove)
                        .count()
            })
            .collect::<Vec<_>>();
        let mut raw_counts = vec![0usize; old_run_count + 1];
        for (position, _) in &self.extra_xml {
            raw_counts[(*position).min(old_run_count)] += 1;
        }
        let mut raw_prefixes = vec![0usize; old_run_count + 1];
        for boundary in 1..=old_run_count {
            if boundary_map[boundary] == boundary_map[boundary - 1] {
                raw_prefixes[boundary] = raw_prefixes[boundary - 1] + raw_counts[boundary - 1];
            }
        }

        self.extra_xml.sort_by_key(|(position, _)| *position);
        self.comment_ranges
            .sort_by_key(|marker| (marker.run_index(), marker.raw_before()));
        self.bookmark_markers.sort_by_key(|marker| marker.run_index);
        self.content_controls
            .sort_by_key(|(at, raw_before, markers_before, _)| (*at, *raw_before, *markers_before));

        for control_index in 0..self.content_controls.len() {
            let (run_index, raw_before, markers_before) = {
                let control = &self.content_controls[control_index];
                (control.0, control.1, control.2)
            };
            let old_boundary = run_index.min(old_run_count);
            let old_raw_slot = raw_before.min(raw_counts[old_boundary]);
            let new_boundary = boundary_map[old_boundary];
            let new_raw_slot = raw_prefixes[old_boundary] + old_raw_slot;
            let preceding_markers = self
                .comment_ranges
                .iter()
                .filter(|marker| {
                    let marker_boundary = marker.run_index().min(old_run_count);
                    marker_boundary < old_boundary
                        && boundary_map[marker_boundary] == new_boundary
                        && raw_prefixes[marker_boundary]
                            + marker.raw_before().min(raw_counts[marker_boundary])
                            == new_raw_slot
                })
                .count();
            let local_marker_count = self
                .comment_ranges
                .iter()
                .filter(|marker| {
                    marker.run_index() == old_boundary
                        && marker.raw_before().min(raw_counts[old_boundary]) == old_raw_slot
                })
                .count();
            let control = &mut self.content_controls[control_index];
            control.0 = new_boundary;
            control.1 = new_raw_slot;
            control.2 = preceding_markers + markers_before.min(local_marker_count);
        }
        for marker in &mut self.comment_ranges {
            match marker {
                CommentRangeMarker::Start {
                    run_index,
                    raw_before,
                    ..
                }
                | CommentRangeMarker::End {
                    run_index,
                    raw_before,
                    ..
                } => {
                    let old_boundary = (*run_index).min(old_run_count);
                    *run_index = boundary_map[old_boundary];
                    *raw_before =
                        raw_prefixes[old_boundary] + (*raw_before).min(raw_counts[old_boundary]);
                }
            }
        }
        for marker in &mut self.bookmark_markers {
            let old_boundary = marker.run_index.min(old_run_count);
            marker.run_index = boundary_map[old_boundary];
            marker.raw_before =
                raw_prefixes[old_boundary] + marker.raw_before.min(raw_counts[old_boundary]);
        }
        let hyperlink_revision_counts = self
            .hyperlinks
            .iter()
            .enumerate()
            .map(|(hyperlink_index, hyperlink)| {
                let start = hyperlink.run_start.min(old_run_count);
                let end = hyperlink.run_end.min(old_run_count);
                (start..=end)
                    .map(|boundary| {
                        self.revisions
                            .iter()
                            .filter(|(at, slot, _)| {
                                *at == boundary
                                    && hyperlink_revision_index(*slot) == Some(hyperlink_index)
                            })
                            .count()
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        for (run_index, raw_before, _) in &mut self.revisions {
            let old_boundary = (*run_index).min(old_run_count);
            *run_index = boundary_map[old_boundary];
            if hyperlink_revision_index(*raw_before).is_none() {
                *raw_before =
                    raw_prefixes[old_boundary] + (*raw_before).min(raw_counts[old_boundary]);
            }
        }
        for (position, _) in &mut self.extra_xml {
            *position = boundary_map[(*position).min(old_run_count)];
        }
        for (position, raw_before, _) in &mut self.equations {
            let old_boundary = (*position).min(old_run_count);
            *position = boundary_map[old_boundary];
            *raw_before = raw_prefixes[old_boundary] + (*raw_before).min(raw_counts[old_boundary]);
        }
        let old_hyperlinks = std::mem::take(&mut self.hyperlinks);
        let mut hyperlink_map = vec![None; old_hyperlinks.len()];
        for (old_index, mut hyperlink) in old_hyperlinks.into_iter().enumerate() {
            let old_start = hyperlink.run_start.min(old_run_count);
            let old_end = hyperlink.run_end.min(old_run_count);
            if let Some(raw_before) = hyperlink.preserved_raw_before {
                hyperlink.preserved_raw_before =
                    Some(raw_prefixes[old_start] + raw_before.min(raw_counts[old_start]));
            }
            let revision_counts = &hyperlink_revision_counts[old_index];
            let new_start = boundary_map[old_start];
            let new_end = boundary_map[old_end];
            for (boundary, revisions_before, _) in &mut hyperlink.extra_xml {
                let old_relative = (*boundary).min(old_end - old_start);
                let old_boundary = old_start + old_relative;
                let new_boundary = boundary_map[old_boundary];
                let collapsed_before = (0..old_relative)
                    .filter(|relative| boundary_map[old_start + relative] == new_boundary)
                    .map(|relative| revision_counts[relative])
                    .sum::<usize>();
                *boundary = new_boundary.saturating_sub(new_start);
                *revisions_before =
                    collapsed_before + (*revisions_before).min(revision_counts[old_relative]);
            }
            let owns_revision = self
                .revisions
                .iter()
                .any(|(_, slot, _)| hyperlink_revision_index(*slot) == Some(old_index));
            hyperlink.run_start = new_start;
            hyperlink.run_end = new_end;
            if new_start < new_end || owns_revision || !hyperlink.extra_xml.is_empty() {
                hyperlink_map[old_index] = Some(self.hyperlinks.len());
                self.hyperlinks.push(hyperlink);
            }
        }
        for (_, slot, _) in &mut self.revisions {
            let Some(old_index) = hyperlink_revision_index(*slot) else {
                continue;
            };
            if let Some(new_index) = hyperlink_map.get(old_index).copied().flatten() {
                *slot = hyperlink_revision_slot(new_index);
            }
        }
        self.runs = self
            .runs
            .drain(..)
            .zip(removed)
            .filter_map(|(run, remove)| (!remove).then_some(run))
            .collect();
    }

    /// Insert a canonical bookmark start marker at a direct-run boundary.
    pub fn insert_bookmark_start(&mut self, run_index: usize, id: i32, name: &str) -> bool {
        if run_index > self.runs.len() {
            return false;
        }
        let mut value = itoa::Buffer::new();
        let mut element = BytesStart::new("w:bookmarkStart");
        element.push_attribute(("w:id", value.format(id)));
        element.push_attribute(("w:name", name));
        let mut raw = Vec::new();
        if Writer::new(&mut raw)
            .write_event(Event::Empty(element))
            .is_err()
        {
            return false;
        }
        let raw_before = self
            .extra_xml
            .iter()
            .filter(|(at, _)| *at == run_index)
            .count();
        self.extra_xml.push((run_index, raw));
        self.bookmark_markers.push(BookmarkMarker::new(
            true,
            id,
            Some(name.to_owned()),
            run_index,
            raw_before,
        ));
        true
    }

    /// Insert a canonical bookmark end marker at a direct-run boundary.
    pub fn insert_bookmark_end(&mut self, run_index: usize, id: i32) -> bool {
        if run_index > self.runs.len() {
            return false;
        }
        let mut value = itoa::Buffer::new();
        let mut element = BytesStart::new("w:bookmarkEnd");
        element.push_attribute(("w:id", value.format(id)));
        let mut raw = Vec::new();
        if Writer::new(&mut raw)
            .write_event(Event::Empty(element))
            .is_err()
        {
            return false;
        }
        let raw_before = self
            .extra_xml
            .iter()
            .filter(|(at, _)| *at == run_index)
            .count();
        self.extra_xml.push((run_index, raw));
        self.bookmark_markers
            .push(BookmarkMarker::new(false, id, None, run_index, raw_before));
        true
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(
            reader,
            &[
                "w".to_string(),
                format!("\0r\0{R_NS}"),
                format!("\0mc\0{}", crate::namespace::MC_NS),
            ],
        )
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        reader.config_mut().trim_text(false);
        let mut properties = None;
        let mut runs = Vec::new();
        let mut run_sources = Vec::new();
        let mut hyperlinks = Vec::new();
        let mut comment_ranges = Vec::new();
        let mut bookmark_markers = Vec::new();
        let mut extra_xml = Vec::new();
        let mut content_controls = Vec::new();
        let mut revisions = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"pPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        properties = Some(parse_scoped_ppr(&raw, word_prefixes)?);
                    } else if is_word_element(name.as_ref(), b"r", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        runs.push(parse_run_raw(&raw, &prefixes)?);
                        run_sources.push(Some(raw));
                    } else if matches_local_name(name.as_ref(), b"r")
                        && !element_prefix_has_binding(name.as_ref(), &prefixes)
                    {
                        let prefixes = prefixes_with_assumed_word_owner(name.as_ref(), &prefixes)?;
                        let raw = capture_element(reader, e)?;
                        runs.push(parse_run_raw(&raw, &prefixes)?);
                        run_sources.push(Some(raw));
                    } else if is_word_element(name.as_ref(), b"sdt", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        if let Some(sdt) = CT_Sdt::from_raw(&raw, &prefixes) {
                            let raw_before = raw_xml_count_at(&extra_xml, runs.len());
                            let markers_before = comment_ranges
                                .iter()
                                .filter(|marker: &&CommentRangeMarker| {
                                    marker.run_index() == runs.len()
                                        && marker.raw_before() == raw_before
                                })
                                .count();
                            content_controls.push((runs.len(), raw_before, markers_before, sdt));
                        } else {
                            extra_xml.push((runs.len(), raw));
                        }
                    } else if is_word_element(name.as_ref(), b"hyperlink", &prefixes) {
                        let (rel_id, anchor, tooltip, doc_location, extra_attributes) =
                            parse_hyperlink_attributes(e, &prefixes)?;
                        let raw = capture_element(reader, e)?;
                        let parsed = parse_hyperlink_children(&raw, &prefixes)?;
                        let run_start = runs.len();
                        if parsed.runs.is_empty() && parsed.revisions.is_empty() {
                            extra_xml.push((run_start, raw));
                        } else {
                            let hyperlink_index = hyperlinks.len();
                            let run_end = run_start + parsed.runs.len();
                            let preserved_raw_before = parsed
                                .runs
                                .is_empty()
                                .then(|| raw_xml_count_at(&extra_xml, run_start));
                            runs.extend(parsed.runs);
                            run_sources.extend(parsed.run_sources);
                            hyperlinks.push(HyperlinkSpan {
                                rel_id,
                                anchor,
                                tooltip,
                                doc_location,
                                run_start,
                                run_end,
                                extra_attributes,
                                extra_xml: parsed.extra_xml,
                                preserved_raw_before,
                            });
                            revisions.extend(parsed.revisions.into_iter().map(|(at, revision)| {
                                (
                                    run_start + at,
                                    hyperlink_revision_slot(hyperlink_index),
                                    revision,
                                )
                            }));
                            if preserved_raw_before.is_some() {
                                extra_xml.push((run_start, raw));
                            }
                        }
                    } else if is_word_element(name.as_ref(), b"fldSimple", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        if let Some(field) = parse_simple_field(&raw, &prefixes)? {
                            runs.push(field_run(field, None));
                            run_sources.push(None);
                        } else {
                            extra_xml.push((runs.len(), raw));
                        }
                    } else if is_word_element(name.as_ref(), b"commentRangeStart", &prefixes)
                        || is_word_element(name.as_ref(), b"commentRangeEnd", &prefixes)
                    {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        push_comment_marker(
                            &mut comment_ranges,
                            &extra_xml,
                            runs.len(),
                            id,
                            is_word_element(name.as_ref(), b"commentRangeStart", &prefixes),
                        );
                    } else if is_word_element(name.as_ref(), b"bookmarkStart", &prefixes)
                        || is_word_element(name.as_ref(), b"bookmarkEnd", &prefixes)
                    {
                        let start = is_word_element(name.as_ref(), b"bookmarkStart", &prefixes);
                        let id = optional_word_attribute(e, b"id", &prefixes)
                            .and_then(|value| value.parse().ok());
                        let bookmark_name = optional_word_attribute(e, b"name", &prefixes);
                        bookmark_markers.push(BookmarkMarker {
                            start,
                            id,
                            name: bookmark_name,
                            run_index: runs.len(),
                            raw_before: raw_xml_count_at(&extra_xml, runs.len()),
                        });
                        extra_xml.push((runs.len(), capture_element(reader, e)?));
                    } else if is_word_element(name.as_ref(), b"ins", &prefixes)
                        || is_word_element(name.as_ref(), b"del", &prefixes)
                        || is_word_element(name.as_ref(), b"moveFrom", &prefixes)
                        || is_word_element(name.as_ref(), b"moveTo", &prefixes)
                    {
                        let raw_before = raw_xml_count_at(&extra_xml, runs.len());
                        let raw = capture_element(reader, e)?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            revisions.push((runs.len(), raw_before, revision));
                        }
                        extra_xml.push((runs.len(), raw));
                    } else {
                        // Capture unknown elements (bookmarks, comments, etc.) as raw XML
                        extra_xml.push((runs.len(), capture_element(reader, e)?));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"commentRangeStart", &prefixes)
                        || is_word_element(name.as_ref(), b"commentRangeEnd", &prefixes)
                    {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        push_comment_marker(
                            &mut comment_ranges,
                            &extra_xml,
                            runs.len(),
                            id,
                            is_word_element(name.as_ref(), b"commentRangeStart", &prefixes),
                        );
                    } else if is_word_element(name.as_ref(), b"bookmarkStart", &prefixes)
                        || is_word_element(name.as_ref(), b"bookmarkEnd", &prefixes)
                    {
                        let start = is_word_element(name.as_ref(), b"bookmarkStart", &prefixes);
                        let id = optional_word_attribute(e, b"id", &prefixes)
                            .and_then(|value| value.parse().ok());
                        let bookmark_name = optional_word_attribute(e, b"name", &prefixes);
                        bookmark_markers.push(BookmarkMarker {
                            start,
                            id,
                            name: bookmark_name,
                            run_index: runs.len(),
                            raw_before: raw_xml_count_at(&extra_xml, runs.len()),
                        });
                        extra_xml.push((runs.len(), capture_empty_element(e)?));
                    } else if !matches_local_name(name.as_ref(), b"p") {
                        let raw_before = raw_xml_count_at(&extra_xml, runs.len());
                        let raw = capture_empty_element(e)?;
                        if (is_word_element(name.as_ref(), b"ins", &prefixes)
                            || is_word_element(name.as_ref(), b"del", &prefixes)
                            || is_word_element(name.as_ref(), b"moveFrom", &prefixes)
                            || is_word_element(name.as_ref(), b"moveTo", &prefixes))
                            && let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes)
                        {
                            revisions.push((runs.len(), raw_before, revision));
                        }
                        extra_xml.push((runs.len(), raw));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"p") => {
                    break;
                }
                Ok(Event::Text(ref text)) if is_xml_whitespace(text.as_ref()) => {
                    extra_xml.push((runs.len(), text.as_ref().to_vec()));
                }
                Ok(Event::Comment(ref comment)) => {
                    extra_xml.push((
                        runs.len(),
                        capture_standalone_event(Event::Comment(comment.to_owned().into_owned()))?,
                    ));
                }
                Ok(Event::PI(ref instruction)) => {
                    extra_xml.push((
                        runs.len(),
                        capture_standalone_event(Event::PI(instruction.to_owned().into_owned()))?,
                    ));
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        project_complex_fields(ComplexFieldProjection {
            runs: &mut runs,
            run_sources: &mut run_sources,
            extra_xml: &mut extra_xml,
            comment_ranges: &mut comment_ranges,
            bookmark_markers: &mut bookmark_markers,
            content_controls: &mut content_controls,
            revisions: &mut revisions,
            hyperlinks: &mut hyperlinks,
            word_prefixes,
        })?;
        extra_xml.retain(|(_, raw)| !is_xml_whitespace(raw));
        for hyperlink in &mut hyperlinks {
            hyperlink
                .extra_xml
                .retain(|(_, _, raw)| !is_xml_whitespace(raw));
        }

        let inherited_bindings = namespace_bindings(word_prefixes);
        let mut equations = Vec::new();
        for run_index in 0..=runs.len() {
            for (raw_before, raw) in extra_xml
                .iter()
                .filter(|(position, _)| *position == run_index)
                .map(|(_, raw)| raw)
                .enumerate()
            {
                if let Some(equation) = OfficeMath::from_raw(raw, &inherited_bindings)? {
                    equations.push((run_index, raw_before, equation));
                }
            }
        }

        Ok(CT_P {
            properties,
            runs,
            hyperlinks,
            comment_ranges,
            bookmark_markers,
            extra_xml,
            content_controls,
            revisions,
            equations,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.to_xml_with_para_id(writer, None)
    }

    pub(crate) fn to_xml_with_para_id<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        para_id: Option<&str>,
    ) -> Result<()> {
        let mut start = BytesStart::new("w:p");
        if let Some(para_id) = para_id {
            start.push_attribute(("w14:paraId", para_id));
        }
        writer.write_event(Event::Start(start))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        // Build a set of run indices that are inside hyperlinks
        let mut hyperlink_runs: std::collections::HashMap<usize, usize> =
            std::collections::HashMap::new();
        for (hl_idx, hl) in self.hyperlinks.iter().enumerate() {
            for run_idx in hl.run_start..hl.run_end {
                hyperlink_runs.insert(run_idx, hl_idx);
            }
        }

        let mut current_hyperlink: Option<usize> = None;
        for (run_idx, run) in self.runs.iter().enumerate() {
            let in_hl = hyperlink_runs.get(&run_idx).copied();

            // Paragraph boundary content is a sibling of the hyperlink.
            if current_hyperlink.is_some() && current_hyperlink != in_hl {
                let hyperlink_index = current_hyperlink.expect("open hyperlink exists");
                write_hyperlink_boundary(
                    writer,
                    &self.revisions,
                    hyperlink_index,
                    &self.hyperlinks[hyperlink_index],
                    run_idx,
                )?;
                writer.write_event(Event::End(BytesEnd::new(hyperlink_qname(
                    &self.hyperlinks[hyperlink_index],
                ))))?;
                current_hyperlink = None;
            }

            write_paragraph_boundary(
                writer,
                ParagraphBoundary {
                    extra_xml: &self.extra_xml,
                    content_controls: &self.content_controls,
                    markers: &self.comment_ranges,
                    hyperlinks: &self.hyperlinks,
                    revisions: &self.revisions,
                    equations: &self.equations,
                },
                run_idx,
            )?;

            write_empty_hyperlinks(writer, &self.hyperlinks, &self.revisions, run_idx)?;

            // Open hyperlink if entering one
            if let Some(hl_idx) = in_hl
                && current_hyperlink != in_hl
            {
                let hl = &self.hyperlinks[hl_idx];
                write_hyperlink_start(writer, hl)?;
                current_hyperlink = in_hl;
            }

            if let Some(hyperlink_index) = current_hyperlink {
                write_hyperlink_boundary(
                    writer,
                    &self.revisions,
                    hyperlink_index,
                    &self.hyperlinks[hyperlink_index],
                    run_idx,
                )?;
            }

            // Fields are paragraph children even though the facade stores them
            // in synthetic runs for the existing inline traversal contract.
            if run.content.len() == 1
                && let RunContent::Field(field) = &run.content[0]
            {
                write_field(
                    writer,
                    field,
                    current_hyperlink
                        .and_then(|index| shadowed_word_namespace(&self.hyperlinks[index])),
                )?;
                continue;
            }

            if let Some(hyperlink_index) = current_hyperlink {
                run.to_xml_with_word_override(
                    writer,
                    shadowed_word_namespace(&self.hyperlinks[hyperlink_index]),
                )?;
            } else {
                run.to_xml(writer)?;
            }
        }

        // Close any remaining open hyperlink
        if let Some(hyperlink_index) = current_hyperlink {
            write_hyperlink_boundary(
                writer,
                &self.revisions,
                hyperlink_index,
                &self.hyperlinks[hyperlink_index],
                self.runs.len(),
            )?;
            writer.write_event(Event::End(BytesEnd::new(hyperlink_qname(
                &self.hyperlinks[hyperlink_index],
            ))))?;
        }

        write_paragraph_boundary(
            writer,
            ParagraphBoundary {
                extra_xml: &self.extra_xml,
                content_controls: &self.content_controls,
                markers: &self.comment_ranges,
                hyperlinks: &self.hyperlinks,
                revisions: &self.revisions,
                equations: &self.equations,
            },
            self.runs.len(),
        )?;
        write_empty_hyperlinks(writer, &self.hyperlinks, &self.revisions, self.runs.len())?;

        writer.write_event(Event::End(BytesEnd::new("w:p")))?;
        Ok(())
    }

    pub(crate) fn collect_runs<'a>(&'a self, runs: &mut Vec<&'a CT_R>) {
        for index in 0..=self.runs.len() {
            for (_, _, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _, _)| *at == index)
            {
                sdt.collect_runs(runs);
            }
            if let Some(run) = self.runs.get(index) {
                runs.push(run);
            }
        }
    }

    pub(crate) fn collect_controls<'a>(&'a self, controls: &mut Vec<&'a CT_Sdt>) {
        for (_, _, _, sdt) in &self.content_controls {
            controls.push(sdt);
            sdt.collect_controls(controls);
        }
    }
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    !value.is_empty()
        && value
            .iter()
            .all(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn capture_standalone_event(event: Event<'static>) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event)?;
    Ok(writer.into_inner())
}

fn write_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    if field.is_unchanged()
        && let FieldSource::Parsed { raw_xml, .. } = &field.source
    {
        return write_raw_with_word_override(writer, raw_xml, foreign_word_namespace);
    }

    if let FieldSource::Parsed {
        form,
        raw_xml,
        original_instruction,
        original_cached_result,
        original_dirty,
        word_prefixes,
        ..
    } = &field.source
        && instruction_structure_eq(&field.instruction, original_instruction)
        && instruction_source_identity_eq(&field.instruction, original_instruction)
        && field.instruction.raw == original_instruction.raw
    {
        let cached_changed = field.cached_result != *original_cached_result;
        let updated = match form {
            FieldForm::Simple => {
                update_simple_field_source(field, raw_xml, word_prefixes, cached_changed)?
            }
            FieldForm::Complex => {
                let updated = update_nested_field_sources(field, raw_xml, word_prefixes)?;
                if cached_changed || field.dirty != *original_dirty {
                    update_complex_field_source(field, &updated, word_prefixes, cached_changed)?
                } else {
                    updated
                }
            }
        };
        return write_raw_with_word_override(writer, &updated, foreign_word_namespace);
    }

    let mut field = field.clone();
    field.instruction = instruction_for_write(&field);

    let form = match &field.source {
        FieldSource::Parsed { form, .. } => *form,
        FieldSource::New { .. } => FieldForm::Simple,
    };
    let form = if instruction_contains_nested(&field.instruction) {
        FieldForm::Complex
    } else {
        form
    };
    match form {
        FieldForm::Simple => write_simple_field(writer, &field, foreign_word_namespace),
        FieldForm::Complex => write_complex_field(writer, &field, foreign_word_namespace),
    }
}

fn update_nested_field_sources(
    field: &Field,
    raw: &[u8],
    word_prefixes: &[String],
) -> Result<Vec<u8>> {
    let scan = scan_complex_source(raw, word_prefixes)?;
    let nested_fields = field.nested_fields_in_source_order();
    if scan.nested.len() != nested_fields.len() {
        return Err(OxmlError::MissingElement(
            "nested field spans in owning complex field".to_owned(),
        ));
    }
    let mut edits = Vec::new();
    for (nested, span) in nested_fields.into_iter().zip(scan.nested) {
        if nested.is_unchanged() {
            continue;
        }
        let replacement =
            rewrite_isolated_nested_field(nested, raw, &span, &scan.runs, word_prefixes)?;
        edits.push((span.start, span.end, replacement));
    }

    let mut updated = raw.to_vec();
    for (start, end, replacement) in edits.into_iter().rev() {
        updated.splice(start..end, replacement);
    }
    Ok(updated)
}

#[derive(Debug)]
struct ComplexSourceScan {
    runs: Vec<RunSourceSpan>,
    nested: Vec<NestedComplexSpan>,
}

#[derive(Debug)]
struct RunSourceSpan {
    start: usize,
    content_start: usize,
    content_end: usize,
    end: usize,
    field_content: Vec<(usize, usize)>,
}

#[derive(Debug)]
struct NestedComplexSpan {
    start: usize,
    end: usize,
    start_run: usize,
    end_run: usize,
}

fn is_word_source_name(name: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return word_prefixes.iter().any(String::is_empty);
    };
    word_prefixes
        .iter()
        .any(|prefix| prefix.as_bytes() == &name[..separator])
}

fn scan_complex_source(raw: &[u8], word_prefixes: &[String]) -> Result<ComplexSourceScan> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut runs = Vec::new();
    let mut open_fields = Vec::<OpenComplexSourceField>::new();
    let mut nested = Vec::new();
    loop {
        let run_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let run_prefixes = word_prefixes_at(&element, word_prefixes)?;
                if !is_word_element(element.name().as_ref(), b"r", &run_prefixes) {
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                    buffer.clear();
                    continue;
                }
                let run_index = runs.len();
                let content_start = reader.buffer_position() as usize;
                let mut field_content = Vec::new();
                loop {
                    buffer.clear();
                    let child_start = reader.buffer_position() as usize;
                    match reader.read_event_into(&mut buffer)? {
                        Event::Start(child) => {
                            let prefixes = word_prefixes_at(&child, &run_prefixes)?;
                            let is_field_content =
                                !is_word_element(child.name().as_ref(), b"rPr", &prefixes)
                                    && is_word_source_name(child.name().as_ref(), &prefixes);
                            if is_word_element(child.name().as_ref(), b"fldChar", &prefixes) {
                                let kind =
                                    optional_word_attribute(&child, b"fldCharType", &prefixes);
                                reader.read_to_end_into(child.name(), &mut Vec::new())?;
                                record_complex_source_marker(
                                    kind.as_deref(),
                                    child_start,
                                    reader.buffer_position() as usize,
                                    run_index,
                                    &mut open_fields,
                                    &mut nested,
                                );
                            } else {
                                reader.read_to_end_into(child.name(), &mut Vec::new())?;
                            }
                            if is_field_content {
                                field_content
                                    .push((child_start, reader.buffer_position() as usize));
                            }
                        }
                        Event::Empty(child) => {
                            let prefixes = word_prefixes_at(&child, &run_prefixes)?;
                            let is_field_content =
                                !is_word_element(child.name().as_ref(), b"rPr", &prefixes)
                                    && is_word_source_name(child.name().as_ref(), &prefixes);
                            if is_word_element(child.name().as_ref(), b"fldChar", &prefixes) {
                                let kind =
                                    optional_word_attribute(&child, b"fldCharType", &prefixes);
                                record_complex_source_marker(
                                    kind.as_deref(),
                                    child_start,
                                    reader.buffer_position() as usize,
                                    run_index,
                                    &mut open_fields,
                                    &mut nested,
                                );
                            }
                            if is_field_content {
                                field_content
                                    .push((child_start, reader.buffer_position() as usize));
                            }
                        }
                        Event::End(end) if matches_local_name(end.name().as_ref(), b"r") => {
                            runs.push(RunSourceSpan {
                                start: run_start,
                                content_start,
                                content_end: child_start,
                                end: reader.buffer_position() as usize,
                                field_content,
                            });
                            break;
                        }
                        Event::Eof => {
                            return Err(OxmlError::MissingElement(
                                "end of field source run".to_owned(),
                            ));
                        }
                        _ => {}
                    }
                }
            }
            Event::Eof => return Ok(ComplexSourceScan { runs, nested }),
            _ => {}
        }
        buffer.clear();
    }
}

fn field_source_has_unmodeled_semantic_attributes(
    raw: &[u8],
    form: FieldForm,
    word_prefixes: &[String],
) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                let prefixes = word_prefixes_at(&element, word_prefixes)?;
                let allowed = match form {
                    FieldForm::Simple
                        if is_word_element(element.name().as_ref(), b"fldSimple", &prefixes) =>
                    {
                        &[b"instr".as_slice(), b"dirty".as_slice()][..]
                    }
                    FieldForm::Complex
                        if is_word_element(element.name().as_ref(), b"fldChar", &prefixes) =>
                    {
                        &[b"fldCharType".as_slice(), b"dirty".as_slice()][..]
                    }
                    _ => {
                        buffer.clear();
                        continue;
                    }
                };

                for attribute in element.attributes() {
                    let attribute = attribute?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    if allowed
                        .iter()
                        .any(|local| is_word_attribute(key, local, &prefixes))
                    {
                        continue;
                    }
                    return Ok(true);
                }

                if matches!(form, FieldForm::Simple) {
                    return Ok(false);
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn record_complex_source_marker(
    kind: Option<&str>,
    start: usize,
    end: usize,
    run_index: usize,
    open_fields: &mut Vec<OpenComplexSourceField>,
    nested: &mut Vec<NestedComplexSpan>,
) {
    match kind {
        Some("begin") => open_fields.push(OpenComplexSourceField {
            start,
            start_run: run_index,
            separated: false,
        }),
        Some("separate") => {
            if let Some(field) = open_fields.last_mut() {
                field.separated = true;
            }
        }
        Some("end") => {
            let Some(field) = open_fields.pop() else {
                return;
            };
            if open_fields.len() == 1 && !open_fields[0].separated {
                nested.push(NestedComplexSpan {
                    start: field.start,
                    end,
                    start_run: field.start_run,
                    end_run: run_index,
                });
            }
        }
        _ => {}
    }
}

struct OpenComplexSourceField {
    start: usize,
    start_run: usize,
    separated: bool,
}

fn rewrite_isolated_nested_field(
    field: &Field,
    raw: &[u8],
    field_span: &NestedComplexSpan,
    runs: &[RunSourceSpan],
    word_prefixes: &[String],
) -> Result<Vec<u8>> {
    let start_run = &runs[field_span.start_run];
    let end_run = &runs[field_span.end_run];
    let mut isolated = raw[start_run.start..start_run.content_start].to_vec();
    if field_span.start_run == field_span.end_run {
        isolated.extend_from_slice(&raw[field_span.start..field_span.end]);
    } else {
        isolated.extend_from_slice(&raw[field_span.start..start_run.end]);
        isolated.extend_from_slice(&raw[start_run.end..end_run.content_start]);
        isolated.extend_from_slice(&raw[end_run.content_start..field_span.end]);
    }
    isolated.extend_from_slice(&raw[end_run.content_end..end_run.end]);

    let FieldSource::Parsed {
        original_instruction,
        original_cached_result,
        original_dirty,
        ..
    } = &field.source
    else {
        return Err(OxmlError::MissingElement(
            "parsed nested field source".to_owned(),
        ));
    };
    let raw_only_instruction_changed = field.instruction.raw != original_instruction.raw
        && instruction_structure_eq(&field.instruction, original_instruction);
    let mut updated = if raw_only_instruction_changed {
        remove_nested_field_sources(&isolated, word_prefixes)?
    } else {
        update_nested_field_sources(field, &isolated, word_prefixes)?
    };
    if raw_only_instruction_changed {
        updated = update_complex_instruction_source(field, &updated, word_prefixes)?;
    }
    let cached_changed = field.cached_result != *original_cached_result;
    if cached_changed || field.dirty != *original_dirty {
        updated = update_complex_field_source(field, &updated, word_prefixes, cached_changed)?;
    }
    let updated_scan = scan_complex_source(&updated, word_prefixes)?;
    let (Some(first_run), Some(last_run)) = (updated_scan.runs.first(), updated_scan.runs.last())
    else {
        return Err(OxmlError::MissingElement(
            "updated nested field runs".to_owned(),
        ));
    };
    Ok(updated[first_run.content_start..last_run.content_end].to_vec())
}

fn remove_nested_field_sources(raw: &[u8], word_prefixes: &[String]) -> Result<Vec<u8>> {
    let scan = scan_complex_source(raw, word_prefixes)?;
    let mut removals = Vec::new();
    for span in &scan.nested {
        for run in &scan.runs[span.start_run..=span.end_run] {
            removals.extend(
                run.field_content
                    .iter()
                    .copied()
                    .filter(|(start, end)| *start >= span.start && *end <= span.end),
            );
        }
    }
    removals.sort_unstable();
    removals.dedup();
    let mut updated = raw.to_vec();
    for (start, end) in removals.into_iter().rev() {
        updated.drain(start..end);
    }
    Ok(updated)
}

struct FieldRewriteContext {
    prefixes: Vec<String>,
    word_run: bool,
    canonical_end: Option<&'static str>,
}

fn update_simple_field_source(
    field: &Field,
    raw: &[u8],
    word_prefixes: &[String],
    cached_changed: bool,
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    let mut contexts = Vec::<FieldRewriteContext>::new();
    let mut buffer = Vec::new();
    let mut wrote_result = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let depth = contexts.len();
                let parent_is_word_run = contexts.last().is_some_and(|context| context.word_run);
                if depth == 0 && is_word_element(element.name().as_ref(), b"fldSimple", &prefixes) {
                    write_rewritten_field_element(
                        &mut writer,
                        &element,
                        "w:fldSimple",
                        &prefixes,
                        &[("instr", field.instruction.raw.as_str())],
                        field.dirty,
                        true,
                        false,
                    )?;
                    contexts.push(FieldRewriteContext {
                        prefixes,
                        word_run: false,
                        canonical_end: Some("w:fldSimple"),
                    });
                } else if depth == 2
                    && parent_is_word_run
                    && (is_word_element(element.name().as_ref(), b"t", &prefixes)
                        || is_word_element(element.name().as_ref(), b"delText", &prefixes))
                {
                    if cached_changed {
                        write_updated_text_element(
                            &mut reader,
                            &mut writer,
                            &element,
                            (!wrote_result).then_some(field.cached_result.as_str()),
                        )?;
                        wrote_result = true;
                    } else {
                        writer.write_event(Event::Start(element.into_owned()))?;
                        contexts.push(FieldRewriteContext {
                            prefixes,
                            word_run: false,
                            canonical_end: None,
                        });
                    }
                } else if cached_changed
                    && depth == 2
                    && parent_is_word_run
                    && (is_word_element(element.name().as_ref(), b"tab", &prefixes)
                        || is_word_element(element.name().as_ref(), b"br", &prefixes))
                {
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                    if !wrote_result {
                        write_field_result_content(&mut writer, &field.cached_result)?;
                        wrote_result = true;
                    }
                } else {
                    let word_run =
                        depth == 1 && is_word_element(element.name().as_ref(), b"r", &prefixes);
                    writer.write_event(Event::Start(element.into_owned()))?;
                    contexts.push(FieldRewriteContext {
                        prefixes,
                        word_run,
                        canonical_end: None,
                    });
                }
            }
            Event::Empty(element) => {
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let depth = contexts.len();
                if depth == 2
                    && contexts.last().is_some_and(|context| context.word_run)
                    && (is_word_element(element.name().as_ref(), b"t", &prefixes)
                        || is_word_element(element.name().as_ref(), b"delText", &prefixes))
                {
                    if cached_changed {
                        write_updated_empty_text_element(
                            &mut writer,
                            &element,
                            (!wrote_result).then_some(field.cached_result.as_str()),
                        )?;
                        wrote_result = true;
                    } else {
                        writer.write_event(Event::Empty(element.into_owned()))?;
                    }
                } else if cached_changed
                    && depth == 2
                    && contexts.last().is_some_and(|context| context.word_run)
                    && (is_word_element(element.name().as_ref(), b"tab", &prefixes)
                        || is_word_element(element.name().as_ref(), b"br", &prefixes))
                {
                    if !wrote_result {
                        write_field_result_content(&mut writer, &field.cached_result)?;
                        wrote_result = true;
                    }
                } else {
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) => {
                let Some(context) = contexts.pop() else {
                    writer.write_event(Event::End(element.into_owned()))?;
                    buffer.clear();
                    continue;
                };
                if cached_changed && context.canonical_end == Some("w:fldSimple") && !wrote_result {
                    write_field_result_run(&mut writer, field, None)?;
                    wrote_result = true;
                }
                if let Some(name) = context.canonical_end {
                    writer.write_event(Event::End(BytesEnd::new(name)))?;
                } else {
                    writer.write_event(Event::End(element.into_owned()))?;
                }
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    Ok(output)
}

fn update_complex_instruction_source(
    field: &Field,
    raw: &[u8],
    word_prefixes: &[String],
) -> Result<Vec<u8>> {
    let effective = instruction_for_write(field);
    let instruction = format!(" {} ", canonical_instruction_text(&effective));
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    let mut contexts = Vec::<FieldRewriteContext>::new();
    let mut buffer = Vec::new();
    let mut field_depth = 0usize;
    let mut outer_result = false;
    let mut wrote_instruction = false;
    let canonical_word_binding = !word_prefixes.iter().any(|prefix| prefix == "w");

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let direct_run_child =
                    contexts.len() == 1 && contexts.last().is_some_and(|context| context.word_run);
                if direct_run_child
                    && is_word_element(element.name().as_ref(), b"fldChar", &prefixes)
                {
                    let kind = optional_word_attribute(&element, b"fldCharType", &prefixes);
                    update_complex_phase(kind.as_deref(), &mut field_depth, &mut outer_result);
                }
                if direct_run_child
                    && field_depth == 1
                    && !outer_result
                    && is_word_element(element.name().as_ref(), b"instrText", &prefixes)
                {
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                    write_updated_instruction_text(
                        &mut writer,
                        &element,
                        (!wrote_instruction).then_some(instruction.as_str()),
                        canonical_word_binding,
                    )?;
                    wrote_instruction = true;
                } else {
                    let word_run = contexts.is_empty()
                        && is_word_element(element.name().as_ref(), b"r", &prefixes);
                    writer.write_event(Event::Start(element.into_owned()))?;
                    contexts.push(FieldRewriteContext {
                        prefixes,
                        word_run,
                        canonical_end: None,
                    });
                }
            }
            Event::Empty(element) => {
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let direct_run_child =
                    contexts.len() == 1 && contexts.last().is_some_and(|context| context.word_run);
                if direct_run_child
                    && is_word_element(element.name().as_ref(), b"fldChar", &prefixes)
                {
                    let kind = optional_word_attribute(&element, b"fldCharType", &prefixes);
                    update_complex_phase(kind.as_deref(), &mut field_depth, &mut outer_result);
                    writer.write_event(Event::Empty(element.into_owned()))?;
                } else if direct_run_child
                    && field_depth == 1
                    && !outer_result
                    && is_word_element(element.name().as_ref(), b"instrText", &prefixes)
                {
                    write_updated_instruction_text(
                        &mut writer,
                        &element,
                        (!wrote_instruction).then_some(instruction.as_str()),
                        canonical_word_binding,
                    )?;
                    wrote_instruction = true;
                } else {
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) => {
                contexts.pop();
                writer.write_event(Event::End(element.into_owned()))?;
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    if !wrote_instruction {
        return Err(OxmlError::MissingElement(
            "nested complex field instruction text".to_owned(),
        ));
    }
    Ok(output)
}

fn write_updated_instruction_text<W: std::io::Write>(
    writer: &mut Writer<W>,
    source: &BytesStart<'_>,
    value: Option<&str>,
    canonical_word_binding: bool,
) -> Result<()> {
    let value = value.unwrap_or_default();
    let mut element = BytesStart::new("w:instrText");
    for attribute in source.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() != b"xml:space"
            && !(canonical_word_binding && attribute.key.as_ref() == b"xmlns:w")
        {
            element.push_attribute(attribute);
        }
    }
    if canonical_word_binding {
        element.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    if needs_space_preservation(value) {
        element.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(element))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new("w:instrText")))?;
    Ok(())
}

fn update_complex_field_source(
    field: &Field,
    raw: &[u8],
    word_prefixes: &[String],
    cached_changed: bool,
) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    let mut contexts = Vec::<FieldRewriteContext>::new();
    let mut buffer = Vec::new();
    let mut field_depth = 0usize;
    let mut outer_result = false;
    let mut wrote_result = false;
    let canonical_word_binding = !word_prefixes.iter().any(|prefix| prefix == "w");

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let depth = contexts.len();
                let direct_run_child =
                    depth == 1 && contexts.last().is_some_and(|context| context.word_run);
                if direct_run_child
                    && is_word_element(element.name().as_ref(), b"fldChar", &prefixes)
                {
                    let kind = optional_word_attribute(&element, b"fldCharType", &prefixes);
                    let depth_before = field_depth;
                    let outer_marker = matches!(kind.as_deref(), Some("begin")) && field_depth == 0
                        || matches!(kind.as_deref(), Some("separate" | "end")) && field_depth == 1;
                    if outer_marker
                        && kind.as_deref() == Some("end")
                        && outer_result
                        && !wrote_result
                        && cached_changed
                    {
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            canonical_word_binding,
                        )?;
                        wrote_result = true;
                    }
                    if cached_changed
                        && outer_result
                        && depth_before == 1
                        && kind.as_deref() == Some("begin")
                        && !wrote_result
                    {
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            canonical_word_binding,
                        )?;
                        wrote_result = true;
                    }
                    update_complex_phase(kind.as_deref(), &mut field_depth, &mut outer_result);
                    if outer_marker {
                        let kind = kind.as_deref().unwrap_or_default();
                        write_rewritten_field_element(
                            &mut writer,
                            &element,
                            "w:fldChar",
                            &prefixes,
                            &[("fldCharType", kind)],
                            (kind == "begin").then_some(field.dirty).flatten(),
                            canonical_word_binding,
                            false,
                        )?;
                        contexts.push(FieldRewriteContext {
                            prefixes,
                            word_run: false,
                            canonical_end: Some("w:fldChar"),
                        });
                        buffer.clear();
                        continue;
                    }
                    if cached_changed
                        && wrote_result
                        && outer_result
                        && (depth_before > 1
                            || depth_before == 1 && kind.as_deref() == Some("begin"))
                    {
                        reader.read_to_end_into(element.name(), &mut Vec::new())?;
                        buffer.clear();
                        continue;
                    }
                }
                if direct_run_child
                    && outer_result
                    && field_depth == 1
                    && is_word_element(element.name().as_ref(), b"t", &prefixes)
                {
                    if cached_changed {
                        reader.read_to_end_into(element.name(), &mut Vec::new())?;
                        if !wrote_result {
                            write_field_result_content_with_binding(
                                &mut writer,
                                &field.cached_result,
                                canonical_word_binding,
                            )?;
                            wrote_result = true;
                        }
                    } else {
                        writer.write_event(Event::Start(element.into_owned()))?;
                        contexts.push(FieldRewriteContext {
                            prefixes,
                            word_run: false,
                            canonical_end: None,
                        });
                    }
                } else if direct_run_child
                    && cached_changed
                    && ((outer_result
                        && field_depth == 1
                        && (is_word_element(element.name().as_ref(), b"tab", &prefixes)
                            || is_word_element(element.name().as_ref(), b"br", &prefixes)))
                        || outer_result
                            && field_depth > 1
                            && (is_word_element(element.name().as_ref(), b"instrText", &prefixes)
                                || is_word_element(element.name().as_ref(), b"t", &prefixes)
                                || is_word_element(element.name().as_ref(), b"tab", &prefixes)
                                || is_word_element(element.name().as_ref(), b"br", &prefixes)))
                {
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                    if outer_result && field_depth == 1 && !wrote_result {
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            canonical_word_binding,
                        )?;
                        wrote_result = true;
                    }
                } else {
                    let word_run =
                        depth == 0 && is_word_element(element.name().as_ref(), b"r", &prefixes);
                    writer.write_event(Event::Start(element.into_owned()))?;
                    contexts.push(FieldRewriteContext {
                        prefixes,
                        word_run,
                        canonical_end: None,
                    });
                }
            }
            Event::Empty(element) => {
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let direct_run_child =
                    contexts.len() == 1 && contexts.last().is_some_and(|context| context.word_run);
                if direct_run_child
                    && is_word_element(element.name().as_ref(), b"fldChar", &prefixes)
                {
                    let kind = optional_word_attribute(&element, b"fldCharType", &prefixes);
                    let depth_before = field_depth;
                    let outer_marker = matches!(kind.as_deref(), Some("begin")) && field_depth == 0
                        || matches!(kind.as_deref(), Some("separate" | "end")) && field_depth == 1;
                    if outer_marker
                        && kind.as_deref() == Some("end")
                        && outer_result
                        && !wrote_result
                        && cached_changed
                    {
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            canonical_word_binding,
                        )?;
                        wrote_result = true;
                    }
                    if cached_changed
                        && outer_result
                        && depth_before == 1
                        && kind.as_deref() == Some("begin")
                        && !wrote_result
                    {
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            canonical_word_binding,
                        )?;
                        wrote_result = true;
                    }
                    update_complex_phase(kind.as_deref(), &mut field_depth, &mut outer_result);
                    if outer_marker {
                        let kind = kind.as_deref().unwrap_or_default();
                        write_rewritten_field_element(
                            &mut writer,
                            &element,
                            "w:fldChar",
                            &prefixes,
                            &[("fldCharType", kind)],
                            (kind == "begin").then_some(field.dirty).flatten(),
                            canonical_word_binding,
                            true,
                        )?;
                        buffer.clear();
                        continue;
                    }
                    if cached_changed
                        && wrote_result
                        && outer_result
                        && (depth_before > 1
                            || depth_before == 1 && kind.as_deref() == Some("begin"))
                    {
                        buffer.clear();
                        continue;
                    }
                }
                if direct_run_child
                    && outer_result
                    && field_depth == 1
                    && is_word_element(element.name().as_ref(), b"t", &prefixes)
                {
                    if cached_changed {
                        if !wrote_result {
                            write_field_result_content_with_binding(
                                &mut writer,
                                &field.cached_result,
                                canonical_word_binding,
                            )?;
                            wrote_result = true;
                        }
                    } else {
                        writer.write_event(Event::Empty(element.into_owned()))?;
                    }
                } else if direct_run_child
                    && cached_changed
                    && ((outer_result
                        && field_depth == 1
                        && (is_word_element(element.name().as_ref(), b"tab", &prefixes)
                            || is_word_element(element.name().as_ref(), b"br", &prefixes)))
                        || outer_result
                            && field_depth > 1
                            && (is_word_element(element.name().as_ref(), b"instrText", &prefixes)
                                || is_word_element(element.name().as_ref(), b"t", &prefixes)
                                || is_word_element(element.name().as_ref(), b"tab", &prefixes)
                                || is_word_element(element.name().as_ref(), b"br", &prefixes)))
                {
                    if outer_result && field_depth == 1 && !wrote_result {
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            canonical_word_binding,
                        )?;
                        wrote_result = true;
                    }
                } else {
                    writer.write_event(Event::Empty(element.into_owned()))?;
                }
            }
            Event::End(element) => {
                if let Some(name) = contexts.pop().and_then(|context| context.canonical_end) {
                    writer.write_event(Event::End(BytesEnd::new(name)))?;
                } else {
                    writer.write_event(Event::End(element.into_owned()))?;
                }
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
    Ok(output)
}

fn update_complex_phase(kind: Option<&str>, field_depth: &mut usize, outer_result: &mut bool) {
    match kind {
        Some("begin") => *field_depth += 1,
        Some("separate") if *field_depth == 1 => *outer_result = true,
        Some("end") if *field_depth == 1 => {
            *outer_result = false;
            *field_depth = 0;
        }
        Some("end") if *field_depth > 1 => *field_depth -= 1,
        _ => {}
    }
}

#[allow(clippy::too_many_arguments)]
fn write_rewritten_field_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    source: &BytesStart<'_>,
    name: &'static str,
    word_prefixes: &[String],
    replacements: &[(&str, &str)],
    dirty: Option<bool>,
    canonical_word_binding: bool,
    empty: bool,
) -> Result<()> {
    let mut element = BytesStart::new(name);
    for attribute in source.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if canonical_word_binding && key == b"xmlns:w" {
            continue;
        }
        if replacements
            .iter()
            .any(|(local, _)| is_field_word_attribute(key, local.as_bytes(), word_prefixes))
            || is_field_word_attribute(key, b"dirty", word_prefixes)
        {
            continue;
        }
        element.push_attribute(attribute);
    }
    if canonical_word_binding {
        element.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    for (local, value) in replacements {
        let name = match *local {
            "instr" => "w:instr",
            "fldCharType" => "w:fldCharType",
            _ => {
                return Err(OxmlError::MissingElement(format!(
                    "field attribute {local}"
                )));
            }
        };
        element.push_attribute((name, *value));
    }
    push_dirty_attribute(&mut element, dirty);
    writer.write_event(if empty {
        Event::Empty(element)
    } else {
        Event::Start(element)
    })?;
    Ok(())
}

fn is_field_word_attribute(key: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = key.iter().position(|byte| *byte == b':') else {
        return false;
    };
    key.get(separator + 1..) == Some(local)
        && word_prefixes
            .iter()
            .any(|prefix| prefix.as_bytes() == &key[..separator])
}

fn write_updated_text_element<W: std::io::Write>(
    reader: &mut Reader<&[u8]>,
    writer: &mut Writer<W>,
    source: &BytesStart<'_>,
    value: Option<&str>,
) -> Result<()> {
    reader.read_to_end_into(source.name(), &mut Vec::new())?;
    if let Some(value) = value {
        write_field_result_content_with_source(writer, value, Some(source))?;
    } else {
        write_field_text_with_source(writer, "", Some(source), false)?;
    }
    Ok(())
}

fn write_updated_empty_text_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    source: &BytesStart<'_>,
    value: Option<&str>,
) -> Result<()> {
    if let Some(value) = value {
        write_field_result_content_with_source(writer, value, Some(source))?;
    } else {
        write_field_text_with_source(writer, "", Some(source), false)?;
    }
    Ok(())
}

fn write_field_result_content<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &str,
) -> Result<()> {
    write_field_result_content_with_binding(writer, value, false)
}

fn write_field_result_content_with_source<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &str,
    source: Option<&BytesStart<'_>>,
) -> Result<()> {
    write_field_result_content_inner(writer, value, source, false)
}

fn write_field_result_content_with_binding<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &str,
    canonical_word_binding: bool,
) -> Result<()> {
    write_field_result_content_inner(writer, value, None, canonical_word_binding)
}

fn write_field_result_content_inner<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &str,
    source: Option<&BytesStart<'_>>,
    canonical_word_binding: bool,
) -> Result<()> {
    let mut start = 0usize;
    let mut used_source = false;
    for (index, character) in value
        .char_indices()
        .chain(std::iter::once((value.len(), '\0')))
    {
        let element = match character {
            '\t' => Some(("w:tab", None)),
            '\n' => Some(("w:br", None)),
            '\u{000c}' => Some(("w:br", Some("page"))),
            '\u{000b}' => Some(("w:br", Some("column"))),
            '\0' if index == value.len() => None,
            _ => continue,
        };
        if start < index || value.is_empty() {
            write_field_text_with_source(
                writer,
                &value[start..index],
                (!used_source).then_some(source).flatten(),
                canonical_word_binding,
            )?;
            used_source = true;
        }
        if let Some((name, break_type)) = element {
            if !used_source && source.is_some() {
                write_field_text_with_source(writer, "", source, canonical_word_binding)?;
                used_source = true;
            }
            let mut element = BytesStart::new(name);
            if canonical_word_binding {
                element.push_attribute(("xmlns:w", crate::namespace::W_NS));
            }
            if let Some(break_type) = break_type {
                element.push_attribute(("w:type", break_type));
            }
            writer.write_event(Event::Empty(element))?;
            start = index + character.len_utf8();
        }
    }
    Ok(())
}

fn write_field_text_with_source<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &str,
    source: Option<&BytesStart<'_>>,
    canonical_word_binding: bool,
) -> Result<()> {
    let mut element = BytesStart::new("w:t");
    if let Some(source) = source {
        for attribute in source.attributes() {
            let attribute = attribute?;
            if attribute.key.as_ref() != b"xml:space" {
                element.push_attribute(attribute);
            }
        }
    } else if canonical_word_binding {
        element.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    if needs_space_preservation(value) {
        element.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(element))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new("w:t")))?;
    Ok(())
}

fn needs_space_preservation(value: &str) -> bool {
    value.chars().next().is_some_and(char::is_whitespace)
        || value.chars().next_back().is_some_and(char::is_whitespace)
}

fn instruction_for_write(field: &Field) -> FieldInstruction {
    let original = match &field.source {
        FieldSource::New {
            original_instruction,
        }
        | FieldSource::Parsed {
            original_instruction,
            ..
        } => original_instruction,
    };
    let raw_changed = field.instruction.raw != original.raw;
    let structured_changed = !instruction_structure_eq(&field.instruction, original);

    if raw_changed && !structured_changed {
        return parse_field_instruction(&field.instruction.raw);
    }
    if structured_changed {
        let mut instruction = field.instruction.clone();
        instruction.raw = canonical_instruction_text(&instruction);
        return instruction;
    }
    field.instruction.clone()
}

fn instruction_structure_eq(left: &FieldInstruction, right: &FieldInstruction) -> bool {
    left.name == right.name
        && left.arguments.len() == right.arguments.len()
        && left
            .arguments
            .iter()
            .zip(&right.arguments)
            .all(|(left, right)| match (left, right) {
                (FieldArgument::Text(left), FieldArgument::Text(right)) => left == right,
                (FieldArgument::Nested(left), FieldArgument::Nested(right)) => {
                    instruction_structure_eq(&left.instruction, &right.instruction)
                }
                _ => false,
            })
        && left.switches.len() == right.switches.len()
        && left
            .switches
            .iter()
            .zip(&right.switches)
            .all(|(left, right)| {
                left.name == right.name
                    && match (&left.argument, &right.argument) {
                        (None, None) => true,
                        (Some(FieldArgument::Text(left)), Some(FieldArgument::Text(right))) => {
                            left == right
                        }
                        (Some(FieldArgument::Nested(left)), Some(FieldArgument::Nested(right))) => {
                            instruction_structure_eq(&left.instruction, &right.instruction)
                        }
                        _ => false,
                    }
            })
}

fn instruction_source_identity_eq(left: &FieldInstruction, right: &FieldInstruction) -> bool {
    let argument_identities_match =
        left.arguments
            .iter()
            .zip(&right.arguments)
            .all(|(left, right)| match (left, right) {
                (FieldArgument::Nested(left), FieldArgument::Nested(right)) => {
                    field_source_identity_eq(left, right)
                }
                (FieldArgument::Text(_), FieldArgument::Text(_)) => true,
                _ => false,
            });
    let switch_identities_match = left
        .switches
        .iter()
        .zip(&right.switches)
        .all(|(left, right)| match (&left.argument, &right.argument) {
            (Some(FieldArgument::Nested(left)), Some(FieldArgument::Nested(right))) => {
                field_source_identity_eq(left, right)
            }
            (None, None) | (Some(FieldArgument::Text(_)), Some(FieldArgument::Text(_))) => true,
            _ => false,
        });
    argument_identities_match && switch_identities_match
}

fn field_source_identity_eq(left: &Field, right: &Field) -> bool {
    match (&left.source, &right.source) {
        (
            FieldSource::Parsed {
                source_id: left, ..
            },
            FieldSource::Parsed {
                source_id: right, ..
            },
        ) => left == right,
        _ => false,
    }
}

fn canonical_instruction_text(instruction: &FieldInstruction) -> String {
    let mut text = instruction.name.clone();
    for argument in &instruction.arguments {
        if let FieldArgument::Text(value) = argument {
            push_canonical_field_token(&mut text, value);
        }
    }
    for switch in &instruction.switches {
        text.push(' ');
        text.push('\\');
        text.push_str(&switch.name);
        if let Some(FieldArgument::Text(value)) = &switch.argument {
            push_canonical_field_token(&mut text, value);
        }
    }
    text
}

fn instruction_contains_nested(instruction: &FieldInstruction) -> bool {
    instruction
        .arguments
        .iter()
        .any(|argument| matches!(argument, FieldArgument::Nested(_)))
        || instruction
            .switches
            .iter()
            .any(|switch| matches!(switch.argument, Some(FieldArgument::Nested(_))))
}

fn write_simple_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let mut element = BytesStart::new("w:fldSimple");
    if foreign_word_namespace.is_some() {
        element.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    element.push_attribute(("w:instr", field.instruction.raw.as_str()));
    push_dirty_attribute(&mut element, field.dirty);
    writer.write_event(Event::Start(element))?;
    write_field_result_run(writer, field, foreign_word_namespace)?;
    writer.write_event(Event::End(BytesEnd::new("w:fldSimple")))?;
    Ok(())
}

fn write_complex_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    write_field_char_run(writer, "begin", field.dirty, foreign_word_namespace)?;
    let mut text = format!(" {}", field.instruction.name);
    for argument in &field.instruction.arguments {
        match argument {
            FieldArgument::Text(value) => push_canonical_field_token(&mut text, value),
            FieldArgument::Nested(field) => {
                text.push(' ');
                write_instruction_run(writer, &text, foreign_word_namespace)?;
                text.clear();
                write_nested_instruction_field(writer, field, foreign_word_namespace)?;
            }
        }
    }
    for switch in &field.instruction.switches {
        text.push(' ');
        text.push('\\');
        text.push_str(&switch.name);
        if let Some(argument) = &switch.argument {
            match argument {
                FieldArgument::Text(value) => push_canonical_field_token(&mut text, value),
                FieldArgument::Nested(field) => {
                    text.push(' ');
                    write_instruction_run(writer, &text, foreign_word_namespace)?;
                    text.clear();
                    write_nested_instruction_field(writer, field, foreign_word_namespace)?;
                }
            }
        }
    }
    text.push(' ');
    if !text.trim().is_empty() {
        write_instruction_run(writer, &text, foreign_word_namespace)?;
    }
    write_field_char_run(writer, "separate", None, foreign_word_namespace)?;
    write_field_result_run(writer, field, foreign_word_namespace)?;
    write_field_char_run(writer, "end", None, foreign_word_namespace)?;
    Ok(())
}

fn write_nested_instruction_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let mut field = field.clone();
    field.instruction = instruction_for_write(&field);
    write_complex_field(writer, &field, foreign_word_namespace)
}

fn push_canonical_field_token(output: &mut String, value: &str) {
    output.push(' ');
    if value.is_empty()
        || value.starts_with('\\')
        || value.chars().any(char::is_whitespace)
        || value.contains('"')
    {
        output.push('"');
        for character in value.chars() {
            if matches!(character, '"' | '\\') {
                output.push('\\');
            }
            output.push(character);
        }
        output.push('"');
    } else {
        output.push_str(value);
    }
}

fn write_instruction_run<W: std::io::Write>(
    writer: &mut Writer<W>,
    instruction: &str,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    write_word_run_start(writer, foreign_word_namespace)?;
    let mut element = BytesStart::new("w:instrText");
    element.push_attribute(("xml:space", "preserve"));
    writer.write_event(Event::Start(element))?;
    writer.write_event(Event::Text(BytesText::new(instruction)))?;
    writer.write_event(Event::End(BytesEnd::new("w:instrText")))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn write_field_char_run<W: std::io::Write>(
    writer: &mut Writer<W>,
    kind: &str,
    dirty: Option<bool>,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    write_word_run_start(writer, foreign_word_namespace)?;
    let mut element = BytesStart::new("w:fldChar");
    element.push_attribute(("w:fldCharType", kind));
    push_dirty_attribute(&mut element, dirty);
    writer.write_event(Event::Empty(element))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn write_field_result_run<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    write_word_run_start(writer, foreign_word_namespace)?;
    let display = if field.cached_result.is_empty()
        && matches!(field.instruction.name.as_str(), "PAGE" | "NUMPAGES")
    {
        "1"
    } else {
        &field.cached_result
    };
    write_field_result_content(writer, display)?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn write_word_run_start<W: std::io::Write>(
    writer: &mut Writer<W>,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let mut run = BytesStart::new("w:r");
    if foreign_word_namespace.is_some() {
        run.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    writer.write_event(Event::Start(run))?;
    Ok(())
}

fn push_dirty_attribute(element: &mut BytesStart<'_>, dirty: Option<bool>) {
    if let Some(dirty) = dirty {
        element.push_attribute(("w:dirty", if dirty { "1" } else { "0" }));
    }
}

fn push_comment_marker(
    markers: &mut Vec<CommentRangeMarker>,
    extra_xml: &[(usize, Vec<u8>)],
    run_index: usize,
    id: i32,
    start: bool,
) {
    let raw_before = raw_xml_count_at(extra_xml, run_index);
    let marker = if start {
        CommentRangeMarker::Start {
            id,
            run_index,
            raw_before,
        }
    } else {
        CommentRangeMarker::End {
            id,
            run_index,
            raw_before,
        }
    };
    markers.push(marker);
}

fn raw_xml_count_at(extra_xml: &[(usize, Vec<u8>)], run_index: usize) -> usize {
    extra_xml
        .iter()
        .filter(|(position, raw)| *position == run_index && !is_xml_whitespace(raw))
        .count()
}

fn parse_hyperlink_children(
    raw: &[u8],
    word_prefixes: &[String],
) -> Result<ParsedHyperlinkChildren> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut runs = Vec::new();
    let mut run_sources = Vec::new();
    let mut revisions = Vec::new();
    let mut extra_xml = Vec::new();
    let mut revisions_at_boundary = 0usize;
    let mut buffer = Vec::new();
    let mut inside = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if !inside => {
                inside = true;
                let _ = word_prefixes_at(&element, word_prefixes)?;
            }
            Event::Start(element) => {
                let prefixes = word_prefixes_at(&element, word_prefixes)?;
                if is_word_element(element.name().as_ref(), b"r", &prefixes) {
                    let raw = capture_element(&mut reader, &element)?;
                    runs.push(parse_run_raw(&raw, &prefixes)?);
                    run_sources.push(Some(raw));
                    revisions_at_boundary = 0;
                } else if is_content_revision_element(element.name().as_ref(), &prefixes) {
                    let raw = capture_element(&mut reader, &element)?;
                    if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                        revisions.push((runs.len(), revision));
                        revisions_at_boundary += 1;
                    } else {
                        extra_xml.push((runs.len(), revisions_at_boundary, raw));
                    }
                } else {
                    extra_xml.push((
                        runs.len(),
                        revisions_at_boundary,
                        capture_element(&mut reader, &element)?,
                    ));
                }
            }
            Event::Empty(element) => {
                let prefixes = word_prefixes_at(&element, word_prefixes)?;
                let raw = capture_empty_element(&element)?;
                if is_content_revision_element(element.name().as_ref(), &prefixes) {
                    if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                        revisions.push((runs.len(), revision));
                        revisions_at_boundary += 1;
                    } else {
                        extra_xml.push((runs.len(), revisions_at_boundary, raw));
                    }
                } else {
                    extra_xml.push((runs.len(), revisions_at_boundary, raw));
                }
            }
            Event::End(element)
                if inside && matches_local_name(element.name().as_ref(), b"hyperlink") =>
            {
                break;
            }
            Event::Text(text) if is_xml_whitespace(text.as_ref()) => {
                extra_xml.push((runs.len(), revisions_at_boundary, text.as_ref().to_vec()));
            }
            Event::Comment(comment) => {
                extra_xml.push((
                    runs.len(),
                    revisions_at_boundary,
                    capture_standalone_event(Event::Comment(comment.into_owned()))?,
                ));
            }
            Event::PI(instruction) => {
                extra_xml.push((
                    runs.len(),
                    revisions_at_boundary,
                    capture_standalone_event(Event::PI(instruction.into_owned()))?,
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(ParsedHyperlinkChildren {
        runs,
        run_sources,
        revisions,
        extra_xml,
    })
}

fn parse_hyperlink_attributes(
    element: &BytesStart<'_>,
    scope: &[String],
) -> Result<ParsedHyperlinkAttributes> {
    let mut rel_id = None;
    let mut anchor = None;
    let mut tooltip = None;
    let mut doc_location = None;
    let mut extra = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        if attribute_in_namespace(name, b"id", R_NS, scope) {
            rel_id = Some(value);
        } else if attribute_in_namespace(name, b"anchor", crate::namespace::W_NS, scope) {
            anchor = Some(value);
        } else if attribute_in_namespace(name, b"tooltip", crate::namespace::W_NS, scope) {
            tooltip = Some(value);
        } else if attribute_in_namespace(name, b"docLocation", crate::namespace::W_NS, scope) {
            doc_location = Some(value);
        } else {
            extra.push((std::str::from_utf8(name)?.to_owned(), value));
        }
    }
    Ok((rel_id, anchor, tooltip, doc_location, extra))
}

fn attribute_in_namespace(name: &[u8], local: &[u8], namespace: &str, scope: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return false;
    };
    if name.get(separator + 1..) != Some(local) {
        return false;
    }
    let prefix = &name[..separator];
    let bound_namespace = scope.iter().find_map(|binding| {
        binding
            .strip_prefix('\0')
            .and_then(|binding| binding.split_once('\0'))
            .filter(|(candidate, _)| candidate.as_bytes() == prefix)
            .map(|(_, value)| value)
    });
    match bound_namespace {
        Some(value) => value == namespace,
        None if namespace == crate::namespace::W_NS => {
            scope.iter().any(|candidate| candidate.as_bytes() == prefix)
        }
        None => false,
    }
}

fn is_element_in_namespace(name: &[u8], local: &[u8], namespace: &str, scope: &[String]) -> bool {
    let (prefix, actual_local) = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or((&b""[..], name), |separator| {
            (&name[..separator], &name[separator + 1..])
        });
    actual_local == local
        && scope.iter().any(|binding| {
            binding
                .strip_prefix('\0')
                .and_then(|binding| binding.split_once('\0'))
                .is_some_and(|(candidate, value)| {
                    candidate.as_bytes() == prefix && value == namespace
                })
        })
}

fn element_prefix_has_binding(name: &[u8], scope: &[String]) -> bool {
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |separator| &name[..separator]);
    scope.iter().any(|binding| {
        binding
            .strip_prefix('\0')
            .and_then(|binding| binding.split_once('\0'))
            .is_some_and(|(candidate, _)| candidate.as_bytes() == prefix)
    })
}

fn prefixes_with_assumed_word_owner(name: &[u8], scope: &[String]) -> Result<Vec<String>> {
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |separator| &name[..separator]);
    let mut prefixes = scope.to_vec();
    prefixes.push(std::str::from_utf8(prefix)?.to_owned());
    Ok(prefixes)
}

fn is_content_revision_element(name: &[u8], word_prefixes: &[String]) -> bool {
    is_word_element(name, b"ins", word_prefixes)
        || is_word_element(name, b"del", word_prefixes)
        || is_word_element(name, b"moveFrom", word_prefixes)
        || is_word_element(name, b"moveTo", word_prefixes)
}

struct ParagraphBoundary<'a> {
    extra_xml: &'a [(usize, Vec<u8>)],
    content_controls: &'a [(usize, usize, usize, CT_Sdt)],
    markers: &'a [CommentRangeMarker],
    hyperlinks: &'a [HyperlinkSpan],
    revisions: &'a [(usize, usize, CT_Revision)],
    equations: &'a [(usize, usize, OfficeMath)],
}

fn write_paragraph_boundary<W: std::io::Write>(
    writer: &mut Writer<W>,
    boundary: ParagraphBoundary<'_>,
    run_index: usize,
) -> Result<()> {
    let extras = boundary
        .extra_xml
        .iter()
        .filter(|(position, _)| *position == run_index)
        .map(|(_, raw)| raw)
        .collect::<Vec<_>>();
    for raw_index in 0..=extras.len() {
        let boundary_markers = boundary
            .markers
            .iter()
            .filter(|marker| {
                marker.run_index() == run_index
                    && marker.raw_before().min(extras.len()) == raw_index
            })
            .collect::<Vec<_>>();
        for marker_index in 0..=boundary_markers.len() {
            for (_, _, _, sdt) in
                boundary
                    .content_controls
                    .iter()
                    .filter(|(at, raw_before, markers_before, _)| {
                        *at == run_index
                            && (*raw_before).min(extras.len()) == raw_index
                            && (*markers_before).min(boundary_markers.len()) == marker_index
                    })
            {
                sdt.to_xml(writer)?;
            }
            if let Some(marker) = boundary_markers.get(marker_index) {
                let (tag, id) = match marker {
                    CommentRangeMarker::Start { id, .. } => ("w:commentRangeStart", *id),
                    CommentRangeMarker::End { id, .. } => ("w:commentRangeEnd", *id),
                };
                let mut value = itoa::Buffer::new();
                let mut element = BytesStart::new(tag);
                element.push_attribute(("w:id", value.format(id)));
                writer.write_event(Event::Empty(element))?;
            }
        }
        if let Some(raw) = extras.get(raw_index) {
            if let Some((_, _, equation)) = boundary
                .equations
                .iter()
                .find(|(at, slot, _)| *at == run_index && *slot == raw_index)
            {
                equation.write_xml(writer)?;
            } else if let Some((hyperlink_index, hyperlink)) = boundary
                .hyperlinks
                .iter()
                .enumerate()
                .find(|(_, hyperlink)| {
                    hyperlink.run_start == run_index
                        && hyperlink.run_end == run_index
                        && hyperlink.preserved_raw_before == Some(raw_index)
                })
            {
                let mut replacement = Vec::new();
                let mut replacement_writer = Writer::new(&mut replacement);
                write_hyperlink_start(&mut replacement_writer, hyperlink)?;
                write_hyperlink_boundary(
                    &mut replacement_writer,
                    boundary.revisions,
                    hyperlink_index,
                    hyperlink,
                    run_index,
                )?;
                replacement_writer
                    .write_event(Event::End(BytesEnd::new(hyperlink_qname(hyperlink))))?;
                writer.get_mut().write_all(&replacement)?;
            } else {
                writer.get_mut().write_all(raw)?;
            }
        }
    }
    Ok(())
}

fn write_hyperlink_boundary<W: std::io::Write>(
    writer: &mut Writer<W>,
    revisions: &[(usize, usize, CT_Revision)],
    hyperlink_index: usize,
    hyperlink: &HyperlinkSpan,
    run_index: usize,
) -> Result<()> {
    let boundary_revisions = revisions
        .iter()
        .filter(|(at, slot, _)| {
            *at == run_index && hyperlink_revision_index(*slot) == Some(hyperlink_index)
        })
        .map(|(_, _, revision)| revision)
        .collect::<Vec<_>>();
    let relative_boundary = run_index.saturating_sub(hyperlink.run_start);
    for revision_index in 0..=boundary_revisions.len() {
        for (_, _, raw) in hyperlink.extra_xml.iter().filter(|(boundary, before, _)| {
            *boundary == relative_boundary
                && (*before).min(boundary_revisions.len()) == revision_index
        }) {
            writer.get_mut().write_all(raw)?;
        }
        if let Some(revision) = boundary_revisions.get(revision_index) {
            revision.write_xml(writer)?;
        }
    }
    Ok(())
}

fn write_hyperlink_start<W: std::io::Write>(
    writer: &mut Writer<W>,
    hyperlink: &HyperlinkSpan,
) -> Result<()> {
    let (word_prefix, declare_word) =
        safe_hyperlink_prefix(hyperlink, "w", "rdocxW", crate::namespace::W_NS);
    let (relationship_prefix, declare_relationship) =
        safe_hyperlink_prefix(hyperlink, "r", "rdocxR", R_NS);
    let mut element = BytesStart::new(format!("{word_prefix}:hyperlink"));
    let word_declaration = format!("xmlns:{word_prefix}");
    let relationship_declaration = format!("xmlns:{relationship_prefix}");
    let relationship_id = format!("{relationship_prefix}:id");
    let anchor_name = format!("{word_prefix}:anchor");
    if declare_word {
        element.push_attribute((word_declaration.as_str(), crate::namespace::W_NS));
    }
    if hyperlink.rel_id.is_some() && declare_relationship {
        element.push_attribute((relationship_declaration.as_str(), R_NS));
    }
    if let Some(rel_id) = &hyperlink.rel_id {
        element.push_attribute((relationship_id.as_str(), rel_id.as_str()));
    }
    if let Some(anchor) = &hyperlink.anchor {
        element.push_attribute((anchor_name.as_str(), anchor.as_str()));
    }
    if let Some(tooltip) = &hyperlink.tooltip {
        let tooltip_name = format!("{word_prefix}:tooltip");
        element.push_attribute((tooltip_name.as_str(), tooltip.as_str()));
    }
    if let Some(doc_location) = &hyperlink.doc_location {
        let doc_location_name = format!("{word_prefix}:docLocation");
        element.push_attribute((doc_location_name.as_str(), doc_location.as_str()));
    }
    for (name, value) in &hyperlink.extra_attributes {
        element.push_attribute((name.as_str(), value.as_str()));
    }
    writer.write_event(Event::Start(element))?;
    Ok(())
}

fn hyperlink_qname(hyperlink: &HyperlinkSpan) -> String {
    let (prefix, _) = safe_hyperlink_prefix(hyperlink, "w", "rdocxW", crate::namespace::W_NS);
    format!("{prefix}:hyperlink")
}

fn safe_hyperlink_prefix(
    hyperlink: &HyperlinkSpan,
    preferred: &str,
    fallback: &str,
    namespace: &str,
) -> (String, bool) {
    let preferred_declaration = format!("xmlns:{preferred}");
    match hyperlink
        .extra_attributes
        .iter()
        .find(|(name, _)| name == &preferred_declaration)
    {
        None => return (preferred.to_owned(), false),
        Some((_, value)) if value == namespace => return (preferred.to_owned(), false),
        Some(_) => {}
    }
    for suffix in 0usize.. {
        let candidate = if suffix == 0 {
            fallback.to_owned()
        } else {
            format!("{fallback}{suffix}")
        };
        let declaration = format!("xmlns:{candidate}");
        if !hyperlink
            .extra_attributes
            .iter()
            .any(|(name, _)| name == &declaration)
        {
            return (candidate, true);
        }
    }
    unreachable!("the finite attribute set cannot occupy every prefix")
}

fn shadowed_word_namespace(hyperlink: &HyperlinkSpan) -> Option<&str> {
    hyperlink
        .extra_attributes
        .iter()
        .find(|(name, value)| name == "xmlns:w" && value != crate::namespace::W_NS)
        .map(|(_, value)| value.as_str())
}

pub(crate) fn write_raw_with_word_override<W: std::io::Write>(
    writer: &mut Writer<W>,
    raw: &[u8],
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let Some(namespace) = foreign_word_namespace else {
        writer.get_mut().write_all(raw)?;
        return Ok(());
    };
    if !raw_uses_external_word_binding(raw) {
        writer.get_mut().write_all(raw)?;
        return Ok(());
    }
    writer
        .get_mut()
        .write_all(&raw_with_root_word_binding(raw, namespace)?)?;
    Ok(())
}

pub(crate) fn raw_with_external_bindings(
    raw: &[u8],
    external_bindings: &[(String, String)],
) -> Result<Vec<u8>> {
    if external_bindings.is_empty() {
        return Ok(raw.to_vec());
    }

    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut scopes = Vec::<Vec<(String, String)>>::new();
    let mut required = Vec::<String>::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let declarations = raw_namespace_declarations(&element)?;
                record_external_bindings_used(
                    &element,
                    &scopes,
                    &declarations,
                    external_bindings,
                    &mut required,
                );
                scopes.push(declarations);
            }
            Event::Empty(element) => {
                let declarations = raw_namespace_declarations(&element)?;
                record_external_bindings_used(
                    &element,
                    &scopes,
                    &declarations,
                    external_bindings,
                    &mut required,
                );
            }
            Event::End(_) => {
                scopes.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    if required.is_empty() {
        return Ok(raw.to_vec());
    }

    let bindings = external_bindings
        .iter()
        .filter(|(prefix, _)| required.contains(prefix))
        .cloned()
        .collect::<Vec<_>>();
    raw_with_root_bindings(raw, &bindings)
}

fn raw_namespace_declarations(element: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut declarations = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            ""
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            std::str::from_utf8(prefix)?
        } else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        declarations.push((prefix.to_owned(), value));
    }
    Ok(declarations)
}

fn record_external_bindings_used(
    element: &BytesStart<'_>,
    scopes: &[Vec<(String, String)>],
    declarations: &[(String, String)],
    external_bindings: &[(String, String)],
    required: &mut Vec<String>,
) {
    let element_name = element.name();
    let element_prefix = qualified_name_prefix(element_name.as_ref()).unwrap_or("");
    record_external_prefix_used(
        element_prefix,
        scopes,
        declarations,
        external_bindings,
        required,
    );
    for attribute in element.attributes().filter_map(|attribute| attribute.ok()) {
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if let Some(prefix) = qualified_name_prefix(name) {
            record_external_prefix_used(prefix, scopes, declarations, external_bindings, required);
        }
    }
}

fn record_external_prefix_used(
    prefix: &str,
    scopes: &[Vec<(String, String)>],
    declarations: &[(String, String)],
    external_bindings: &[(String, String)],
    required: &mut Vec<String>,
) {
    let internally_bound = declarations
        .iter()
        .rev()
        .any(|(candidate, _)| candidate == prefix)
        || scopes
            .iter()
            .rev()
            .any(|scope| scope.iter().rev().any(|(candidate, _)| candidate == prefix));
    if !internally_bound
        && external_bindings
            .iter()
            .any(|(candidate, _)| candidate == prefix)
        && !required.iter().any(|candidate| candidate == prefix)
    {
        required.push(prefix.to_owned());
    }
}

fn qualified_name_prefix(name: &[u8]) -> Option<&str> {
    let separator = name.iter().position(|byte| *byte == b':')?;
    std::str::from_utf8(&name[..separator]).ok()
}

fn raw_with_root_bindings(raw: &[u8], bindings: &[(String, String)]) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let end = reader.buffer_position() as usize;
                let mut element = element.into_owned();
                let names = bindings
                    .iter()
                    .map(|(prefix, _)| {
                        if prefix.is_empty() {
                            "xmlns".to_owned()
                        } else {
                            format!("xmlns:{prefix}")
                        }
                    })
                    .collect::<Vec<_>>();
                for ((_, namespace), name) in bindings.iter().zip(&names) {
                    element.push_attribute((name.as_str(), namespace.as_str()));
                }
                let mut output = raw[..start].to_vec();
                Writer::new(&mut output).write_event(Event::Start(element))?;
                output.extend_from_slice(&raw[end..]);
                return Ok(output);
            }
            Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let mut element = element.into_owned();
                let names = bindings
                    .iter()
                    .map(|(prefix, _)| {
                        if prefix.is_empty() {
                            "xmlns".to_owned()
                        } else {
                            format!("xmlns:{prefix}")
                        }
                    })
                    .collect::<Vec<_>>();
                for ((_, namespace), name) in bindings.iter().zip(&names) {
                    element.push_attribute((name.as_str(), namespace.as_str()));
                }
                let mut output = raw[..start].to_vec();
                Writer::new(&mut output).write_event(Event::Empty(element))?;
                output.extend_from_slice(&raw[end..]);
                return Ok(output);
            }
            Event::Eof => return Ok(raw.to_vec()),
            _ => {}
        }
        buffer.clear();
    }
}

fn raw_uses_external_word_binding(raw: &[u8]) -> bool {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut external_binding = vec![true];
    let mut saw_root = false;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                let (declares_word, uses_word) = raw_word_binding_use(&element);
                if !saw_root {
                    saw_root = true;
                    if declares_word {
                        return false;
                    }
                }
                let uses_external =
                    external_binding.last().copied().unwrap_or(true) && !declares_word;
                if uses_external && uses_word {
                    return true;
                }
                external_binding.push(uses_external);
            }
            Ok(Event::Empty(element)) => {
                let (declares_word, uses_word) = raw_word_binding_use(&element);
                if !saw_root {
                    saw_root = true;
                    if declares_word {
                        return false;
                    }
                }
                let uses_external =
                    external_binding.last().copied().unwrap_or(true) && !declares_word;
                if uses_external && uses_word {
                    return true;
                }
            }
            Ok(Event::End(_)) => {
                if external_binding.len() > 1 {
                    external_binding.pop();
                }
            }
            Ok(Event::Eof) => return false,
            Err(_) => return false,
            _ => {}
        }
        buffer.clear();
    }
}

fn raw_word_binding_use(element: &BytesStart<'_>) -> (bool, bool) {
    let declares_word = element
        .attributes()
        .filter_map(|attribute| attribute.ok())
        .any(|attribute| attribute.key.as_ref() == b"xmlns:w");
    let uses_word = qualified_name_uses_prefix(element.name().as_ref(), b"w")
        || element
            .attributes()
            .filter_map(|attribute| attribute.ok())
            .any(|attribute| qualified_name_uses_prefix(attribute.key.as_ref(), b"w"));
    (declares_word, uses_word)
}

fn qualified_name_uses_prefix(name: &[u8], prefix: &[u8]) -> bool {
    name.iter()
        .position(|byte| *byte == b':')
        .is_some_and(|separator| &name[..separator] == prefix)
}

fn raw_with_root_word_binding(raw: &[u8], namespace: &str) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let end = reader.buffer_position() as usize;
                let mut element = element.into_owned();
                element.push_attribute(("xmlns:w", namespace));
                let mut output = raw[..start].to_vec();
                Writer::new(&mut output).write_event(Event::Start(element))?;
                output.extend_from_slice(&raw[end..]);
                return Ok(output);
            }
            Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let mut element = element.into_owned();
                element.push_attribute(("xmlns:w", namespace));
                let mut output = raw[..start].to_vec();
                Writer::new(&mut output).write_event(Event::Empty(element))?;
                output.extend_from_slice(&raw[end..]);
                return Ok(output);
            }
            Event::Eof => return Ok(raw.to_vec()),
            _ => {}
        }
        buffer.clear();
    }
}

fn write_empty_hyperlinks<W: std::io::Write>(
    writer: &mut Writer<W>,
    hyperlinks: &[HyperlinkSpan],
    revisions: &[(usize, usize, CT_Revision)],
    run_index: usize,
) -> Result<()> {
    for (hyperlink_index, hyperlink) in hyperlinks.iter().enumerate().filter(|(_, hyperlink)| {
        hyperlink.run_start == run_index
            && hyperlink.run_end == run_index
            && hyperlink.preserved_raw_before.is_none()
    }) {
        write_hyperlink_start(writer, hyperlink)?;
        write_hyperlink_boundary(writer, revisions, hyperlink_index, hyperlink, run_index)?;
        writer.write_event(Event::End(BytesEnd::new(hyperlink_qname(hyperlink))))?;
    }
    Ok(())
}

fn required_word_i32_attribute(
    element: &BytesStart<'_>,
    local: &[u8],
    word_prefixes: &[String],
) -> Result<i32> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let Some(separator) = key.iter().position(|byte| *byte == b':') else {
            continue;
        };
        if key.get(separator + 1..) == Some(local)
            && word_prefixes
                .iter()
                .any(|prefix| prefix.as_bytes() == &key[..separator])
        {
            return Ok(std::str::from_utf8(&attribute.value)?.parse()?);
        }
    }
    Err(OxmlError::MissingElement(format!(
        "{} attribute",
        String::from_utf8_lossy(local)
    )))
}

fn optional_word_attribute(
    element: &BytesStart<'_>,
    local: &[u8],
    word_prefixes: &[String],
) -> Option<String> {
    for attribute in element.attributes().flatten() {
        let key = attribute.key.as_ref();
        let Some(separator) = key.iter().position(|byte| *byte == b':') else {
            continue;
        };
        if key.get(separator + 1..) == Some(local)
            && word_prefixes
                .iter()
                .any(|prefix| prefix.as_bytes() == &key[..separator])
        {
            let value = std::str::from_utf8(&attribute.value).ok()?;
            return Some(
                quick_xml::escape::unescape(value)
                    .map(|value| value.into_owned())
                    .unwrap_or_else(|_| value.to_owned()),
            );
        }
    }
    None
}

/// Parse a field instruction string into the shared recursive grammar.
fn parse_field_instruction(instr: &str) -> FieldInstruction {
    parse_field_instruction_parts(vec![InstructionPart::Text(instr.to_owned())])
}

#[derive(Debug)]
enum InstructionPart {
    Text(String),
    Nested(Field),
}

struct InstructionToken {
    argument: FieldArgument,
    quoted: bool,
}

fn parse_field_instruction_parts(parts: Vec<InstructionPart>) -> FieldInstruction {
    parse_field_instruction_parts_with_order(parts).0
}

fn parse_field_instruction_parts_with_order(
    parts: Vec<InstructionPart>,
) -> (FieldInstruction, Vec<NestedFieldPosition>) {
    let mut raw = String::new();
    let mut tokens = Vec::new();
    let mut text = String::new();
    for part in parts {
        match part {
            InstructionPart::Text(value) => {
                raw.push_str(&value);
                text.push_str(&value);
            }
            InstructionPart::Nested(field) => {
                tokens.extend(lex_field_text(&text));
                text.clear();
                tokens.push(InstructionToken {
                    argument: FieldArgument::Nested(Box::new(field)),
                    quoted: false,
                });
            }
        }
    }
    tokens.extend(lex_field_text(&text));

    let mut tokens = tokens.into_iter();
    let name = match tokens.next().map(|token| token.argument) {
        Some(FieldArgument::Text(value)) => value.to_uppercase(),
        Some(FieldArgument::Nested(_)) | None => String::new(),
    };
    let remaining = tokens.collect::<Vec<_>>();
    let mut arguments = Vec::new();
    let mut switches = Vec::new();
    let mut nested_order = Vec::new();
    let mut index = 0usize;
    while index < remaining.len() {
        let InstructionToken {
            argument: FieldArgument::Text(value),
            quoted,
        } = &remaining[index]
        else {
            if matches!(&remaining[index].argument, FieldArgument::Nested(_)) {
                nested_order.push(NestedFieldPosition::Argument(arguments.len()));
            }
            arguments.push(remaining[index].argument.clone());
            index += 1;
            continue;
        };
        let Some(switch_name) = (!quoted)
            .then(|| value.strip_prefix('\\'))
            .flatten()
            .filter(|name| !name.is_empty())
        else {
            if matches!(&remaining[index].argument, FieldArgument::Nested(_)) {
                nested_order.push(NestedFieldPosition::Argument(arguments.len()));
            }
            arguments.push(remaining[index].argument.clone());
            index += 1;
            continue;
        };
        let takes_argument = switch_takes_argument(&name, switch_name);
        let argument = if takes_argument && remaining.get(index + 1).is_some_and(|next| {
            next.quoted
                || !matches!(&next.argument, FieldArgument::Text(text) if text.starts_with('\\'))
        }) {
            index += 1;
            Some(remaining[index].argument.clone())
        } else {
            None
        };
        if matches!(&argument, Some(FieldArgument::Nested(_))) {
            nested_order.push(NestedFieldPosition::Switch(switches.len()));
        }
        switches.push(FieldSwitch {
            name: switch_name.to_ascii_lowercase(),
            argument,
        });
        index += 1;
    }

    (
        FieldInstruction {
            raw: raw.trim().to_owned(),
            name,
            arguments,
            switches,
        },
        nested_order,
    )
}

fn switch_takes_argument(field_name: &str, switch_name: &str) -> bool {
    matches!(
        switch_name.to_ascii_lowercase().as_str(),
        "*" | "#" | "@" | "r" | "s" | "d" | "b" | "f"
    ) || field_name == "INCLUDETEXT" && switch_name.eq_ignore_ascii_case("c")
}

fn lex_field_text(input: &str) -> Vec<InstructionToken> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut characters = input.chars().peekable();
    let mut quoted = false;
    let mut token_was_quoted = false;
    let push_token =
        |tokens: &mut Vec<InstructionToken>, current: &mut String, quoted: &mut bool| {
            if !current.is_empty() || *quoted {
                tokens.push(InstructionToken {
                    argument: FieldArgument::Text(std::mem::take(current)),
                    quoted: *quoted,
                });
                *quoted = false;
            }
        };
    while let Some(character) = characters.next() {
        if quoted {
            match character {
                '"' => quoted = false,
                '\\' if characters
                    .peek()
                    .is_some_and(|next| matches!(next, '"' | '\\')) =>
                {
                    current.push(characters.next().expect("peeked character exists"));
                }
                _ => current.push(character),
            }
        } else {
            match character {
                '"' => {
                    quoted = true;
                    token_was_quoted = true;
                }
                character if character.is_whitespace() => {
                    push_token(&mut tokens, &mut current, &mut token_was_quoted);
                }
                _ => current.push(character),
            }
        }
    }
    push_token(&mut tokens, &mut current, &mut token_was_quoted);
    tokens
}

impl Default for CT_P {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_paragraph(xml: &str) -> CT_P {
        let full = format!("<w:p>{xml}</w:p>");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"p") => break,
                _ => {}
            }
            buf.clear();
        }
        CT_P::from_xml(&mut reader).unwrap()
    }

    #[test]
    fn parse_simple_paragraph() {
        let p = parse_paragraph(r#"<w:r><w:t>Hello World</w:t></w:r>"#);
        assert_eq!(p.text(), "Hello World");
        assert_eq!(p.runs.len(), 1);
    }

    #[test]
    fn complex_hyperlink_field_exposes_its_target_and_cached_text() {
        let p = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve"> HYPERLINK &quot;https://example.test/path&quot; </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>Example link</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));

        assert_eq!(p.text(), "Example link");
        assert_eq!(
            p.complex_field_hyperlinks(),
            vec![ComplexFieldHyperlink {
                run_start: 0,
                run_end: 1,
                target: "https://example.test/path".to_owned(),
            }]
        );
    }

    #[test]
    fn form_text_field_imports_its_cached_plain_text() {
        let p = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText> FORMTEXT </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>Entered value</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));

        assert_eq!(p.text(), "Entered value");
        assert!(p.complex_field_hyperlinks().is_empty());
    }

    #[test]
    fn unsafe_complex_fields_are_not_exposed_as_links() {
        let cases = [
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText> HYPERLINK &quot;https://example.test&quot; </w:instrText></w:r>"#,
            ),
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>"#,
                r#"<w:r><w:instrText> HYPERLINK &quot;https://example.test&quot; </w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>Cached</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText> HYPERLINK </w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>Cached</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText> HYPERLINK &quot;https://example.test&quot; </w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>"#,
                r#"<w:r><w:t>Cached</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText> HYPERLINK &quot;https://example.test&quot; </w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
        ];

        for case in cases {
            let paragraph = parse_paragraph(case);
            assert!(paragraph.complex_field_hyperlinks().is_empty(), "{case}");
        }
    }

    #[test]
    fn parse_paragraph_with_properties() {
        let p = parse_paragraph(
            r#"<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r>"#,
        );
        assert_eq!(p.text(), "Centered");
        assert!(p.properties.is_some());
        assert_eq!(
            p.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn direct_paragraph_parser_accepts_explicit_property_binding() {
        let xml = format!(
            r#"<outer xmlns:ext="urn:producer"><q:p xmlns:q="{}"><ext:pPr><ext:jc ext:val="right"/></ext:pPr><q:pPr xmlns:q="{}"><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr><q:r><q:t>Direct</q:t></q:r></q:p></outer>"#,
            crate::namespace::W_NS,
            crate::namespace::W_NS
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"p" => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Ok(Event::Eof) => panic!("missing paragraph"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert_eq!(parsed.text(), "Direct");
        assert_eq!(
            parsed.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn direct_paragraph_parser_does_not_invent_foreign_word_identity() {
        let xml = r#"<outer><ext:p xmlns:ext="urn:producer"><ext:pPr xmlns:ext="urn:producer"><ext:jc ext:val="right"/></ext:pPr><ext:r><ext:t>Foreign</ext:t></ext:r></ext:p></outer>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"p" => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Ok(Event::Eof) => panic!("missing paragraph"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert!(parsed.properties.is_none());
    }

    #[test]
    fn direct_paragraph_parser_accepts_default_word_namespace() {
        let xml = format!(
            r#"<outer xmlns:ext="urn:producer"><p xmlns="{0}" xmlns:w="{0}"><ext:pPr><ext:jc ext:val="right"/></ext:pPr><pPr xmlns="{0}" xmlns:w="{0}"><ext:jc ext:val="right"/><jc w:val="center"/></pPr><r><t>Direct</t></r></p></outer>"#,
            crate::namespace::W_NS
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"p" => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Ok(Event::Eof) => panic!("missing paragraph"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert_eq!(parsed.text(), "Direct");
        assert_eq!(
            parsed.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn parse_run_with_formatting() {
        let p = parse_paragraph(r#"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold Italic</w:t></w:r>"#);
        let run = &p.runs[0];
        let rpr = run.properties.as_ref().unwrap();
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.italic, Some(true));
    }

    #[test]
    fn parse_multiple_runs() {
        let p = parse_paragraph(r#"<w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r>"#);
        assert_eq!(p.runs.len(), 2);
        assert_eq!(p.text(), "Hello World");
    }

    #[test]
    fn parse_hyperlink() {
        let p = parse_paragraph(
            r#"<w:hyperlink r:id="rId5"><w:r><w:t>Click here</w:t></w:r></w:hyperlink>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert_eq!(p.text(), "Click here");
        assert_eq!(p.hyperlinks.len(), 1);
        assert_eq!(p.hyperlinks[0].rel_id, Some("rId5".to_string()));
        assert_eq!(p.hyperlinks[0].run_start, 0);
        assert_eq!(p.hyperlinks[0].run_end, 1);
    }

    #[test]
    fn parse_hyperlink_with_anchor() {
        let p = parse_paragraph(
            r#"<w:hyperlink w:anchor="section1" w:tooltip="Jump" w:docLocation="target"><w:r><w:t>Go to section</w:t></w:r></w:hyperlink>"#,
        );
        assert_eq!(p.hyperlinks.len(), 1);
        assert_eq!(p.hyperlinks[0].anchor, Some("section1".to_string()));
        assert_eq!(p.hyperlinks[0].tooltip.as_deref(), Some("Jump"));
        assert_eq!(p.hyperlinks[0].doc_location.as_deref(), Some("target"));
        assert!(p.hyperlinks[0].rel_id.is_none());
    }

    #[test]
    fn parse_hyperlink_uses_local_namespace_aliases_without_reporting_declarations() {
        let p = parse_paragraph(concat!(
            r#"<q:hyperlink xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"q:tooltip="Jump" q:docLocation="target" q:history="1">"#,
            r#"<q:r><q:t>Go</q:t></q:r></q:hyperlink>"#,
        ));
        assert_eq!(p.hyperlinks[0].tooltip.as_deref(), Some("Jump"));
        assert_eq!(p.hyperlinks[0].doc_location.as_deref(), Some("target"));
        assert!(
            p.hyperlinks[0]
                .extra_attributes
                .contains(&("q:history".to_owned(), "1".to_owned()))
        );
        assert!(
            p.hyperlinks[0]
                .extra_attributes
                .iter()
                .any(|(name, _)| name == "xmlns:q")
        );
    }

    #[test]
    fn parse_hyperlink_multiple_runs() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Before </w:t></w:r><w:hyperlink r:id="rId6"><w:r><w:t>link </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r>"#,
        );
        assert_eq!(p.runs.len(), 4);
        assert_eq!(p.text(), "Before link text after");
        assert_eq!(p.hyperlinks.len(), 1);
        assert_eq!(p.hyperlinks[0].run_start, 1);
        assert_eq!(p.hyperlinks[0].run_end, 3);
    }

    #[test]
    fn round_trip_hyperlink() {
        let mut p = CT_P::new();
        p.add_run("Before ");
        p.add_run("link text");
        p.add_run(" after");
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: Some("rId7".to_string()),
            anchor: None,
            tooltip: None,
            doc_location: None,
            run_start: 1,
            run_end: 2,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.text(), "Before link text after");
        assert_eq!(parsed.hyperlinks.len(), 1);
        assert_eq!(parsed.hyperlinks[0].rel_id, Some("rId7".to_string()));
        assert_eq!(parsed.hyperlinks[0].run_start, 1);
        assert_eq!(parsed.hyperlinks[0].run_end, 2);
    }

    #[test]
    fn comment_reference_is_typed_and_round_trips() {
        let p = parse_paragraph(
            r#"<w:r><w:t>flagged</w:t></w:r><w:r><w:commentReference w:id="1"/></w:r>"#,
        );
        assert_eq!(p.runs.len(), 2);
        assert!(matches!(
            p.runs[1].content.as_slice(),
            [RunContent::CommentReference { id: 1, .. }]
        ));

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();
        assert!(
            xml.contains(r#"<w:commentReference w:id="1"/>"#),
            "comment reference must survive round-trip: {xml}"
        );
        assert!(xml.contains("flagged"));
    }

    #[test]
    fn alternate_content_is_preserved_once_and_parsed_for_layout() {
        // A shape as Word writes it: DrawingML in mc:Choice, VML in the
        // fallback. The block has to come back out verbatim, exactly once,
        // while still being visible to layout.
        let src = concat!(
            r#"<w:r><mc:AlternateContent><mc:Choice Requires="wps">"#,
            r#"<w:drawing><wp:anchor behindDoc="0">"#,
            r#"<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>"#,
            r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
            r#"<wp:extent cx="914400" cy="457200"/>"#,
            r#"<a:graphic><a:graphicData><wps:wsp><wps:spPr>"#,
            r#"<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="729FCF"/></a:solidFill>"#,
            r#"<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>"#,
            r#"</wps:spPr><wps:txbx><w:txbxContent>"#,
            r#"<w:p><w:r><w:t>boxed</w:t></w:r></w:p>"#,
            r#"</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic>"#,
            r#"</wp:anchor></w:drawing></mc:Choice>"#,
            r#"<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback>"#,
            r#"</mc:AlternateContent></w:r>"#,
        );
        let p = parse_paragraph(src);
        assert_eq!(p.runs.len(), 1);
        let run = &p.runs[0];

        // Visible to layout.
        assert_eq!(
            run.alt_drawings.len(),
            1,
            "the drawing must reach the model"
        );
        let anchor = run.alt_drawings[0]
            .anchor
            .as_ref()
            .expect("should be an anchored drawing");
        let shape = anchor.shape.as_ref().expect("should carry shape content");
        assert_eq!(shape.preset.as_deref(), Some("rect"));
        assert_eq!(
            shape.solid_fill.as_deref(),
            Some("729FCF"),
            "the fill colour must win over the outline colour"
        );
        assert_eq!(shape.text.len(), 1);
        assert_eq!(shape.text[0].text(), "boxed");

        // Preserved verbatim, exactly once.
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();
        assert_eq!(
            xml.matches("<mc:AlternateContent").count(),
            1,
            "the block must not be duplicated: {xml}"
        );
        assert_eq!(
            xml.matches("<mc:Fallback").count(),
            1,
            "the VML fallback must survive"
        );
        assert!(xml.contains(r#"<a:prstGeom prst="rect"/>"#));
    }

    #[test]
    fn vml_only_alternate_content_is_preserved_without_a_layout_projection() {
        let p = parse_paragraph(concat!(
            r#"<w:r><mc:AlternateContent><mc:Choice Requires="vml">"#,
            r#"<w:pict><v:shape id="legacy"/></w:pict>"#,
            r#"</mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"/></w:pict>"#,
            r#"</mc:Fallback></mc:AlternateContent></w:r>"#,
        ));

        assert!(p.runs[0].alt_drawings.is_empty());
        let mut output = Vec::new();
        p.to_xml(&mut Writer::new(&mut output))
            .expect("paragraph writes");
        let xml = String::from_utf8(output).expect("XML is UTF-8");
        assert_eq!(xml.matches("<mc:AlternateContent").count(), 1, "{xml}");
        assert!(xml.contains(r#"<v:shape id="legacy"/>"#), "{xml}");
    }

    #[test]
    fn empty_run_properties_are_not_moved_after_content() {
        // extra_xml is written after the run content, and CT_R requires
        // w:rPr first, so a self-closing <w:rPr/> must not be captured.
        // Capturing it would emit <w:t> before <w:rPr/> and break the schema.
        let p = parse_paragraph(r#"<w:r><w:rPr/><w:t>x</w:t></w:r>"#);
        assert_eq!(p.runs.len(), 1);
        assert!(
            p.runs[0].extra_xml.is_empty(),
            "an empty w:rPr must not be captured as extra_xml"
        );

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();
        assert!(
            !xml.contains(r#"<w:t>x</w:t><w:rPr/>"#),
            "w:rPr must never follow run content: {xml}"
        );
    }

    #[test]
    fn parse_fld_simple_page() {
        let p = parse_paragraph(
            r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert_eq!(p.runs[0].content.len(), 1);
        assert_eq!(parsed_field(&p, 0).instruction.name, "PAGE");
    }

    #[test]
    fn parse_fld_simple_numpages() {
        let p = parse_paragraph(
            r#"<w:fldSimple w:instr=" NUMPAGES \* MERGEFORMAT "><w:r><w:t>5</w:t></w:r></w:fldSimple>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert_eq!(parsed_field(&p, 0).instruction.name, "NUMPAGES");
    }

    fn parsed_field(paragraph: &CT_P, run_index: usize) -> &Field {
        match &paragraph.runs[run_index].content[0] {
            RunContent::Field(field) => field,
            other => panic!("expected field, got {other:?}"),
        }
    }

    fn serialized_paragraph(paragraph: &CT_P) -> String {
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        String::from_utf8(output).unwrap()
    }

    #[test]
    fn field_instruction_corpus_parses_every_simple_complex_split_and_nested_form() {
        let corpus = [
            (r"PAGE \* MERGEFORMAT", "PAGE"),
            ("NUMPAGES \\# \"0\"", "NUMPAGES"),
            ("REF \"destination\" \\h", "REF"),
            (r"PAGEREF destination \p", "PAGEREF"),
            (r"SEQ Figure \r 1", "SEQ"),
            ("DOCPROPERTY \"Last Saved By\" \\* Upper", "DOCPROPERTY"),
            ("DOCVARIABLE \"Customer Name\"", "DOCVARIABLE"),
            ("STYLEREF \"Heading 1\" \\l", "STYLEREF"),
            ("INCLUDETEXT \"chapter one.docx\" bookmark", "INCLUDETEXT"),
            ("DATE \\@ \"MMMM d, yyyy\"", "DATE"),
            ("TIME \\@ \"HH:mm\"", "TIME"),
            (r"FILENAME \p", "FILENAME"),
            (r"AUTHOR \* Caps", "AUTHOR"),
            ("MERGEFIELD \"First Name\" \\* MERGEFORMAT", "MERGEFIELD"),
            ("IF \"10\" >= \"2\" \"yes\" \"no\"", "IF"),
        ];

        for (instruction, expected_name) in corpus {
            let escaped = instruction.replace('&', "&amp;").replace('"', "&quot;");
            let paragraph = parse_paragraph(&format!(
                r#"<w:fldSimple w:instr="{escaped}"><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
            ));
            let field = parsed_field(&paragraph, 0);
            assert_eq!(field.instruction.raw, instruction, "{instruction}");
            assert_eq!(field.instruction.name, expected_name, "{instruction}");
            assert_eq!(field.cached_result, "cached", "{instruction}");
            let expected_switch = match expected_name {
                "PAGE" | "DOCPROPERTY" | "AUTHOR" | "MERGEFIELD" => Some("*"),
                "NUMPAGES" => Some("#"),
                "REF" => Some("h"),
                "PAGEREF" | "FILENAME" => Some("p"),
                "SEQ" => Some("r"),
                "STYLEREF" => Some("l"),
                "DATE" | "TIME" => Some("@"),
                _ => None,
            };
            if let Some(expected_switch) = expected_switch {
                assert!(
                    field
                        .instruction
                        .switches
                        .iter()
                        .any(|switch| switch.name == expected_switch),
                    "{instruction}"
                );
            }
            if expected_name == "IF" {
                assert_eq!(
                    field.instruction.arguments,
                    ["10", ">=", "2", "yes", "no"]
                        .map(|value| FieldArgument::Text(value.to_owned()))
                );
            }
        }

        let escaped = parse_paragraph(
            r#"<w:fldSimple w:instr="MERGEFIELD &quot;Customer \&quot;Code\&quot;&quot; \* MERGEFORMAT"><w:r><w:t>value</w:t></w:r></w:fldSimple>"#,
        );
        let escaped = parsed_field(&escaped, 0);
        assert!(matches!(
            escaped.instruction.arguments.first(),
            Some(FieldArgument::Text(value)) if value == "Customer \"Code\""
        ));
        assert!(escaped
            .instruction
            .switches
            .iter()
            .any(|switch| switch.name == "*"
                && matches!(&switch.argument, Some(FieldArgument::Text(value)) if value == "MERGEFORMAT")));

        let complex = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve"> IF 1 = </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>inside</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve"> &quot;yes&quot; &quot;no&quot; </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>yes</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let complex = parsed_field(&complex, 0);
        assert_eq!(complex.instruction.name, "IF");
        assert_eq!(complex.cached_result, "yes");
        assert_eq!(complex.dirty, Some(true));
        assert!(
            complex
                .instruction
                .arguments
                .iter()
                .any(|argument| matches!(argument, FieldArgument::Nested(field)
                if field.instruction.name == "REF" && field.cached_result == "inside"))
        );
    }

    #[test]
    fn same_run_complex_markers_are_safe_and_keep_the_cached_result() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/>"#,
            r#"<w:instrText>DATE</w:instrText>"#,
            r#"<w:fldChar w:fldCharType="separate"/>"#,
            r#"<w:t>17 August 2026</w:t>"#,
            r#"<w:fldChar w:fldCharType="end"/></w:r>"#,
        ));

        let field = parsed_field(&paragraph, 0);
        assert_eq!(field.instruction.name, "DATE");
        assert_eq!(field.cached_result, "17 August 2026");
    }

    #[test]
    fn field_discovery_uses_direct_run_children_and_restores_namespace_scope() {
        let word_namespace = crate::namespace::W_NS;
        let nested_marker = parse_paragraph(concat!(
            r#"<w:r><x:extension xmlns:x="urn:producer"><w:fldChar w:fldCharType="begin"/></x:extension></w:r>"#,
            r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>literal</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        assert!(!nested_marker.runs.iter().any(|run| {
            run.content
                .iter()
                .any(|content| matches!(content, RunContent::Field(_)))
        }));
        assert_eq!(nested_marker.text(), "literal");

        let shadowed = parse_paragraph(&format!(
            r#"<q:r xmlns:q="{}"><q:fldChar q:fldCharType="begin"/><x:extension xmlns:x="urn:producer" xmlns:q="urn:not-word"/><q:instrText>DATE</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>cached</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
            word_namespace,
        ));
        let field = parsed_field(&shadowed, 0);
        assert_eq!(field.instruction.name, "DATE");
        assert_eq!(field.cached_result, "cached");
    }

    #[test]
    fn quoted_backslash_and_empty_operands_remain_arguments() {
        let paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="INCLUDETEXT &quot;\\\\server\\share\\file.docx&quot; &quot;&quot;"><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
        );
        let field = parsed_field(&paragraph, 0);
        assert_eq!(
            field.instruction.arguments,
            [r#"\\server\share\file.docx"#, ""].map(|value| FieldArgument::Text(value.to_owned()))
        );
        assert!(field.instruction.switches.is_empty());
    }

    #[test]
    fn mergefield_and_includetext_switches_own_their_arguments() {
        let merge = parse_paragraph(
            r#"<w:fldSimple w:instr="MERGEFIELD Name \b &quot;before text&quot; \f &quot;after text&quot;"><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
        );
        let merge = parsed_field(&merge, 0);
        assert_eq!(
            merge.instruction.arguments,
            vec![FieldArgument::Text("Name".to_owned())]
        );
        assert_eq!(
            merge.instruction.switches,
            vec![
                FieldSwitch {
                    name: "b".to_owned(),
                    argument: Some(FieldArgument::Text("before text".to_owned())),
                },
                FieldSwitch {
                    name: "f".to_owned(),
                    argument: Some(FieldArgument::Text("after text".to_owned())),
                },
            ]
        );

        let include = parse_paragraph(
            r#"<w:fldSimple w:instr="INCLUDETEXT &quot;chapter.docx&quot; \c &quot;MSWord8&quot;"><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
        );
        let include = parsed_field(&include, 0);
        assert_eq!(
            include.instruction.arguments,
            vec![FieldArgument::Text("chapter.docx".to_owned())]
        );
        assert_eq!(
            include.instruction.switches,
            vec![FieldSwitch {
                name: "c".to_owned(),
                argument: Some(FieldArgument::Text("MSWord8".to_owned())),
            }]
        );
    }

    #[test]
    fn malformed_simple_fields_remain_opaque_and_byte_identical() {
        for source in [
            r#"<w:fldSimple><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
            r#"<w:fldSimple w:instr=""><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
        ] {
            let paragraph = parse_paragraph(source);
            assert!(!paragraph.runs.iter().any(|run| {
                run.content
                    .iter()
                    .any(|content| matches!(content, RunContent::Field(_)))
            }));
            assert_eq!(
                serialized_paragraph(&paragraph),
                format!("<w:p>{source}</w:p>")
            );
        }
    }

    #[test]
    fn field_mutation_preserves_source_runs_formatting_and_unmodelled_xml() {
        let word_namespace = crate::namespace::W_NS;
        let mut simple = parse_paragraph(&format!(
            r#"<q:fldSimple xmlns:q="{word_namespace}" xmlns:x="urn:producer" q:instr="DATE" q:dirty="1"><q:r data="one"><q:rPr><q:b/></q:rPr><q:t>old</q:t><x:custom><x:item/></x:custom></q:r><q:r data="two"><q:rPr><q:i/></q:rPr><q:t> result</q:t></q:r></q:fldSimple>"#,
        ));
        let RunContent::Field(field) = &mut simple.runs[0].content[0] else {
            panic!("expected simple field")
        };
        field.cached_result = "new result".to_owned();
        field.dirty = Some(false);
        let output = serialized_paragraph(&simple);
        assert!(
            output.contains(r#"<x:custom><x:item/></x:custom>"#),
            "{output}"
        );
        assert!(output.contains(r#"<q:rPr><q:b/></q:rPr>"#), "{output}");
        assert!(
            output.contains(r#"<q:r data="two"><q:rPr><q:i/></q:rPr>"#),
            "{output}"
        );
        assert_eq!(output.matches("<q:r ").count(), 2, "{output}");
        assert!(output.contains("new result"), "{output}");
        assert!(output.contains(r#"w:dirty="0""#), "{output}");

        let mut complex = parse_paragraph(concat!(
            r#"<w:r data="begin"><w:fldChar w:fldCharType="begin" w:dirty="1"/><w:producerBegin/></w:r>"#,
            r#"<w:r data="instruction"><w:rPr><w:b/></w:rPr><w:instrText>DATE</w:instrText><w:producerInstruction/></w:r>"#,
            r#"<w:proofErr w:type="spellStart"/>"#,
            r#"<w:r data="separator"><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r data="result-one"><w:rPr><w:i/></w:rPr><w:t>old</w:t><w:producerResult/></w:r>"#,
            r#"<w:r data="result-two"><w:rPr><w:b/></w:rPr><w:t> result</w:t></w:r>"#,
            r#"<w:r data="end"><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let RunContent::Field(field) = &mut complex.runs[0].content[0] else {
            panic!("expected complex field")
        };
        field.cached_result = "new result".to_owned();
        field.dirty = Some(false);
        let output = serialized_paragraph(&complex);
        for preserved in [
            r#"<w:producerBegin/>"#,
            r#"<w:producerInstruction/>"#,
            r#"<w:proofErr w:type="spellStart"/>"#,
            r#"<w:rPr><w:i/></w:rPr>"#,
            r#"<w:producerResult/>"#,
            r#"<w:r data="result-two"><w:rPr><w:b/></w:rPr>"#,
        ] {
            assert!(output.contains(preserved), "missing {preserved}: {output}");
        }
        assert_eq!(output.matches("<w:r ").count(), 6, "{output}");
        assert!(output.contains("new result"), "{output}");
        assert!(output.contains(r#"w:dirty="0""#), "{output}");
    }

    #[test]
    fn nested_cache_mutation_preserves_operand_order_and_producer_xml() {
        let word_namespace = crate::namespace::W_NS;
        let mut paragraph = parse_paragraph(&format!(
            concat!(
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer" data="outer-begin"><q:fldChar q:fldCharType="begin"><x:outerBegin/></q:fldChar></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer" data="outer-instruction"><q:rPr><q:b/></q:rPr><q:instrText xml:space="preserve">IF </q:instrText><x:outerInstruction/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer" data="nested-begin"><q:fldChar q:fldCharType="begin" q:dirty="on"><x:nestedBegin/></q:fldChar></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer" data="nested-instruction"><q:rPr><q:i/></q:rPr><q:instrText>REF destination</q:instrText><x:nestedInstruction/></q:r>"#,
                r#"<q:r xmlns:q="{0}" data="nested-separator"><q:fldChar q:fldCharType="separate"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer" data="nested-result"><q:rPr><q:u/></q:rPr><q:t>old</q:t><x:nestedResult/></q:r>"#,
                r#"<q:r xmlns:q="{0}" data="nested-end"><q:fldChar q:fldCharType="end"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer" data="outer-tail"><q:instrText xml:space="preserve"> = &quot;x&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><x:outerTail/></q:r>"#,
                r#"<q:r xmlns:q="{0}" data="outer-separator"><q:fldChar q:fldCharType="separate"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" data="outer-result"><q:t>yes</q:t></q:r>"#,
                r#"<q:r xmlns:q="{0}" data="outer-end"><q:fldChar q:fldCharType="end"/></q:r>"#,
            ),
            word_namespace,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected outer field")
        };
        let Some(FieldArgument::Nested(nested)) = field.instruction.arguments.first_mut() else {
            panic!("expected first positional operand to be nested")
        };
        nested.cached_result = "x".to_owned();
        nested.dirty = Some(false);

        let output = serialized_paragraph(&paragraph);
        for preserved in [
            r#"data="outer-begin""#,
            r#"<x:outerBegin/>"#,
            r#"<q:rPr><q:b/></q:rPr>"#,
            r#"<x:outerInstruction/>"#,
            r#"data="nested-instruction""#,
            r#"<q:rPr><q:i/></q:rPr>"#,
            r#"<x:nestedInstruction/>"#,
            r#"<q:rPr><q:u/></q:rPr>"#,
            r#"<x:nestedResult/>"#,
            r#"<x:outerTail/>"#,
        ] {
            assert!(output.contains(preserved), "missing {preserved}: {output}");
        }
        assert!(output.contains(r#"w:dirty="0""#), "{output}");
        assert!(output.contains(">x</w:t>"), "{output}");
        let nested_begin = output.find(r#"data="nested-begin""#).unwrap();
        let comparison = output.find(" = &quot;x&quot;").unwrap();
        assert!(nested_begin < comparison, "{output}");

        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        let field = parsed_field(&reparsed, 0);
        assert!(matches!(
            field.instruction.arguments.as_slice(),
            [FieldArgument::Nested(nested), FieldArgument::Text(operator), FieldArgument::Text(right), FieldArgument::Text(yes), FieldArgument::Text(no)]
                if nested.cached_result == "x"
                    && operator == "="
                    && right == "x"
                    && yes == "yes"
                    && no == "no"
        ));
    }

    #[test]
    fn same_instruction_nested_replacements_use_canonical_source_identity() {
        let outer_source = concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve">IF </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>old nested</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve"> = &quot;x&quot; &quot;yes&quot; &quot;no&quot;</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>old outer</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        );
        let word_namespace = crate::namespace::W_NS;
        let foreign = parse_paragraph(&format!(
            r#"<q:fldSimple xmlns:q="{word_namespace}" xmlns:x="urn:foreign" q:instr="REF destination"><q:r><q:t>foreign replacement</q:t><x:foreign/></q:r></q:fldSimple>"#,
        ));
        let replacements = [
            Field::new("REF destination", "new replacement"),
            parsed_field(&foreign, 0).clone(),
        ];

        for (replacement, expected) in replacements
            .into_iter()
            .zip(["new replacement", "foreign replacement"])
        {
            let mut paragraph = parse_paragraph(outer_source);
            let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
                panic!("expected outer field")
            };
            let Some(FieldArgument::Nested(nested)) = outer.instruction.arguments.first_mut()
            else {
                panic!("expected nested positional field")
            };
            **nested = replacement;

            let output = serialized_paragraph(&paragraph);
            let reparsed = parse_paragraph(
                output
                    .strip_prefix("<w:p>")
                    .and_then(|value| value.strip_suffix("</w:p>"))
                    .unwrap(),
            );
            let field = parsed_field(&reparsed, 0);
            let Some(FieldArgument::Nested(nested)) = field.instruction.arguments.first() else {
                panic!("replacement field was discarded: {output}")
            };
            assert_eq!(nested.cached_result, expected, "{output}");
        }
    }

    #[test]
    fn nested_updates_skip_opaque_lookalikes_and_map_identical_siblings() {
        let word_namespace = crate::namespace::W_NS;
        let nested_source = format!(
            concat!(
                r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="begin"/></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:instrText>REF destination</q:instrText></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="separate"/></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:t>stored nested</q:t></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="end"/></q:r>"#,
            ),
            word_namespace,
        );
        let source = format!(
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r xmlns:x="urn:producer"><w:instrText xml:space="preserve">IF </w:instrText><x:opaque>{0}</x:opaque></w:r>"#,
                "{0}",
                r#"<w:r><w:instrText xml:space="preserve"> </w:instrText></w:r>"#,
                "{0}",
                r#"<w:r><w:instrText xml:space="preserve"> = &quot;x&quot; &quot;yes&quot; &quot;no&quot;</w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>stored outer</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            nested_source,
        );
        let mut paragraph = parse_paragraph(&source);
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            panic!("expected outer field")
        };
        let mut nested =
            outer
                .instruction
                .arguments
                .iter_mut()
                .filter_map(|argument| match argument {
                    FieldArgument::Nested(field) => Some(field.as_mut()),
                    FieldArgument::Text(_) => None,
                });
        nested.next().unwrap().cached_result = "first typed".to_owned();
        nested.next().unwrap().cached_result = "second typed".to_owned();
        assert!(nested.next().is_none());

        let output = serialized_paragraph(&paragraph);
        assert!(
            output.contains(&format!("<x:opaque>{nested_source}</x:opaque>")),
            "{output}"
        );
        assert_eq!(output.matches("stored nested").count(), 1, "{output}");
        assert_eq!(output.matches("first typed").count(), 1, "{output}");
        assert_eq!(output.matches("second typed").count(), 1, "{output}");
        assert!(
            output.find("stored nested").unwrap() < output.find("first typed").unwrap()
                && output.find("first typed").unwrap() < output.find("second typed").unwrap(),
            "{output}"
        );
    }

    #[test]
    fn same_run_nested_siblings_have_distinct_update_spans() {
        let word_namespace = crate::namespace::W_NS;
        let nested_source = concat!(
            r#"<q:fldChar q:fldCharType="begin"/>"#,
            r#"<q:instrText>MERGEFIELD Same</q:instrText>"#,
            r#"<q:fldChar q:fldCharType="separate"/>"#,
            r#"<q:t>stored sibling</q:t>"#,
            r#"<q:fldChar q:fldCharType="end"/>"#,
        );
        let source = format!(
            concat!(
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer"><x:before/>"#,
                r#"<q:fldChar q:fldCharType="begin"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
                "{1}",
                r#"<q:instrText xml:space="preserve"> </q:instrText>"#,
                "{1}",
                r#"<q:instrText xml:space="preserve"> = &quot;x&quot; &quot;yes&quot; &quot;no&quot;</q:instrText>"#,
                r#"<q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t>"#,
                r#"<q:fldChar q:fldCharType="end"/><x:after/></q:r>"#,
            ),
            word_namespace, nested_source,
        );
        let mut paragraph = parse_paragraph(&source);
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            panic!("expected outer field")
        };
        let mut nested =
            outer
                .instruction
                .arguments
                .iter_mut()
                .filter_map(|argument| match argument {
                    FieldArgument::Nested(field) => Some(field.as_mut()),
                    FieldArgument::Text(_) => None,
                });
        nested.next().unwrap().cached_result = "first same run".to_owned();
        nested.next().unwrap().cached_result = "second same run".to_owned();
        assert!(nested.next().is_none());

        let output = serialized_paragraph(&paragraph);
        assert_eq!(output.matches("first same run").count(), 1, "{output}");
        assert_eq!(output.matches("second same run").count(), 1, "{output}");
        assert!(!output.contains("stored sibling"), "{output}");
        assert!(output.contains("<x:before/>"), "{output}");
        assert!(output.contains("<x:after/>"), "{output}");
        assert!(
            output.find("first same run").unwrap() < output.find("second same run").unwrap(),
            "{output}"
        );
    }

    #[test]
    fn shared_boundary_run_nested_siblings_have_non_overlapping_spans() {
        let word_namespace = crate::namespace::W_NS;
        let source = format!(
            concat!(
                r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="begin"/><q:instrText xml:space="preserve">IF </q:instrText></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="begin"/><q:instrText>MERGEFIELD First</q:instrText></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer"><q:fldChar q:fldCharType="separate"/><q:t>stored first</q:t><q:fldChar q:fldCharType="end"/><x:between/><q:instrText xml:space="preserve"> = </q:instrText><q:fldChar q:fldCharType="begin"/></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:instrText>MERGEFIELD Second</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored second</q:t><q:fldChar q:fldCharType="end"/><q:instrText xml:space="preserve"> &quot;yes&quot; &quot;no&quot;</q:instrText></q:r>"#,
                r#"<q:r xmlns:q="{0}"><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
            ),
            word_namespace,
        );
        let mut paragraph = parse_paragraph(&source);
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            panic!("expected outer field")
        };
        let mut nested =
            outer
                .instruction
                .arguments
                .iter_mut()
                .filter_map(|argument| match argument {
                    FieldArgument::Nested(field) => Some(field.as_mut()),
                    FieldArgument::Text(_) => None,
                });
        nested.next().unwrap().cached_result = "updated first".to_owned();
        nested.next().unwrap().cached_result = "updated second".to_owned();
        assert!(nested.next().is_none());

        let output = serialized_paragraph(&paragraph);
        assert_eq!(output.matches("updated first").count(), 1, "{output}");
        assert_eq!(output.matches("updated second").count(), 1, "{output}");
        assert!(!output.contains("stored first"), "{output}");
        assert!(!output.contains("stored second"), "{output}");
        assert!(output.contains("<x:between/>"), "{output}");
        assert!(
            output.find("updated first").unwrap() < output.find("<x:between/>").unwrap()
                && output.find("<x:between/>").unwrap() < output.find("updated second").unwrap(),
            "{output}"
        );
    }

    #[test]
    fn isolated_same_run_nested_field_applies_its_raw_instruction_edit() {
        let word_namespace = crate::namespace::W_NS;
        let source = format!(
            concat!(
                r#"<q:r xmlns:q="{0}" xmlns:x="urn:producer"><q:fldChar q:fldCharType="begin"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText>MERGEFIELD Old</q:instrText><x:nestedInstruction/><q:fldChar q:fldCharType="separate"/><q:t>stored nested</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
            ),
            word_namespace,
        );
        let mut paragraph = parse_paragraph(&source);
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            panic!("expected outer field")
        };
        let Some(FieldArgument::Nested(nested)) = outer.instruction.arguments.first_mut() else {
            panic!("expected nested field")
        };
        nested.instruction.raw = "MERGEFIELD New".to_owned();
        nested.cached_result = "new value".to_owned();
        nested.dirty = Some(false);

        let output = serialized_paragraph(&paragraph);
        assert!(output.contains("MERGEFIELD New"), "{output}");
        assert!(!output.contains("MERGEFIELD Old"), "{output}");
        assert!(output.contains("new value"), "{output}");
        assert!(output.contains(r#"w:dirty="0""#), "{output}");
        assert!(output.contains("<x:nestedInstruction/>"), "{output}");

        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        let outer = parsed_field(&reparsed, 0);
        let Some(FieldArgument::Nested(nested)) = outer.instruction.arguments.first() else {
            panic!("expected reopened nested field")
        };
        assert_eq!(nested.instruction.name, "MERGEFIELD");
        assert!(matches!(
            nested.instruction.arguments.first(),
            Some(FieldArgument::Text(value)) if value == "New"
        ));
        assert_eq!(nested.cached_result, "new value");
        assert_eq!(nested.dirty, Some(false));
    }

    #[test]
    fn complex_source_retains_comments_and_processing_instructions() {
        let source = concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            "<!-- before instruction -->",
            r#"<w:r><w:instrText>MERGEFIELD Producer</w:instrText></w:r>"#,
            "<?producer before-separator?>",
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            "<!-- before result -->",
            r#"<w:r><w:t>stored producer</w:t></w:r>"#,
            "<?producer before-end?>",
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        );
        let mut paragraph = parse_paragraph(source);
        let field = parsed_field(&paragraph, 0);
        let (captured, _) = field.source_replacement().unwrap().unwrap();
        assert_eq!(captured, source.as_bytes());

        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "updated producer".to_owned();
        field.dirty = Some(false);
        let output = serialized_paragraph(&paragraph);
        for preserved in [
            "<!-- before instruction -->",
            "<?producer before-separator?>",
            "<!-- before result -->",
            "<?producer before-end?>",
        ] {
            assert!(output.contains(preserved), "missing {preserved}: {output}");
        }
        assert!(output.contains("updated producer"), "{output}");
    }

    #[test]
    fn pretty_printed_complex_source_retains_inter_run_whitespace() {
        let source = concat!(
            "\n    ",
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            "\n    ",
            r#"<w:r><w:instrText>MERGEFIELD Pretty</w:instrText></w:r>"#,
            "\n    ",
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            "\n    ",
            r#"<w:r><w:t>stored pretty</w:t></w:r>"#,
            "\n    ",
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            "\n  ",
        );
        let paragraph = parse_paragraph(source);
        let field = parsed_field(&paragraph, 0);
        let (captured, _) = field.source_replacement().unwrap().unwrap();
        assert!(
            captured
                .windows(b"</w:r>\n    <w:r>".len())
                .any(|window| window == b"</w:r>\n    <w:r>"),
            "{}",
            String::from_utf8_lossy(captured)
        );
    }

    #[test]
    fn simple_and_complex_dirty_flags_preserve_aliases_and_unmodelled_content() {
        let word_namespace = crate::namespace::W_NS;
        let mut simple = parse_paragraph(&format!(
            r#"<q:fldSimple xmlns:q="{word_namespace}" xmlns:x="urn:producer" q:instr="REF destination" q:dirty="on" x:token="simple"><x:before/><q:r><q:t>old</q:t><x:inside/></q:r><x:after/></q:fldSimple>"#,
        ));
        let RunContent::Field(field) = &mut simple.runs[0].content[0] else {
            panic!("expected simple field")
        };
        assert_eq!(field.dirty, Some(true));
        field.cached_result = "new".to_owned();
        field.dirty = Some(false);
        let output = serialized_paragraph(&simple);
        assert!(output.contains(r#"<w:fldSimple xmlns:q="#,), "{output}");
        assert!(output.contains(r#"w:dirty="0""#), "{output}");
        assert!(output.contains(r#"x:token="simple""#), "{output}");
        assert!(output.contains("<x:before/>"), "{output}");
        assert!(output.contains("<x:inside/>"), "{output}");
        assert!(output.contains("<x:after/>"), "{output}");

        let mut complex = parse_paragraph(&format!(
            concat!(
                r#"<q:r xmlns:q="{}" xmlns:x="urn:producer"><q:fldChar q:fldCharType="begin" q:dirty="off"><x:begin/></q:fldChar></q:r>"#,
                r#"<q:r xmlns:q="{}" xmlns:x="urn:producer"><q:instrText>DATE</q:instrText><x:instruction/></q:r>"#,
                r#"<x:between xmlns:x="urn:producer"/>"#,
                r#"<q:r xmlns:q="{}"><q:fldChar q:fldCharType="separate" q:dirty="off"/></q:r>"#,
                r#"<q:r xmlns:q="{}" xmlns:x="urn:producer"><q:t>old</q:t><x:result/></q:r>"#,
                r#"<q:r xmlns:q="{}"><q:fldChar q:fldCharType="end" q:dirty="off"/></q:r>"#,
            ),
            word_namespace, word_namespace, word_namespace, word_namespace, word_namespace,
        ));
        let RunContent::Field(field) = &mut complex.runs[0].content[0] else {
            panic!("expected complex field")
        };
        assert_eq!(field.dirty, Some(false));
        field.cached_result = "new".to_owned();
        field.dirty = Some(true);
        let output = serialized_paragraph(&complex);
        assert!(
            output.contains(r#"<w:fldChar w:fldCharType="begin" w:dirty="1">"#),
            "{output}"
        );
        assert_eq!(output.matches(r#"w:dirty="1""#).count(), 1, "{output}");
        assert!(output.contains("<x:begin/>"), "{output}");
        assert!(output.contains("<x:instruction/>"), "{output}");
        assert!(output.contains("<x:between"), "{output}");
        assert!(output.contains("<x:result/>"), "{output}");
    }

    #[test]
    fn complex_field_projection_keeps_result_formatting_and_controls() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three</w:t><w:br w:type="page"/><w:t>four</w:t><w:br w:type="column"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));

        let field = parsed_field(&paragraph, 0);
        assert_eq!(field.cached_result, "one\ttwo\nthree\u{000c}four\u{000b}");
        let segments = field.cached_display_segments();
        assert_eq!(segments.len(), 1);
        assert_eq!(segments[0].1.unwrap().bold, Some(true));
    }

    #[test]
    fn complex_field_projection_keeps_each_result_runs_formatting() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>"#,
            r#"<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));

        let segments = parsed_field(&paragraph, 0).cached_display_segments();
        assert_eq!(segments.len(), 2);
        assert_eq!(segments[0].0, "bold");
        assert_eq!(segments[0].1.unwrap().bold, Some(true));
        assert_eq!(segments[1].0, "italic");
        assert_eq!(segments[1].1.unwrap().italic, Some(true));
    }

    #[test]
    fn edited_complex_cache_keeps_the_first_result_runs_formatting() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>old</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "fresh".to_owned();

        let segments = field.cached_display_segments();
        assert_eq!(segments.len(), 1);
        assert_eq!(segments[0].0, "fresh");
        assert_eq!(segments[0].1.unwrap().bold, Some(true));
        assert_eq!(segments[0].1.unwrap().italic, Some(true));
    }

    #[test]
    fn complex_field_mutation_clears_non_begin_dirty_and_adds_missing_result_text() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>"#,
            r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate" w:dirty="1"/></w:r>"#,
            r#"<w:r><w:producerResult/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end" w:dirty="1"/></w:r>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "fresh".to_owned();
        field.dirty = Some(false);

        let output = serialized_paragraph(&paragraph);
        assert_eq!(output.matches("w:dirty=").count(), 1, "{output}");
        assert!(output.contains(r#"w:dirty="0""#), "{output}");
        assert!(output.contains("<w:t>fresh</w:t>"), "{output}");
        assert!(output.contains("<w:producerResult/>"), "{output}");

        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        let field = parsed_field(&reparsed, 0);
        assert_eq!(field.cached_result, "fresh");
        assert_eq!(field.dirty, Some(false));
    }

    #[test]
    fn dirty_only_complex_rewrite_does_not_duplicate_result_controls() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>"#,
            r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>one</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three</w:t><w:br w:type="page"/><w:t>four</w:t><w:br w:type="column"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let expected = parsed_field(&paragraph, 0).cached_result.clone();
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.dirty = Some(false);

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert_eq!(parsed_field(&reparsed, 0).cached_result, expected);
    }

    #[test]
    fn simple_field_cache_mutation_replaces_old_result_controls() {
        let mut paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="DATE"><w:r><w:t>one</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three</w:t><w:br w:type="page"/><w:t>four</w:t><w:br w:type="column"/></w:r></w:fldSimple>"#,
        );
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "fresh".to_owned();

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert_eq!(parsed_field(&reparsed, 0).cached_result, "fresh");
        assert!(!output.contains("<w:tab"), "{output}");
        assert!(!output.contains("<w:br"), "{output}");
    }

    #[test]
    fn simple_control_only_cache_mutation_keeps_result_run_formatting() {
        let mut paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="DATE"><w:r><w:rPr><w:b/><w:i/></w:rPr><w:tab/></w:r></w:fldSimple>"#,
        );
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "fresh".to_owned();

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        let field = parsed_field(&reparsed, 0);
        assert_eq!(field.cached_result, "fresh");
        let segments = field.cached_display_segments();
        assert_eq!(segments.len(), 1);
        assert_eq!(segments[0].1.unwrap().bold, Some(true));
        assert_eq!(segments[0].1.unwrap().italic, Some(true));
        assert_eq!(output.matches("<w:r>").count(), 1, "{output}");
    }

    #[test]
    fn simple_cache_mutation_skips_empty_leading_result_runs_for_formatting() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:fldSimple w:instr="DATE">"#,
            r#"<w:r><w:rPr><w:b/></w:rPr></w:r>"#,
            r#"<w:r><w:rPr><w:i/></w:rPr><w:t>old</w:t></w:r>"#,
            r#"</w:fldSimple>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "fresh".to_owned();
        let in_memory = field.cached_display_segments();
        assert_eq!(in_memory[0].1.unwrap().bold, None);
        assert_eq!(in_memory[0].1.unwrap().italic, Some(true));

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        let reopened = parsed_field(&reparsed, 0).cached_display_segments();
        assert_eq!(reopened[0].0, "fresh");
        assert_eq!(reopened[0].1.unwrap().bold, None);
        assert_eq!(reopened[0].1.unwrap().italic, Some(true));
    }

    #[test]
    fn simple_cache_mutation_uses_an_empty_text_runs_formatting() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:fldSimple w:instr="DATE">"#,
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t/></w:r>"#,
            r#"<w:r><w:rPr><w:i/></w:rPr><w:t>old</w:t></w:r>"#,
            r#"</w:fldSimple>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "fresh".to_owned();
        let in_memory = field.cached_display_segments();
        assert_eq!(in_memory[0].1.unwrap().bold, Some(true));
        assert_eq!(in_memory[0].1.unwrap().italic, None);

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        let reopened = parsed_field(&reparsed, 0).cached_display_segments();
        assert_eq!(reopened[0].0, "fresh");
        assert_eq!(reopened[0].1.unwrap().bold, Some(true));
        assert_eq!(reopened[0].1.unwrap().italic, None);
    }

    #[test]
    fn expanded_simple_result_controls_contribute_to_the_cached_display() {
        let paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="DATE"><w:r><w:t>one</w:t><w:tab></w:tab><w:t>two</w:t><w:br></w:br><w:t>three</w:t><w:br w:type="page"></w:br><w:t>four</w:t><w:br w:type="column"></w:br></w:r></w:fldSimple>"#,
        );
        assert_eq!(
            parsed_field(&paragraph, 0).cached_result,
            "one\ttwo\nthree\u{000c}four\u{000b}"
        );
    }

    #[test]
    fn cache_mutation_replaces_a_nested_only_complex_result() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>IF 1 = 1 yes no</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>old nested result</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "new result".to_owned();

        let output = serialized_paragraph(&paragraph);
        assert!(!output.contains("old nested result"), "{output}");
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert_eq!(parsed_field(&reparsed, 0).cached_result, "new result");
    }

    #[test]
    fn canonical_instruction_preserves_empty_and_backslash_leading_operands() {
        let mut paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="INCLUDETEXT old"><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
        );
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.instruction.arguments = vec![
            FieldArgument::Text(String::new()),
            FieldArgument::Text(r#"\\server\share\file.docx"#.to_owned()),
        ];

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert_eq!(
            parsed_field(&reparsed, 0).instruction.arguments,
            vec![
                FieldArgument::Text(String::new()),
                FieldArgument::Text(r#"\\server\share\file.docx"#.to_owned()),
            ]
        );
    }

    #[test]
    fn complex_field_inside_explicit_hyperlink_is_projected() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:hyperlink w:anchor="destination">"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>cached</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"</w:hyperlink>"#,
        ));

        assert_eq!(paragraph.runs.len(), 1);
        assert_eq!(parsed_field(&paragraph, 0).instruction.name, "REF");
        assert_eq!(paragraph.hyperlinks[0].run_start, 0);
        assert_eq!(paragraph.hyperlinks[0].run_end, 1);
    }

    #[test]
    fn hyperlink_local_raw_children_stay_inside_a_projected_field_boundary() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:hyperlink w:anchor="destination">"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<x:producer xmlns:x="urn:producer"/>"#,
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>cached</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"</w:hyperlink>"#,
        ));

        assert_eq!(paragraph.runs.len(), 1);
        let output = serialized_paragraph(&paragraph);
        assert_eq!(output.matches("<x:producer").count(), 1, "{output}");
        assert!(
            output.find("fldCharType=\"begin\"").unwrap() < output.find("<x:producer").unwrap(),
            "{output}"
        );
        assert!(
            output.find("<x:producer").unwrap() < output.find("<w:instrText").unwrap(),
            "{output}"
        );
    }

    #[test]
    fn hyperlink_complex_source_retains_inter_run_trivia() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:hyperlink w:anchor="destination">"#,
            "\n  ",
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            "\n  <!-- hyperlink-before-instruction -->",
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            "<?producer hyperlink-before-separator?>",
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            "\n  <!-- hyperlink-before-result -->",
            r#"<w:r><w:t>cached</w:t></w:r>"#,
            "<?producer hyperlink-before-end?>",
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            "\n",
            r#"</w:hyperlink>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected hyperlink field")
        };
        let (captured, _) = field.source_replacement().unwrap().unwrap();
        for preserved in [
            "\n  <!-- hyperlink-before-instruction -->",
            "<?producer hyperlink-before-separator?>",
            "\n  <!-- hyperlink-before-result -->",
            "<?producer hyperlink-before-end?>",
        ] {
            assert!(
                String::from_utf8_lossy(captured).contains(preserved),
                "missing {preserved}: {}",
                String::from_utf8_lossy(captured)
            );
        }

        field.cached_result = "updated hyperlink".to_owned();
        field.dirty = Some(false);
        let output = serialized_paragraph(&paragraph);
        for preserved in [
            "<!-- hyperlink-before-instruction -->",
            "<?producer hyperlink-before-separator?>",
            "<!-- hyperlink-before-result -->",
            "<?producer hyperlink-before-end?>",
        ] {
            assert!(output.contains(preserved), "missing {preserved}: {output}");
        }
        assert!(output.contains("updated hyperlink"), "{output}");
        assert!(output.contains("<w:hyperlink"), "{output}");
    }

    #[test]
    fn cache_mutation_marks_significant_result_whitespace() {
        for source in [
            r#"<w:fldSimple w:instr="DATE"><w:r><w:t>old</w:t></w:r></w:fldSimple>"#,
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>old</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
        ] {
            let mut paragraph = parse_paragraph(source);
            let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
                panic!("expected field")
            };
            field.cached_result = " value ".to_owned();
            let output = serialized_paragraph(&paragraph);
            assert!(
                output.contains(r#"<w:t xml:space="preserve"> value </w:t>"#),
                "{output}"
            );
        }
    }

    #[test]
    fn nested_structured_edit_converts_a_simple_source_to_complex_form() {
        let mut paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="IF 1 = 1 yes no"><w:r><w:t>yes</w:t></w:r></w:fldSimple>"#,
        );
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.instruction.arguments = vec![
            FieldArgument::Nested(Box::new(Field::new("REF destination", "cached"))),
            FieldArgument::Text("=".to_owned()),
            FieldArgument::Text("cached".to_owned()),
            FieldArgument::Text("yes".to_owned()),
            FieldArgument::Text("no".to_owned()),
        ];

        let output = serialized_paragraph(&paragraph);
        assert!(!output.starts_with("<w:p><w:fldSimple"), "{output}");
        assert!(output.contains("REF destination"), "{output}");
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert!(matches!(
            parsed_field(&reparsed, 0).instruction.arguments.first(),
            Some(FieldArgument::Nested(nested)) if nested.instruction.name == "REF"
        ));
    }

    #[test]
    fn malformed_nested_result_field_invalidates_the_outer_field() {
        let source = concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>IF 1 = 1 yes no</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>nested</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r><w:t>outer</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        );
        let paragraph = parse_paragraph(source);

        assert!(!paragraph.runs.iter().any(|run| {
            run.content
                .iter()
                .any(|content| matches!(content, RunContent::Field(_)))
        }));
        assert!(serialized_paragraph(&paragraph).contains(source));
    }

    #[test]
    fn public_instruction_edits_serialize_consistently_for_both_forms() {
        let sources = [
            r#"<w:fldSimple w:instr="DATE"><w:r><w:t>cached</w:t></w:r></w:fldSimple>"#,
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText>DATE</w:instrText></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>cached</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
        ];

        for source in sources {
            let mut raw_edit = parse_paragraph(source);
            let RunContent::Field(field) = &mut raw_edit.runs[0].content[0] else {
                panic!("expected field")
            };
            field.instruction.raw = r"TIME \@ HH:mm".to_owned();
            let output = serialized_paragraph(&raw_edit);
            assert!(output.contains("TIME"), "raw edit absent: {output}");
            assert!(
                !output.contains(">DATE<") && !output.contains("&quot;DATE&quot;"),
                "{output}"
            );

            let mut structured_edit = parse_paragraph(source);
            let RunContent::Field(field) = &mut structured_edit.runs[0].content[0] else {
                panic!("expected field")
            };
            field.instruction.name = "TIME".to_owned();
            field.instruction.arguments.clear();
            field.instruction.switches = vec![FieldSwitch {
                name: "@".to_owned(),
                argument: Some(FieldArgument::Text("HH:mm".to_owned())),
            }];
            let output = serialized_paragraph(&structured_edit);
            assert!(output.contains("TIME"), "structured edit absent: {output}");
            assert!(
                !output.contains(">DATE<") && !output.contains("&quot;DATE&quot;"),
                "{output}"
            );
        }
    }

    #[test]
    fn malformed_complex_fields_remain_untyped_and_preserved() {
        for source in [
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r><w:instrText>REF destination</w:instrText></w:r>"#,
                r#"<w:r><w:t>literal</w:t></w:r>"#,
            ),
            concat!(
                r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r><w:t>literal</w:t></w:r>"#,
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
        ] {
            let paragraph = parse_paragraph(source);
            assert!(!paragraph.runs.iter().any(|run| {
                run.content
                    .iter()
                    .any(|content| matches!(content, RunContent::Field(_)))
            }));
            assert_eq!(paragraph.text(), "literal");

            let mut output = Vec::new();
            paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
            let output = String::from_utf8(output).unwrap();
            assert!(output.contains(source), "{output}");
        }
    }

    #[test]
    fn unchanged_complex_fields_keep_source_runs_and_unmodelled_neighbours() {
        let source = concat!(
            r#"<w:customBefore data="left"/>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> DATE </w:instrText><w:producer/></w:r>"#,
            r#"<w:proofErr w:type="spellStart"/>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:rPr><w:i/></w:rPr><w:t>17 August 2026</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:customAfter data="right"/>"#,
        );
        let mut paragraph = parse_paragraph(source);
        assert_eq!(paragraph.runs.len(), 1);
        assert_eq!(parsed_field(&paragraph, 0).instruction.name, "DATE");

        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            format!("<w:p>{source}</w:p>")
        );

        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected complex field")
        };
        field.cached_result = "18 August 2026".to_owned();
        field.dirty = Some(false);
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        let begin = output.find(r#"w:fldCharType="begin""#).unwrap();
        let instruction = output.find("<w:instrText").unwrap();
        let separate = output.find(r#"w:fldCharType="separate""#).unwrap();
        let result = output.find("<w:t>18 August 2026</w:t>").unwrap();
        let end = output.find(r#"w:fldCharType="end""#).unwrap();
        assert!(begin < instruction && instruction < separate && separate < result && result < end);

        let word_namespace = crate::namespace::W_NS;
        let mut aliased = parse_paragraph(&format!(
            r#"<q:fldSimple xmlns:q="{word_namespace}" q:instr="REF destination" q:dirty="true"><q:r><q:t>old</q:t><q:custom/></q:r></q:fldSimple>"#,
        ));
        let RunContent::Field(field) = &mut aliased.runs[0].content[0] else {
            panic!("expected aliased field")
        };
        field.cached_result = "new".to_owned();
        field.dirty = Some(false);

        let mut output = Vec::new();
        aliased.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<w:fldSimple"), "{output}");
        assert!(output.contains(r#"w:instr="REF destination""#), "{output}");
        assert!(output.contains(r#"w:dirty="0""#), "{output}");
        assert!(output.contains("<w:t>new</w:t>"), "{output}");
        assert!(!output.contains("<q:fldSimple"), "{output}");
    }

    #[test]
    fn ref_and_pageref_instructions_keep_targets_and_switches() {
        let ref_field = parse_paragraph(
            r#"<w:fldSimple w:instr=" REF destination \h \* MERGEFORMAT "><w:r><w:t>cached text</w:t></w:r></w:fldSimple>"#,
        );
        let page_ref = parse_paragraph(
            r#"<w:fldSimple w:instr=" PAGEREF destination \p "><w:r><w:t>7</w:t></w:r></w:fldSimple>"#,
        );

        let ref_field_model = parsed_field(&ref_field, 0);
        assert_eq!(ref_field_model.instruction.name, "REF");
        assert!(matches!(
            ref_field_model.instruction.arguments.first(),
            Some(FieldArgument::Text(bookmark)) if bookmark == "destination"
        ));
        assert_eq!(
            ref_field_model.instruction.raw,
            r"REF destination \h \* MERGEFORMAT"
        );
        assert_eq!(ref_field_model.cached_result, "cached text");
        let page_ref_model = parsed_field(&page_ref, 0);
        assert_eq!(page_ref_model.instruction.name, "PAGEREF");
        assert!(matches!(
            page_ref_model.instruction.arguments.first(),
            Some(FieldArgument::Text(bookmark)) if bookmark == "destination"
        ));
        assert_eq!(page_ref_model.instruction.raw, r"PAGEREF destination \p");
        assert_eq!(page_ref_model.cached_result, "7");

        let mut output = Vec::new();
        ref_field.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"w:instr=" REF destination \h \* MERGEFORMAT ""#));
        assert!(output.contains("<w:t>cached text</w:t>"));
    }

    #[test]
    fn empty_cross_reference_displays_remain_empty() {
        let paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr=" REF destination "><w:r><w:t></w:t></w:r></w:fldSimple><w:fldSimple w:instr=" PAGEREF destination "><w:r><w:t></w:t></w:r></w:fldSimple>"#,
        );

        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert_eq!(output.matches("<w:t></w:t>").count(), 2, "{output}");
        assert!(!output.contains("<w:t>1</w:t>"), "{output}");
    }

    #[test]
    fn field_parser_uses_expanded_names_and_accepts_word_aliases() {
        let word_namespace = crate::namespace::W_NS;
        let paragraph = parse_paragraph(&format!(
            r#"<x:fldSimple xmlns:x="urn:producer" x:instr="REF foreign"><x:r><x:t>foreign</x:t></x:r></x:fldSimple><q:fldSimple xmlns:q="{word_namespace}" q:instr="REF destination"><q:r><q:t>cached</q:t></q:r></q:fldSimple>"#,
        ));

        assert_eq!(paragraph.runs.len(), 1);
        let field = parsed_field(&paragraph, 0);
        assert_eq!(field.instruction.name, "REF");
        assert!(matches!(
            field.instruction.arguments.first(),
            Some(FieldArgument::Text(bookmark)) if bookmark == "destination"
        ));
        assert_eq!(field.cached_result, "cached");
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<x:fldSimple"), "{output}");
        assert!(output.contains("x:instr=\"REF foreign\""), "{output}");
        assert!(output.contains("<q:fldSimple"), "{output}");
    }

    #[test]
    fn bookmark_markers_keep_range_order_and_unmodelled_neighbours() {
        let p = parse_paragraph(
            r#"<w:customBefore/><w:bookmarkStart w:id="4" w:name="destination"/><w:r><w:t>inside</w:t></w:r><w:bookmarkEnd w:id="4"/><w:customAfter/>"#,
        );

        assert_eq!(p.bookmark_markers.len(), 2);
        assert!(p.bookmark_markers[0].is_start());
        assert_eq!(p.bookmark_markers[0].id(), Some(4));
        assert_eq!(p.bookmark_markers[0].name(), Some("destination"));
        assert_eq!(p.bookmark_markers[0].run_index(), 0);
        assert!(!p.bookmark_markers[1].is_start());
        assert_eq!(p.bookmark_markers[1].run_index(), 1);

        let mut output = Vec::new();
        p.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        let before = output.find("<w:customBefore/>").unwrap();
        let start = output.find("<w:bookmarkStart").unwrap();
        let run = output.find("<w:r>").unwrap();
        let end = output.find("<w:bookmarkEnd").unwrap();
        let after = output.find("<w:customAfter/>").unwrap();
        assert!(before < start && start < run && run < end && end < after);
    }

    #[test]
    fn parse_fld_simple_mixed_with_text() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>5</w:t></w:r></w:fldSimple>"#,
        );
        assert_eq!(p.runs.len(), 4);
        assert_eq!(p.text(), "Page  of ");
        assert_eq!(parsed_field(&p, 1).instruction.name, "PAGE");
        assert_eq!(parsed_field(&p, 3).instruction.name, "NUMPAGES");
    }

    #[test]
    fn round_trip_fld_simple() {
        let mut p = CT_P::new();
        p.add_run("Page ");
        p.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(" PAGE ", "1"))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.runs.len(), 2);
        assert_eq!(parsed_field(&parsed, 1).instruction.name, "PAGE");
    }

    #[test]
    fn round_trip_paragraph() {
        let mut p = CT_P::new();
        p.add_run("Hello ");
        let run = p.add_run("World");
        run.properties = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.text(), "Hello World");
        assert_eq!(parsed.runs.len(), 2);
        assert_eq!(parsed.runs[1].properties.as_ref().unwrap().bold, Some(true));
    }

    #[test]
    fn parse_footnote_reference() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Some text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>"#,
        );
        assert_eq!(p.runs.len(), 2);
        assert_eq!(p.runs[0].text(), "Some text");
        assert_eq!(p.runs[1].content.len(), 1);
        assert!(matches!(
            p.runs[1].content[0],
            RunContent::FootnoteRef { id: 1 }
        ));
    }

    #[test]
    fn parse_endnote_reference() {
        let p = parse_paragraph(r#"<w:r><w:endnoteReference w:id="3"/></w:r>"#);
        assert_eq!(p.runs.len(), 1);
        assert!(matches!(
            p.runs[0].content[0],
            RunContent::EndnoteRef { id: 3 }
        ));
    }

    #[test]
    fn round_trip_footnote_reference() {
        let mut p = CT_P::new();
        p.add_run("Text before");
        p.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::FootnoteRef { id: 2 }],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        p.add_run(" text after");

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.runs.len(), 3);
        assert!(matches!(
            parsed.runs[1].content[0],
            RunContent::FootnoteRef { id: 2 }
        ));
    }

    #[test]
    fn comment_anchors_are_typed_without_moving_neighbouring_xml() {
        let word_namespace = crate::namespace::W_NS;
        let p = parse_paragraph(&format!(
            r#"<w:r><w:t>before</w:t></w:r><ext:marker xmlns:ext="urn:producer"/><q:commentRangeStart xmlns:q="{word_namespace}" q:id="7"/><w:r><w:t>inside</w:t></w:r><q:commentRangeEnd xmlns:q="{word_namespace}" q:id="7"/><w:r><ext:runBefore xmlns:ext="urn:producer"/><q:commentReference xmlns:q="{word_namespace}" q:id="7"/><ext:runAfter xmlns:ext="urn:producer"/></w:r><ext:tail xmlns:ext="urn:producer"/>"#
        ));

        assert_eq!(p.comment_ranges.len(), 2);
        assert!(matches!(
            p.comment_ranges[0],
            CommentRangeMarker::Start {
                id: 7,
                run_index: 1,
                ..
            }
        ));
        assert!(matches!(
            p.comment_ranges[1],
            CommentRangeMarker::End {
                id: 7,
                run_index: 2,
                ..
            }
        ));
        assert!(matches!(
            p.runs[2].content[0],
            RunContent::CommentReference { id: 7, .. }
        ));

        let mut output = Vec::new();
        p.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(
            r#"<ext:marker xmlns:ext="urn:producer"/><w:commentRangeStart w:id="7"/><w:r><w:t>inside</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><ext:runBefore xmlns:ext="urn:producer"/><w:commentReference w:id="7"/><ext:runAfter xmlns:ext="urn:producer"/></w:r><ext:tail xmlns:ext="urn:producer"/>"#
        ));
    }

    #[test]
    fn edited_nested_field_order_matches_canonical_serialization() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve">MERGEFIELD \b </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>AUTHOR</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>stored author</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r><w:instrText xml:space="preserve"> </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>MERGEFIELD Name</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>stored name</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>stored outer</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        assert_eq!(
            parsed_field(&paragraph, 0)
                .nested_fields_in_source_order()
                .iter()
                .map(|field| field.instruction.name.as_str())
                .collect::<Vec<_>>(),
            ["AUTHOR", "MERGEFIELD"]
        );

        let mut raw_edited = paragraph.clone();
        let RunContent::Field(field) = &mut raw_edited.runs[0].content[0] else {
            unreachable!()
        };
        field.instruction.raw = "AUTHOR".to_owned();
        assert!(field.nested_fields_in_source_order().is_empty());
        assert_eq!(field.effective_instruction().name, "AUTHOR");
        let output = serialized_paragraph(&raw_edited);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert_eq!(parsed_field(&reparsed, 0).instruction.name, "AUTHOR");
        assert!(
            parsed_field(&reparsed, 0)
                .nested_fields_in_source_order()
                .is_empty()
        );

        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        field.instruction.arguments.insert(
            0,
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD First", "stored first"))),
        );
        let expected = ["MERGEFIELD First", "MERGEFIELD Name", "AUTHOR"];
        assert_eq!(
            field
                .nested_fields_in_source_order()
                .iter()
                .map(|field| field.instruction.raw.as_str())
                .collect::<Vec<_>>(),
            expected
        );

        let output = serialized_paragraph(&paragraph);
        let reparsed = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .and_then(|value| value.strip_suffix("</w:p>"))
                .unwrap(),
        );
        assert_eq!(
            parsed_field(&reparsed, 0)
                .nested_fields_in_source_order()
                .iter()
                .map(|field| field.instruction.raw.as_str())
                .collect::<Vec<_>>(),
            expected
        );
    }

    #[test]
    fn malformed_comment_anchor_id_is_rejected() {
        let full = format!(
            r#"<w:p xmlns:w="{}"><w:commentRangeStart w:id="not-a-number"/></w:p>"#,
            crate::namespace::W_NS
        );
        let mut reader = Reader::from_str(&full);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element))
                    if matches_local_name(element.name().as_ref(), b"p") =>
                {
                    break;
                }
                Ok(Event::Eof) => panic!("missing paragraph"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        }
        assert!(CT_P::from_xml(&mut reader).is_err());
    }

    #[test]
    fn undecodable_ordinary_and_deleted_text_are_rejected() {
        for child in [
            b"<w:t>before\xff</w:t>".as_slice(),
            b"<w:delText>before\xff</w:delText>".as_slice(),
        ] {
            let mut xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r>"#.to_vec();
            xml.extend_from_slice(child);
            xml.extend_from_slice(b"</w:r></w:p>");
            let mut reader = Reader::from_reader(xml.as_slice());
            let mut buffer = Vec::new();
            assert!(matches!(
                reader.read_event_into(&mut buffer),
                Ok(Event::Start(_))
            ));
            assert!(matches!(
                CT_P::from_xml(&mut reader),
                Err(OxmlError::InvalidValue(_))
            ));
        }

        let valid = parse_paragraph(
            "<w:r><w:t>one &amp; two</w:t><w:delText>three &lt; four</w:delText></w:r>",
        );
        assert_eq!(valid.runs[0].text(), "one & twothree < four");
    }

    #[test]
    fn collapsed_run_boundaries_rebase_equation_raw_slots() {
        let math_namespace = crate::namespace::M_NS;
        let mut paragraph = parse_paragraph(&format!(
            r#"<m:oMath xmlns:m="{math_namespace}"><m:r><m:t>first</m:t></m:r></m:oMath><w:r><w:commentReference w:id="7"/></w:r><m:oMath xmlns:m="{math_namespace}"><m:r><m:t>second</m:t></m:r></m:oMath><w:r><w:t>kept</w:t></w:r>"#
        ));
        paragraph.remove_comment_anchors(&[7]);
        assert_eq!(paragraph.equations.len(), 2);
        assert_eq!(paragraph.equations[0].0, 0);
        assert_eq!(paragraph.equations[0].1, 0);
        assert_eq!(paragraph.equations[1].0, 0);
        assert_eq!(paragraph.equations[1].1, 1);
        let OfficeMath::Inline(second) = &mut paragraph.equations[1].2 else {
            panic!("inline equation")
        };
        let crate::math::MathExpression::Run(run) = &mut second.expressions[0] else {
            panic!("math run")
        };
        run.text = "changed".to_owned();

        let output = serialized_paragraph(&paragraph);
        assert!(output.find("first").unwrap() < output.find("changed").unwrap());
        assert!(!output.contains("second"));
    }

    #[test]
    fn nested_math_text_extension_remains_raw_through_paragraph_reopen() {
        let math_namespace = crate::namespace::M_NS;
        let relationships_namespace = crate::namespace::R_NS;
        let compatibility_namespace = crate::namespace::MC_NS;
        let equation = format!(
            r#"<m:oMath xmlns:m="{math_namespace}" xmlns:r="{relationships_namespace}" xmlns:mc="{compatibility_namespace}"><m:r><m:t>x<x:nested xmlns:x="urn:test"/></m:t></m:r></m:oMath>"#
        );
        let source =
            format!(r#"<w:r><w:t>before</w:t></w:r>{equation}<w:r><w:t>after</w:t></w:r>"#);
        let paragraph = parse_paragraph(&source);
        assert_eq!(paragraph.text(), "beforeafter");
        assert_eq!(paragraph.equations.len(), 1);
        assert!(paragraph.equations[0].2.has_unsupported_content());

        let first = serialized_paragraph(&paragraph);
        assert!(first.contains(&equation), "{first}");
        let reopened = parse_paragraph(
            first
                .strip_prefix("<w:p>")
                .and_then(|xml| xml.strip_suffix("</w:p>"))
                .expect("serialized paragraph wrapper"),
        );
        assert_eq!(reopened.text(), "beforeafter");
        assert_eq!(reopened.equations.len(), 1);
        assert!(reopened.equations[0].2.has_unsupported_content());
        let second = serialized_paragraph(&reopened);
        assert!(second.contains(&equation), "{second}");
    }
}
