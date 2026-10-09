//! Text content elements: `CT_P` (paragraph), `CT_R` (run), `CT_Text`.

use std::borrow::Cow;
use std::cell::Cell;
use std::collections::HashSet;
use std::sync::atomic::{AtomicU64, Ordering};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::{NsReader, Reader, Writer, XmlVersion};

use crate::content_control::{CT_Sdt, SdtContent, SdtOwner};
use crate::drawing::CT_Drawing;
use crate::error::{OxmlError, Result};
use crate::math::OfficeMath;
use crate::namespace::{R_NS, matches_local_name};
use crate::numbering::{namespace_bindings, parse_scoped_ppr, word_prefixes_at};
use crate::properties::{CT_PPr, CT_RPr, is_word_attribute, is_word_element};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::{CT_Revision, RevisionKind};
use crate::ruby::CT_Ruby;
use crate::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};

static NEXT_FIELD_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

const ROOT_ATTRIBUTES_ELEMENT: &[u8] = b"rdocxRootAttributes";
pub(crate) const ROOT_ATTRIBUTES_POSITION: usize = usize::MAX;
const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
pub(crate) const ROOT_R_BINDING: u8 = 1;
pub(crate) const ROOT_MC_BINDING: u8 = 2;
pub(crate) const ROOT_WP_BINDING: u8 = 4;

thread_local! {
    static ROOT_BINDINGS: Cell<u8> = const { Cell::new(0) };
}

/// Keep part-root guarantees in scope only while that part is serialized.
/// Nested serializations restore the caller's guarantees, including on unwind.
pub(crate) struct RootBindingScope(u8);

pub(crate) fn root_binding_scope(bindings: u8) -> RootBindingScope {
    RootBindingScope(ROOT_BINDINGS.with(|current| current.replace(bindings)))
}

impl Drop for RootBindingScope {
    fn drop(&mut self) {
        ROOT_BINDINGS.with(|current| current.set(self.0));
    }
}

fn namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn root_attribute_namespace(name: &[u8], bindings: &[(String, String)]) -> Result<Option<String>> {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return Ok(None);
    };
    let prefix = &name[..separator];
    if prefix == b"xml" {
        return Ok(Some("http://www.w3.org/XML/1998/namespace".to_owned()));
    }
    bindings
        .iter()
        .find(|(candidate, _)| candidate.as_bytes() == prefix)
        .map(|(_, namespace)| Some(namespace.clone()))
        .ok_or_else(|| {
            OxmlError::InvalidValue(format!(
                "root attribute prefix `{}` is unbound",
                String::from_utf8_lossy(prefix)
            ))
        })
}

pub(crate) fn capture_root_attribute_record(
    start: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<Option<Vec<u8>>> {
    let mut bindings = namespace_bindings(prefixes);
    // A plain scope entry names a Word prefix. `word_prefixes_at` adds one
    // only beside its binding, but the default scope of the public `from_xml`
    // entrypoints, `CT_P::from_xml` among them, names `w` by convention and
    // binds nothing. A caller that parses a paragraph cut out of its part in
    // that scope, as the text-box replacement and template walkers do, reaches
    // this capture with it, so the Word prefix resolves here instead of
    // failing on the first `w:rsidR`. An explicit binding always wins.
    for prefix in prefixes {
        if !prefix.starts_with('\0') && !bindings.iter().any(|(bound, _)| bound == prefix) {
            bindings.push((prefix.clone(), crate::namespace::W_NS.to_owned()));
        }
    }
    let mut attributes = Vec::new();
    let mut expanded = HashSet::new();
    let mut used_prefixes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        // A namespace declaration is not retained for its own sake. The loop
        // below re-declares exactly the prefixes these attributes use, and the
        // alias machinery declares the rest on the elements that use them, so
        // recording a declaration here would emit it twice.
        if namespace_declaration(name) {
            continue;
        }
        let namespace = root_attribute_namespace(name, &bindings)?;
        let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
        if !expanded.insert((namespace, local.to_vec())) {
            return Err(OxmlError::InvalidValue(format!(
                "duplicate expanded root attribute `{}`",
                String::from_utf8_lossy(local)
            )));
        }
        if let Some(separator) = name.iter().position(|byte| *byte == b':')
            && let prefix = &name[..separator]
            && prefix != b"xml"
            && !used_prefixes.iter().any(|candidate| candidate == prefix)
        {
            used_prefixes.push(prefix.to_vec());
        }
        attributes.push((std::str::from_utf8(name)?.to_owned(), value));
    }
    if attributes.is_empty() {
        return Ok(None);
    }

    let mut record = BytesStart::new(std::str::from_utf8(ROOT_ATTRIBUTES_ELEMENT)?);
    for (name, value) in &attributes {
        record.push_attribute((name.as_str(), value.as_str()));
    }
    for prefix in used_prefixes {
        let declaration = format!("xmlns:{}", String::from_utf8_lossy(&prefix));
        let namespace = bindings
            .iter()
            .find(|(candidate, _)| candidate.as_bytes() == prefix)
            .map(|(_, namespace)| namespace.as_str())
            .ok_or_else(|| {
                OxmlError::InvalidValue(format!(
                    "root attribute prefix `{}` is unbound",
                    String::from_utf8_lossy(&prefix)
                ))
            })?;
        record.push_attribute((declaration.as_str(), namespace));
    }
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Empty(record))?;
    Ok(Some(writer.into_inner()))
}

#[doc(hidden)]
pub fn is_root_attribute_record(raw: &[u8]) -> bool {
    raw.starts_with(b"<rdocxRootAttributes")
}

pub(crate) fn push_root_attribute_record(
    target: &mut BytesStart<'_>,
    raw: &[u8],
    replaced: Option<(&str, &str)>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(raw);
    let mut buffer = Vec::new();
    let element = loop {
        match reader.read_resolved_event_into(&mut buffer)? {
            (_, Event::Empty(element)) if element.name().as_ref() == ROOT_ATTRIBUTES_ELEMENT => {
                break element.into_owned();
            }
            (_, Event::Eof) => {
                return Err(OxmlError::InvalidValue(
                    "invalid retained root-attribute record".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    };
    let existing = target
        .attributes()
        .filter_map(|attribute| attribute.ok())
        .map(|attribute| attribute.key.as_ref().to_vec())
        .collect::<HashSet<_>>();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        if existing.contains(name) {
            continue;
        }
        // `w14` is bound by the part root that owns the element, the same
        // assumption the authored `w14:paraId` write already makes. Rebinding
        // it here would make a reopened save differ from the save it came
        // from, purely by a declaration that changes nothing. `w` is in scope
        // already, since the element's own `w:` name needs it, and a writer
        // that shadows it declares it on the element first. Word writes
        // `w:rsid*` on nearly every paragraph and run, so repeating `w` here
        // repeated it hundreds of times in an edited part.
        if (name == b"xmlns:w14" && attribute.value.as_ref() == W14_NS.as_bytes())
            || (name == b"xmlns:w" && attribute.value.as_ref() == crate::namespace::W_NS.as_bytes())
        {
            continue;
        }
        let root_bindings = ROOT_BINDINGS.with(Cell::get);
        if (root_bindings & ROOT_R_BINDING != 0
            && name == b"xmlns:r"
            && attribute.value.as_ref() == R_NS.as_bytes())
            || (root_bindings & ROOT_MC_BINDING != 0
                && name == b"xmlns:mc"
                && attribute.value.as_ref() == MC_NS.as_bytes())
            || (root_bindings & ROOT_WP_BINDING != 0
                && name == b"xmlns:wp"
                && attribute.value.as_ref() == WP_NS.as_bytes())
        {
            continue;
        }
        if !namespace_declaration(name)
            && let Some((namespace, local)) = replaced
        {
            let (resolved, resolved_local) = reader.resolver().resolve_attribute(attribute.key);
            if resolved_local.as_ref() == local.as_bytes()
                && matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace.as_bytes())
            {
                continue;
            }
        }
        let name = std::str::from_utf8(name)?;
        let value =
            attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?;
        target.push_attribute((name, value.as_ref()));
    }
    Ok(())
}

/// Remove `w14:paraId` and `w14:textId` from the retained start-tag
/// attributes of a paragraph or table row, given its `extra_xml`.
///
/// A copy must not share them with its source, and Word assigns new ones to
/// an element that has none. Every other retained attribute stays, with the
/// declarations its prefix needs.
#[doc(hidden)]
pub fn drop_w14_paragraph_identities(extra_xml: &mut Vec<(usize, Vec<u8>)>) -> Result<()> {
    let mut index = 0;
    while index < extra_xml.len() {
        let (position, raw) = &extra_xml[index];
        if *position == ROOT_ATTRIBUTES_POSITION && is_root_attribute_record(raw) {
            match root_attribute_record_without_w14_identities(raw)? {
                Some(record) => extra_xml[index].1 = record,
                None => {
                    extra_xml.remove(index);
                    continue;
                }
            }
        }
        index += 1;
    }
    Ok(())
}

/// Return the record without its w14 identities, or `None` when nothing else
/// is left in it.
fn root_attribute_record_without_w14_identities(raw: &[u8]) -> Result<Option<Vec<u8>>> {
    let mut reader = NsReader::from_reader(raw);
    let mut buffer = Vec::new();
    let element = loop {
        match reader.read_resolved_event_into(&mut buffer)? {
            (_, Event::Empty(element)) if element.name().as_ref() == ROOT_ATTRIBUTES_ELEMENT => {
                break element.into_owned();
            }
            (_, Event::Eof) => {
                return Err(OxmlError::InvalidValue(
                    "invalid retained root-attribute record".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    };
    let mut dropped = false;
    let mut attributes = Vec::new();
    let mut declarations = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = std::str::from_utf8(attribute.key.as_ref())?.to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        if namespace_declaration(name.as_bytes()) {
            declarations.push((name, value));
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W14_NS.as_bytes())
            && matches!(local.as_ref(), b"paraId" | b"textId")
        {
            dropped = true;
            continue;
        }
        attributes.push((name, value));
    }
    if !dropped {
        return Ok(Some(raw.to_vec()));
    }
    if attributes.is_empty() {
        return Ok(None);
    }
    let mut record = BytesStart::new(std::str::from_utf8(ROOT_ATTRIBUTES_ELEMENT)?);
    for (name, value) in &attributes {
        record.push_attribute((name.as_str(), value.as_str()));
    }
    // The record declares exactly the prefixes its attributes use, so a
    // declaration only the dropped identities used goes with them.
    for (name, value) in &declarations {
        let prefix = name.strip_prefix("xmlns:").unwrap_or_default();
        if attributes.iter().any(|(attribute, _)| {
            attribute
                .split_once(':')
                .is_some_and(|(used, _)| used == prefix)
        }) {
            record.push_attribute((name.as_str(), value.as_str()));
        }
    }
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Empty(record))?;
    Ok(Some(writer.into_inner()))
}

/// Declare the canonical `w14` prefix on the root of a serialized part whose
/// content uses the prefix while the root does not bind it.
///
/// A retained root-attribute record writes `w14:paraId` and `w14:textId`
/// without their declaration, because Word and python-docx declare `w14` on
/// the part root. A root rdocx wrote does not, and a producer may declare the
/// prefix on the element that uses it, so the part declares it here once. A
/// root that binds `w14` itself is left as it is.
#[doc(hidden)]
pub fn declare_w14_on_part_root(xml: &mut Vec<u8>) -> Result<()> {
    let mut reader = Reader::from_reader(xml.as_slice());
    let mut buffer = Vec::new();
    let root_end = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(root) => {
                for attribute in root.attributes() {
                    if attribute?.key.as_ref() == b"xmlns:w14" {
                        return Ok(());
                    }
                }
                break reader.buffer_position() as usize;
            }
            Event::Empty(_) | Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    };
    // A qualified name starts after `<`, `</` or the whitespace before an
    // attribute. Text that happens to match only adds a declaration.
    let uses_w14 = xml[root_end..].windows(5).any(|window| {
        matches!(window[0], b'<' | b'/' | b' ' | b'\t' | b'\r' | b'\n') && &window[1..] == b"w14:"
    });
    if uses_w14 {
        let declaration = format!(r#" xmlns:w14="{W14_NS}""#);
        xml.splice(root_end - 1..root_end - 1, declaration.into_bytes());
    }
    Ok(())
}

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

    /// Keep the first `at` characters and return the rest.
    ///
    /// Either part preserves space when the original did or when the split
    /// leaves a space at one of its ends.
    fn split_off_at(&mut self, at: usize) -> CT_Text {
        let byte = self
            .text
            .char_indices()
            .nth(at)
            .map_or(self.text.len(), |(byte, _)| byte);
        let rest = self.text.split_off(byte);
        let inherited = self.preserve_space;
        let preserve = |text: &str| inherited || text.starts_with(' ') || text.ends_with(' ');
        let rest = CT_Text {
            preserve_space: preserve(&rest),
            text: rest,
        };
        self.preserve_space = preserve(&self.text);
        rest
    }
}

/// A parsed Word field with its stored result and update marker.
#[derive(Debug, Clone)]
pub struct Field {
    pub instruction: FieldInstruction,
    pub cached_result: String,
    pub dirty: Option<bool>,
    locked: Option<bool>,
    /// Typed legacy form metadata retained below the begin `w:fldChar`.
    pub legacy_form: Option<LegacyFormFieldData>,
    legacy_form_parse_error: bool,
    nested_order: Vec<NestedFieldPosition>,
    cached_fields: Vec<(std::ops::Range<usize>, Field)>,
    simple_cached_runs: Option<Vec<CT_R>>,
    typed_cached_runs: Option<Vec<CT_R>>,
    typed_cached_comment_ranges: Vec<CommentRangeMarker>,
    parsed_cached_comment_ranges: Vec<CommentRangeMarker>,
    parsed_cached_runs: Option<Vec<CT_R>>,
    source: FieldSource,
    /// The physical run this field shares with text outside it, if any.
    span: Option<Box<FieldRunSpan>>,
}

/// The model runs a physical field span was split into on read.
///
/// A run such as `Page {PAGE} of the report` holds text outside the field.
/// The reader turns it into sibling runs so every consumer sees that text,
/// and the writer puts the original bytes back while the runs are untouched.
#[derive(Debug, Clone)]
struct FieldRunSpan {
    /// The span's runs in order, `None` where a field of the span sits.
    runs: Vec<Option<CT_R>>,
    /// Where this field sits among `runs`.
    position: usize,
    /// Which field of the span this is, counted from the first.
    field_index: usize,
}

/// The bounded legacy form kind projected from `w:ffData`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyFormFieldKind {
    TextInput,
    CheckBox,
    DropDownList,
}

/// The current value stored by one legacy form field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LegacyFormFieldValue {
    Text(String),
    Checked(bool),
    SelectedIndex(usize),
}

/// Supported metadata from one structurally valid `w:ffData` subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyFormFieldData {
    pub name: Option<String>,
    pub enabled: bool,
    pub calculate_on_exit: bool,
    pub kind: LegacyFormFieldKind,
    pub value: LegacyFormFieldValue,
    pub choices: Vec<String>,
    pub max_length: Option<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
                form: FieldForm::Simple,
                cached_runs: None,
                original_cached_result: cached_result.to_owned(),
                cached_segments: Vec::new(),
            },
            instruction,
            cached_result: cached_result.to_owned(),
            dirty: None,
            locked: None,
            legacy_form: None,
            legacy_form_parse_error: false,
            nested_order: Vec::new(),
            cached_fields: Vec::new(),
            simple_cached_runs: None,
            typed_cached_runs: None,
            typed_cached_comment_ranges: Vec::new(),
            parsed_cached_comment_ranges: Vec::new(),
            parsed_cached_runs: None,
            span: None,
        }
    }

    /// Construct a checked field from a producer-style instruction.
    pub fn from_raw(instruction: &str, form: FieldForm, cached_runs: Vec<CT_R>) -> Result<Self> {
        validate_raw_field_instruction(instruction)?;
        let parsed = parse_field_instruction(instruction);
        let mut field = Self::from_instruction(parsed, form, cached_runs)?;
        field.instruction.raw = instruction.trim().to_owned();
        if let FieldSource::New {
            original_instruction,
            ..
        } = &mut field.source
        {
            *original_instruction = field.instruction.clone();
        }
        Ok(field)
    }

    /// Construct a checked field with typed operands and ordered cached runs.
    pub fn from_instruction(
        instruction: FieldInstruction,
        form: FieldForm,
        cached_runs: Vec<CT_R>,
    ) -> Result<Self> {
        let instruction = FieldInstruction::new(
            &instruction.name,
            instruction.arguments,
            instruction.switches,
        )?;
        if form == FieldForm::Simple && instruction_contains_nested(&instruction) {
            return Err(OxmlError::InvalidValue(
                "simple field instructions cannot contain nested fields".into(),
            ));
        }
        validate_cached_field_runs(&cached_runs)?;
        let cached_segments = cached_runs
            .iter()
            .enumerate()
            .map(|(run_index, run)| CachedDisplaySegment {
                run_index,
                text: cached_run_display(run),
                properties: run.properties.clone(),
            })
            .collect::<Vec<_>>();
        let cached_result = cached_segments
            .iter()
            .map(|segment| segment.text.as_str())
            .collect::<String>();
        let mut field = Self::new(&instruction.raw, &cached_result);
        field.instruction = instruction.clone();
        field.source = FieldSource::New {
            original_instruction: instruction,
            form,
            cached_runs: Some(cached_runs),
            original_cached_result: cached_result,
            cached_segments,
        };
        Ok(field)
    }

    /// Replace a field's result with checked typed runs while retaining its source controls.
    #[doc(hidden)]
    pub fn set_cached_runs(&mut self, runs: Vec<CT_R>) -> Result<()> {
        if self.locked == Some(true) {
            return Err(OxmlError::InvalidValue(
                "a locked field cannot replace its cached runs".into(),
            ));
        }
        validate_cached_field_runs(&runs)?;
        let mut candidate = self.clone();
        candidate.cached_result = runs.iter().map(cached_run_display).collect();
        candidate.typed_cached_runs = Some(runs);
        candidate.typed_cached_comment_ranges.clear();
        candidate.simple_cached_runs = None;
        candidate.cached_fields.clear();
        let mut writer = Writer::new(Vec::new());
        write_run_field(&mut writer, &candidate, None, None)?;
        let source = writer.into_inner();
        let mut paragraph = FIELD_CACHE_VALIDATION_SCOPE.to_vec();
        paragraph.extend(source);
        paragraph.extend_from_slice(b"</w:p>");
        CT_P::from_xml_fragment(&paragraph)?;
        *self = candidate;
        Ok(())
    }

    /// Replace a checked result and its sibling annotation range markers atomically.
    #[doc(hidden)]
    pub fn set_cached_runs_with_comment_ranges(
        &mut self,
        runs: Vec<CT_R>,
        ranges: Vec<CommentRangeMarker>,
    ) -> Result<()> {
        validate_cached_comment_ranges(&runs, &ranges)?;
        if ranges.is_empty() {
            return self.set_cached_runs(runs);
        }
        let mut candidate = self.clone();
        candidate.set_cached_runs(runs)?;
        candidate.typed_cached_comment_ranges = ranges;
        let mut writer = Writer::new(Vec::new());
        write_run_field(&mut writer, &candidate, None, None)?;
        let mut paragraph = FIELD_CACHE_VALIDATION_SCOPE.to_vec();
        paragraph.extend(writer.into_inner());
        paragraph.extend_from_slice(b"</w:p>");
        let reparsed = CT_P::from_xml_fragment(&paragraph)?;
        let owner = reparsed
            .runs()
            .into_iter()
            .flat_map(|run| &run.content)
            .find_map(|content| match content {
                RunContent::Field(field) => Some(field),
                _ => None,
            })
            .ok_or_else(|| {
                OxmlError::InvalidValue("annotation result did not reopen as a typed field".into())
            })?;
        if owner.effective_instruction() != candidate.effective_instruction()
            || owner.cached_result != candidate.cached_result
            || owner.cached_result_comment_ranges().len()
                != candidate.typed_cached_comment_ranges.len()
        {
            return Err(OxmlError::InvalidValue(
                "annotation result reopened with a different field owner or range graph".into(),
            ));
        }
        if let Some(runs) = owner.cached_result_runs() {
            validate_cached_comment_ranges(runs, owner.cached_result_comment_ranges())?;
        } else {
            return Err(OxmlError::InvalidValue(
                "annotation result has no typed cache after reopen".into(),
            ));
        }
        *self = candidate;
        Ok(())
    }

    /// Return sibling annotation boundaries owned by this field's stored result.
    #[doc(hidden)]
    pub fn cached_result_comment_ranges(&self) -> &[CommentRangeMarker] {
        if self.typed_cache().is_some() {
            &self.typed_cached_comment_ranges
        } else {
            &self.parsed_cached_comment_ranges
        }
    }

    /// Return the ordered typed result runs without replaying cached fields.
    #[doc(hidden)]
    pub fn cached_result_runs(&self) -> Option<&[CT_R]> {
        self.typed_cache()
            .or(self.simple_cached_runs.as_deref())
            .or(self.parsed_cached_runs.as_deref())
            .or(match &self.source {
                FieldSource::New { cached_runs, .. } => cached_runs.as_deref(),
                _ => None,
            })
            .filter(|runs| {
                runs.iter().map(cached_run_display).collect::<String>() == self.cached_result
            })
    }

    fn typed_cache(&self) -> Option<&[CT_R]> {
        self.typed_cached_runs.as_deref().filter(|runs| {
            runs.iter().map(cached_run_display).collect::<String>() == self.cached_result
        })
    }

    /// Return the representation this field writes.
    pub fn form(&self) -> FieldForm {
        if instruction_contains_nested(&self.effective_instruction()) {
            return FieldForm::Complex;
        }
        match &self.source {
            FieldSource::New { form, .. } | FieldSource::Parsed { form, .. } => *form,
        }
    }

    /// The field-local lock state, including an absent toggle.
    pub fn locked(&self) -> Option<bool> {
        self.locked
    }

    /// Set or remove the field-local lock without changing its instruction or cache.
    pub fn set_locked(&mut self, value: Option<bool>) {
        self.locked = value;
    }

    /// Validate the current mutable field before attaching it to a document.
    #[doc(hidden)]
    pub fn validate_for_attachment(&self) -> Result<()> {
        let instruction = self.effective_instruction();
        validate_raw_field_instruction(&instruction.raw)?;
        FieldInstruction::new(
            &instruction.name,
            instruction.arguments,
            instruction.switches,
        )?;
        if let Some(runs) = self.typed_cache() {
            validate_cached_field_runs(runs)?;
        }
        if let FieldSource::New {
            cached_runs: Some(runs),
            ..
        } = &self.source
        {
            validate_cached_field_runs(runs)?;
        }
        let original_cached_result = match &self.source {
            FieldSource::New {
                original_cached_result,
                cached_runs: Some(_),
                ..
            }
            | FieldSource::Parsed {
                original_cached_result,
                ..
            } => Some(original_cached_result),
            _ => None,
        };
        if original_cached_result != Some(&self.cached_result) {
            oxml_core::xml::reject_non_xml_characters("field result", &self.cached_result)?;
        }
        Ok(())
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
        let locked = field_source_lock(&raw_xml, form, &word_prefixes)
            .ok()
            .flatten();
        let source_id = NEXT_FIELD_SOURCE_ID.fetch_add(1, Ordering::Relaxed);
        let original_instruction = instruction.clone();
        let original_cached_result = cached_result.clone();
        let (legacy_form, legacy_form_parse_error) = match parse_legacy_form_data(
            &raw_xml,
            &word_prefixes,
            &instruction.name,
            &cached_result,
        ) {
            Ok(legacy_form) => (legacy_form, false),
            Err(_) => (None, true),
        };
        Self {
            instruction,
            cached_result,
            dirty,
            locked,
            legacy_form: legacy_form.clone(),
            legacy_form_parse_error,
            nested_order: Vec::new(),
            cached_fields: Vec::new(),
            simple_cached_runs: None,
            typed_cached_runs: None,
            typed_cached_comment_ranges: Vec::new(),
            parsed_cached_comment_ranges: Vec::new(),
            parsed_cached_runs: None,
            span: None,
            source: FieldSource::Parsed {
                source_id,
                owner_id: source_id,
                form,
                raw_xml,
                original_instruction,
                original_cached_result,
                cached_segments,
                original_dirty: dirty,
                original_locked: locked,
                original_legacy_form: legacy_form,
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

    /// Direct nested fields in the ordered cached result, without flattening it.
    #[doc(hidden)]
    pub fn cached_fields_in_source_order(&self) -> Vec<&Field> {
        if let Some(runs) = self.typed_cache().or(self.simple_cached_runs.as_deref()) {
            return runs
                .iter()
                .flat_map(|run| &run.content)
                .filter_map(|content| match content {
                    RunContent::Field(field) => Some(field),
                    _ => None,
                })
                .collect();
        }
        match &self.source {
            FieldSource::New {
                cached_runs: Some(runs),
                ..
            } => runs
                .iter()
                .flat_map(|run| &run.content)
                .filter_map(|content| match content {
                    RunContent::Field(field) => Some(field),
                    _ => None,
                })
                .collect(),
            _ => self.cached_fields.iter().map(|(_, field)| field).collect(),
        }
    }

    /// Direct nested operands followed by nested cached-result fields.
    #[doc(hidden)]
    pub fn all_nested_fields_in_source_order(&self) -> Vec<&Field> {
        let mut fields = self.nested_fields_in_source_order();
        fields.extend(self.cached_fields_in_source_order());
        fields
    }

    /// Mutate a direct cached-result field by its physical encounter index.
    /// Preserve a parsed child's identity. Checked serialization rejects a new
    /// replacement without the original simple-cache source span.
    #[doc(hidden)]
    pub fn cached_field_mut(&mut self, index: usize) -> Option<&mut Field> {
        if let Some(runs) = self
            .typed_cached_runs
            .as_mut()
            .or(self.simple_cached_runs.as_mut())
        {
            return runs
                .iter_mut()
                .flat_map(|run| &mut run.content)
                .filter_map(|content| match content {
                    RunContent::Field(field) => Some(field),
                    _ => None,
                })
                .nth(index);
        }
        match &mut self.source {
            FieldSource::New {
                cached_runs: Some(runs),
                ..
            } => runs
                .iter_mut()
                .flat_map(|run| &mut run.content)
                .filter_map(|content| match content {
                    RunContent::Field(field) => Some(field),
                    _ => None,
                })
                .nth(index),
            _ => self.cached_fields.get_mut(index).map(|(_, field)| field),
        }
    }

    /// Refresh the display projection after nested cached fields were edited.
    #[doc(hidden)]
    pub fn refresh_cached_field_projection(&mut self) {
        if let Some(value) = self.nested_cached_projection() {
            self.cached_result = value;
        }
    }

    fn nested_cached_projection(&self) -> Option<String> {
        if let Some(runs) = self.typed_cache().or(self.simple_cached_runs.as_deref()) {
            return Some(runs.iter().map(cached_run_display).collect());
        }
        match &self.source {
            FieldSource::New {
                cached_runs: Some(runs),
                ..
            } => Some(runs.iter().map(cached_run_display).collect()),
            FieldSource::Parsed {
                original_cached_result,
                ..
            } if !self.cached_fields.is_empty() => {
                let mut value = original_cached_result.clone();
                for (range, field) in self.cached_fields.iter().rev() {
                    if !value.is_char_boundary(range.start)
                        || !value.is_char_boundary(range.end)
                        || range.end > value.len()
                    {
                        return None;
                    }
                    value.replace_range(range.clone(), &field.cached_result);
                }
                Some(value)
            }
            _ => None,
        }
    }

    /// Reject a malformed legacy-form owner retained by this field.
    #[doc(hidden)]
    pub fn validate_legacy_form_owner(&self) -> Result<()> {
        if self.legacy_form_parse_error {
            return Err(OxmlError::InvalidValue(
                "malformed legacy form owner".to_owned(),
            ));
        }
        Ok(())
    }

    /// Replace the selected legacy form while retaining original nested-field order.
    #[doc(hidden)]
    pub fn set_nth_legacy_form_value_in_source_order(
        &mut self,
        remaining: &mut usize,
        value: &LegacyFormFieldValue,
    ) -> Result<bool> {
        self.validate_legacy_form_owner()?;
        if self.legacy_form.is_some() {
            if *remaining == 0 {
                self.set_legacy_form_value(value.clone())?;
                return Ok(true);
            }
            *remaining -= 1;
        }

        let original = self.original_instruction().clone();
        let structured_unchanged = instruction_structure_eq(&self.instruction, &original);
        if structured_unchanged && self.instruction.raw != original.raw {
            return Ok(false);
        }
        let ordered = if structured_unchanged {
            self.nested_order.clone()
        } else {
            Vec::new()
        };
        for position in &ordered {
            let nested = match *position {
                NestedFieldPosition::Argument(index) => self.instruction.arguments.get_mut(index),
                NestedFieldPosition::Switch(index) => self
                    .instruction
                    .switches
                    .get_mut(index)
                    .and_then(|field_switch| field_switch.argument.as_mut()),
            };
            if let Some(FieldArgument::Nested(field)) = nested
                && field.set_nth_legacy_form_value_in_source_order(remaining, value)?
            {
                return Ok(true);
            }
        }
        for (index, argument) in self.instruction.arguments.iter_mut().enumerate() {
            if ordered.contains(&NestedFieldPosition::Argument(index)) {
                continue;
            }
            if let FieldArgument::Nested(field) = argument
                && field.set_nth_legacy_form_value_in_source_order(remaining, value)?
            {
                return Ok(true);
            }
        }
        for (index, field_switch) in self.instruction.switches.iter_mut().enumerate() {
            if ordered.contains(&NestedFieldPosition::Switch(index)) {
                continue;
            }
            if let Some(FieldArgument::Nested(field)) = &mut field_switch.argument
                && field.set_nth_legacy_form_value_in_source_order(remaining, value)?
            {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Return the instruction selected by the public raw-versus-structured edit rules.
    #[doc(hidden)]
    pub fn effective_instruction(&self) -> FieldInstruction {
        instruction_for_write(self)
    }

    /// Return the effective instruction spelling without cloning nested field trees.
    #[doc(hidden)]
    pub fn effective_instruction_text(&self) -> String {
        let original = self.original_instruction();
        let structured_changed = !instruction_structure_eq(&self.instruction, original);
        if self.instruction.raw != original.raw && !structured_changed {
            self.instruction.raw.trim().to_owned()
        } else if structured_changed {
            canonical_instruction_text(&self.instruction)
        } else {
            self.instruction.raw.clone()
        }
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
                ..
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
        if self.typed_cached_runs.is_some() {
            return false;
        }
        match &self.source {
            FieldSource::New { .. } => false,
            FieldSource::Parsed {
                original_instruction,
                original_cached_result,
                original_dirty,
                original_locked,
                original_legacy_form,
                ..
            } => {
                self.instruction == *original_instruction
                    && instruction_source_identity_eq(&self.instruction, original_instruction)
                    && self.cached_result == *original_cached_result
                    && self.dirty == *original_dirty
                    && self.locked == *original_locked
                    && self.legacy_form == *original_legacy_form
                    && self
                        .cached_fields_in_source_order()
                        .iter()
                        .all(|field| field.is_unchanged())
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

    /// Whether this field writes a complex `w:fldChar` sequence.
    #[doc(hidden)]
    pub fn is_complex(&self) -> bool {
        self.form() == FieldForm::Complex
    }

    /// Return the text this field contributes to [`CT_R::text`].
    #[doc(hidden)]
    pub fn projected_text(&self) -> Option<&str> {
        (self.is_parsed_complex()
            || matches!(
                self.source,
                FieldSource::New {
                    form: FieldForm::Complex,
                    ..
                }
            ))
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
        self.cached_display_field_segments()
            .into_iter()
            .map(|(text, properties, _)| (text, properties))
            .collect()
    }

    /// Return the stored display split by its original result-run formatting.
    #[doc(hidden)]
    pub fn cached_display_field_segments(&self) -> Vec<(&str, Option<&CT_RPr>, &Field)> {
        let (original_cached_result, cached_segments) = match &self.source {
            FieldSource::Parsed {
                original_cached_result,
                cached_segments,
                ..
            }
            | FieldSource::New {
                original_cached_result,
                cached_segments,
                ..
            } => (original_cached_result, cached_segments),
        };
        if self.nested_cached_projection().as_deref() == Some(self.cached_result.as_str()) {
            let runs = self
                .typed_cache()
                .or(self.simple_cached_runs.as_deref())
                .or(match &self.source {
                    FieldSource::New {
                        cached_runs: Some(runs),
                        ..
                    } => Some(runs.as_slice()),
                    _ => None,
                });
            if let Some(runs) = runs {
                return runs
                    .iter()
                    .flat_map(|run| {
                        run.content.iter().flat_map(|content| match content {
                            RunContent::Field(child) => child.cached_display_field_segments(),
                            RunContent::Break(BreakType::Page) => {
                                vec![("\u{000c}", run.properties.as_ref(), self)]
                            }
                            RunContent::Break(BreakType::Column) => {
                                vec![("\u{000b}", run.properties.as_ref(), self)]
                            }
                            _ => vec![(CT_R::content_text(content), run.properties.as_ref(), self)],
                        })
                    })
                    .collect();
            }
            if !self.cached_fields.is_empty() {
                let mut segments = Vec::new();
                let mut start = 0;
                for (range, child) in &self.cached_fields {
                    segments.extend(cached_display_range(
                        cached_segments,
                        start..range.start,
                        self,
                    ));
                    segments.extend(child.cached_display_field_segments());
                    start = range.end;
                }
                segments.extend(cached_display_range(
                    cached_segments,
                    start..original_cached_result.len(),
                    self,
                ));
                return segments;
            }
        }
        if !cached_segments.is_empty() {
            if self.cached_result == *original_cached_result {
                return cached_segments
                    .iter()
                    .map(|segment| (segment.text.as_str(), segment.properties.as_ref(), self))
                    .collect();
            }
            return vec![(
                self.cached_result.as_str(),
                cached_segments
                    .first()
                    .and_then(|segment| segment.properties.as_ref()),
                self,
            )];
        }
        vec![(self.cached_result.as_str(), None, self)]
    }

    /// Whether this cached display owner is protected by any enclosing field.
    /// An owner outside this field's cached tree is conservatively locked.
    #[doc(hidden)]
    pub fn cached_display_owner_is_locked(&self, owner: &Field) -> bool {
        fn find(field: &Field, owner: &Field, inherited: bool) -> Option<bool> {
            let locked = inherited || field.locked() == Some(true);
            if std::ptr::eq(field, owner) {
                return Some(locked);
            }
            field
                .cached_fields_in_source_order()
                .into_iter()
                .find_map(|child| find(child, owner, locked))
        }
        find(self, owner, false).unwrap_or(true)
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

    /// Return what this field writes on its own when it was read from a run
    /// that also held content outside it.
    #[doc(hidden)]
    pub fn detached_source(&self) -> Result<Option<Vec<u8>>> {
        if self.span.is_none() {
            return Ok(None);
        }
        let mut writer = Writer::new(Vec::new());
        write_run_field(&mut writer, self, None, None)?;
        Ok(Some(writer.into_inner()))
    }

    /// Return the retained physical XML owner identity for a parsed field.
    #[doc(hidden)]
    pub fn source_owner_id(&self) -> Option<u64> {
        match self.source {
            FieldSource::Parsed { owner_id, .. } => Some(owner_id),
            FieldSource::New { .. } => None,
        }
    }

    /// Change the typed value of a legacy form field and its cached display.
    #[doc(hidden)]
    pub fn set_legacy_form_value(&mut self, value: LegacyFormFieldValue) -> Result<()> {
        let form = self
            .legacy_form
            .as_mut()
            .ok_or_else(|| OxmlError::MissingElement("w:ffData".to_owned()))?;
        let display = match (&form.kind, &value) {
            (LegacyFormFieldKind::TextInput, LegacyFormFieldValue::Text(value)) => {
                if form
                    .max_length
                    .is_some_and(|maximum| value.chars().count() > maximum)
                {
                    return Err(OxmlError::InvalidValue(
                        "legacy text form value exceeds w:maxLength".to_owned(),
                    ));
                }
                value.clone()
            }
            (LegacyFormFieldKind::CheckBox, LegacyFormFieldValue::Checked(checked)) => {
                if *checked { "☒" } else { "☐" }.to_owned()
            }
            (LegacyFormFieldKind::DropDownList, LegacyFormFieldValue::SelectedIndex(index)) => {
                let choice = form.choices.get(*index).ok_or_else(|| {
                    OxmlError::InvalidValue("legacy drop-down selection is out of range".to_owned())
                })?;
                choice.clone()
            }
            _ => {
                return Err(OxmlError::InvalidValue(
                    "legacy form value kind does not match the field kind".to_owned(),
                ));
            }
        };
        form.value = value;
        self.cached_result = display;
        Ok(())
    }
}

impl PartialEq for Field {
    fn eq(&self, other: &Self) -> bool {
        self.instruction == other.instruction
            && self.cached_result == other.cached_result
            && self.dirty == other.dirty
            && self.locked == other.locked
            && self.legacy_form == other.legacy_form
            && self.legacy_form_parse_error == other.legacy_form_parse_error
            && self.cached_result_runs() == other.cached_result_runs()
            && self.cached_result_comment_ranges() == other.cached_result_comment_ranges()
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

impl FieldInstruction {
    /// Whether the preserved instruction closes every unescaped quoted operand.
    #[doc(hidden)]
    pub fn quotes_are_balanced(&self) -> bool {
        let mut characters = self.raw.chars().peekable();
        let mut quoted = false;
        while let Some(character) = characters.next() {
            if quoted
                && character == '\\'
                && characters
                    .peek()
                    .is_some_and(|next| matches!(next, '"' | '\\'))
            {
                characters.next();
            } else if character == '"' {
                quoted = !quoted;
            }
        }
        !quoted
    }

    /// Build one canonical instruction from operands that cannot inject tokens.
    pub fn new(
        name: &str,
        arguments: Vec<FieldArgument>,
        switches: Vec<FieldSwitch>,
    ) -> Result<Self> {
        if !valid_field_token(name) || name.starts_with('\\') {
            return Err(OxmlError::InvalidValue("invalid field name".into()));
        }
        for argument in arguments.iter().chain(
            switches
                .iter()
                .filter_map(|switch| switch.argument.as_ref()),
        ) {
            match argument {
                FieldArgument::Text(value) => {
                    oxml_core::xml::reject_non_xml_characters("field argument", value)?
                }
                FieldArgument::Nested(field) => field.validate_for_attachment()?,
            }
        }
        let name = name.to_uppercase();
        let mut switches = switches;
        for switch in &mut switches {
            if !(valid_field_token(&switch.name) || switch.name == "!")
                || switch.name.contains('\\')
            {
                return Err(OxmlError::InvalidValue("invalid field switch name".into()));
            }
            switch.name = switch.name.to_ascii_lowercase();
            if switch.argument.is_some() && switch_is_known_flag(&name, &switch.name) {
                return Err(OxmlError::InvalidValue(format!(
                    "field switch {} does not take an argument",
                    switch.name
                )));
            }
        }
        let mut instruction = Self {
            raw: String::new(),
            name,
            arguments,
            switches,
        };
        instruction.raw = canonical_instruction_text(&instruction);
        Ok(instruction)
    }
}

fn valid_field_token(value: &str) -> bool {
    !value.is_empty()
        && value
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || matches!(c, '_' | '=' | '*' | '#' | '@'))
}

fn validate_raw_field_instruction(value: &str) -> Result<()> {
    oxml_core::xml::reject_non_xml_characters("field instruction", value)?;
    let mut quoted = false;
    let mut characters = value.chars().peekable();
    while let Some(character) = characters.next() {
        if quoted
            && character == '\\'
            && characters
                .peek()
                .is_some_and(|next| matches!(next, '"' | '\\'))
        {
            characters.next();
        } else if character == '"' {
            quoted = !quoted;
        }
    }
    if quoted {
        return Err(OxmlError::InvalidValue(
            "unbalanced field instruction quotes".into(),
        ));
    }
    let tokens = lex_field_text(value);
    let Some(InstructionToken {
        argument: FieldArgument::Text(name),
        quoted: false,
        ..
    }) = tokens.first()
    else {
        return Err(OxmlError::InvalidValue(
            "field instruction must start with a field name".into(),
        ));
    };
    if !valid_field_token(name) {
        return Err(OxmlError::InvalidValue("invalid field name".into()));
    }
    for token in &tokens[1..] {
        if let FieldArgument::Text(value) = &token.argument
            && !token.quoted
            && value.starts_with('\\')
            && !(valid_field_token(&value[1..]) || value == "\\!")
        {
            return Err(OxmlError::InvalidValue("invalid field switch token".into()));
        }
    }
    Ok(())
}

fn cached_display_range<'a>(
    segments: &'a [CachedDisplaySegment],
    range: std::ops::Range<usize>,
    owner: &'a Field,
) -> Vec<(&'a str, Option<&'a CT_RPr>, &'a Field)> {
    let mut offset = 0;
    let mut result = Vec::new();
    for segment in segments {
        let end = offset + segment.text.len();
        let start = range.start.max(offset);
        let stop = range.end.min(end);
        if start < stop
            && let Some(text) = segment.text.get(start - offset..stop - offset)
        {
            result.push((text, segment.properties.as_ref(), owner));
        }
        offset = end;
    }
    result
}

fn cached_run_display(run: &CT_R) -> String {
    run.content
        .iter()
        .map(|content| match content {
            RunContent::Field(field) => field.cached_result.clone(),
            RunContent::Break(BreakType::Page) => "\u{000c}".into(),
            RunContent::Break(BreakType::Column) => "\u{000b}".into(),
            _ => CT_R::content_text(content).to_owned(),
        })
        .collect()
}

const FIELD_CACHE_VALIDATION_SCOPE: &[u8] = concat!(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" ",
    "xmlns:v=\"urn:schemas-microsoft-com:vml\" ",
    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" ",
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ",
    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" ",
    "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
)
.as_bytes();

fn validate_cached_field_runs(runs: &[CT_R]) -> Result<()> {
    for run in runs {
        for content in &run.content {
            match content {
                RunContent::Text(text) | RunContent::DeletedText(text) => {
                    oxml_core::xml::reject_non_xml_characters("field result", &text.text)?
                }
                RunContent::Field(field) => field.validate_for_attachment()?,
                _ => {}
            }
        }
        for (index, raw) in run.extra_xml.iter().enumerate() {
            if run
                .extra_xml_positions
                .get(index)
                .is_some_and(|position| CT_R::raw_child_is_root_attributes(*position))
            {
                continue;
            }
            validate_raw_cached_field_content(run, raw)?;
        }
        let mut writer = Writer::new(Vec::new());
        writer
            .get_mut()
            .extend_from_slice(FIELD_CACHE_VALIDATION_SCOPE);
        run.to_xml(&mut writer)?;
        writer.get_mut().extend_from_slice(b"</w:p>");
        oxml_core::xml::validate_strict_xml_1_0(&writer.into_inner())
            .map_err(|error| OxmlError::InvalidValue(format!("invalid cached XML: {error:?}")))?;
    }
    Ok(())
}

fn validate_raw_cached_field_content(run: &CT_R, raw: &[u8]) -> Result<()> {
    let mut writer = Writer::new(FIELD_CACHE_VALIDATION_SCOPE.to_vec());
    let mut parent = BytesStart::new("w:r");
    for (position, attributes) in run.extra_xml_positions.iter().zip(&run.extra_xml) {
        if CT_R::raw_child_is_root_attributes(*position) {
            push_root_attribute_record(&mut parent, attributes, None)?;
        }
    }
    writer.write_event(Event::Start(parent))?;
    writer.get_mut().extend_from_slice(raw);
    writer.get_mut().extend_from_slice(b"</w:r></w:p>");
    let scoped = writer.into_inner();
    oxml_core::xml::validate_strict_xml_1_0(&scoped)
        .map_err(|error| OxmlError::InvalidValue(format!("invalid cached XML: {error:?}")))?;
    let mut reader = NsReader::from_reader(scoped.as_slice());
    loop {
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if namespace
                    == ResolveResult::Bound(Namespace(crate::namespace::W_NS.as_bytes()))
                    && matches!(
                        element.local_name().as_ref(),
                        b"fldChar"
                            | b"fldSimple"
                            | b"instrText"
                            | b"delInstrText"
                            | b"commentRangeStart"
                            | b"commentRangeEnd"
                    ) =>
            {
                return Err(OxmlError::InvalidValue(
                    "raw Word field delimiters are not cached display content".into(),
                ));
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
    }
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
        form: FieldForm,
        cached_runs: Option<Vec<CT_R>>,
        original_cached_result: String,
        cached_segments: Vec<CachedDisplaySegment>,
    },
    Parsed {
        source_id: u64,
        owner_id: u64,
        form: FieldForm,
        raw_xml: Vec<u8>,
        original_instruction: FieldInstruction,
        original_cached_result: String,
        cached_segments: Vec<CachedDisplaySegment>,
        original_dirty: Option<bool>,
        original_locked: Option<bool>,
        original_legacy_form: Option<LegacyFormFieldData>,
        word_prefixes: Vec<String>,
    },
}

#[derive(Debug, Clone)]
struct CachedDisplaySegment {
    run_index: usize,
    text: String,
    properties: Option<CT_RPr>,
}

/// The representation used to serialize a Word field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldForm {
    /// One `w:fldSimple` element.
    Simple,
    /// Ordered begin, instruction, separate, result and end runs.
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
        /// Custom mark displayed by this reference instead of a numeric label.
        custom_mark: Option<String>,
    },
    /// An endnote reference (`<w:endnoteReference w:id="..."/>`).
    EndnoteRef {
        id: i32,
        /// Custom mark displayed by this reference instead of a numeric label.
        custom_mark: Option<String>,
    },
    /// A comment reference (`<w:commentReference w:id="..."/>`).
    CommentReference {
        id: i32,
        /// Number of raw run children that precede this reference.
        raw_before: usize,
    },
    /// A symbol character from a specific font (`<w:sym w:font="..." w:char="..."/>`).
    Symbol {
        /// The font the code point is looked up in.
        font: String,
        /// The code point, written back as four upper-case hex digits.
        char_code: u16,
    },
    /// One of the Word special characters that carries no text of its own.
    SpecialCharacter(SpecialCharacter),
}

/// A Word special character run child.
///
/// The four share one placement contract, so they share one `RunContent`
/// variant rather than taking one each across the ten files that match on it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpecialCharacter {
    /// `<w:cr/>`, a carriage return that breaks the line.
    CarriageReturn,
    /// `<w:noBreakHyphen/>`, a hyphen that is not a break opportunity.
    NoBreakHyphen,
    /// `<w:softHyphen/>`, a hyphen that renders only at a line break.
    SoftHyphen,
    /// `<w:ptab/>`, an absolutely positioned tab.
    PositionalTab {
        alignment: ST_PTabAlignment,
        relative_to: ST_PTabRelativeTo,
        leader: ST_PTabLeader,
    },
}

/// `ST_PTabAlignment` — how content aligns against a positional tab.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_PTabAlignment {
    Left,
    Center,
    Right,
}

/// `ST_PTabRelativeTo` — what a positional tab position is measured from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_PTabRelativeTo {
    Margin,
    Indent,
}

/// `ST_PTabLeader` — the leader drawn across a positional tab gap.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_PTabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
    MiddleDot,
}

impl ST_PTabAlignment {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "left" => Some(ST_PTabAlignment::Left),
            "center" => Some(ST_PTabAlignment::Center),
            "right" => Some(ST_PTabAlignment::Right),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            ST_PTabAlignment::Left => "left",
            ST_PTabAlignment::Center => "center",
            ST_PTabAlignment::Right => "right",
        }
    }
}

impl ST_PTabRelativeTo {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "margin" => Some(ST_PTabRelativeTo::Margin),
            "indent" => Some(ST_PTabRelativeTo::Indent),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            ST_PTabRelativeTo::Margin => "margin",
            ST_PTabRelativeTo::Indent => "indent",
        }
    }
}

impl ST_PTabLeader {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "none" => Some(ST_PTabLeader::None),
            "dot" => Some(ST_PTabLeader::Dot),
            "hyphen" => Some(ST_PTabLeader::Hyphen),
            "underscore" => Some(ST_PTabLeader::Underscore),
            "middleDot" => Some(ST_PTabLeader::MiddleDot),
            _ => None,
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            ST_PTabLeader::None => "none",
            ST_PTabLeader::Dot => "dot",
            ST_PTabLeader::Hyphen => "hyphen",
            ST_PTabLeader::Underscore => "underscore",
            ST_PTabLeader::MiddleDot => "middleDot",
        }
    }
}

/// Read `<w:sym/>` as typed content.
///
/// A symbol whose `w:char` is not four hex digits, or that carries an
/// attribute outside the modeled pair, stays raw so nothing is lost.
fn parse_symbol(e: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<RunContent>> {
    let mut font = None;
    let mut char_code = None;
    for attribute in e.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let value = std::str::from_utf8(&attribute.value)?;
        if is_word_attribute(key, b"font", prefixes) {
            font = Some(value.to_owned());
        } else if is_word_attribute(key, b"char", prefixes) {
            if value.len() != 4 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Ok(None);
            }
            char_code = u16::from_str_radix(value, 16).ok();
        } else if !key.starts_with(b"xmlns") {
            return Ok(None);
        }
    }
    Ok(match (font, char_code) {
        (Some(font), Some(char_code)) => Some(RunContent::Symbol { font, char_code }),
        _ => None,
    })
}

/// Read `<w:ptab/>` as typed content.
///
/// All three attributes are required by the schema, so a positional tab
/// missing one, or carrying a token outside the inventory, stays raw.
fn parse_positional_tab(e: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<RunContent>> {
    let mut alignment = None;
    let mut relative_to = None;
    let mut leader = None;
    for attribute in e.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let value = std::str::from_utf8(&attribute.value)?;
        if is_word_attribute(key, b"alignment", prefixes) {
            alignment = ST_PTabAlignment::from_str(value);
        } else if is_word_attribute(key, b"relativeTo", prefixes) {
            relative_to = ST_PTabRelativeTo::from_str(value);
        } else if is_word_attribute(key, b"leader", prefixes) {
            leader = ST_PTabLeader::from_str(value);
        } else if !key.starts_with(b"xmlns") {
            return Ok(None);
        }
    }
    Ok(match (alignment, relative_to, leader) {
        (Some(alignment), Some(relative_to), Some(leader)) => Some(RunContent::SpecialCharacter(
            SpecialCharacter::PositionalTab {
                alignment,
                relative_to,
                leader,
            },
        )),
        _ => None,
    })
}

/// Read `<w:cr/>`, `<w:noBreakHyphen/>` and `<w:softHyphen/>` as typed content.
///
/// These are `CT_Empty`, so an element carrying any attribute other than a
/// namespace declaration stays raw.
fn parse_empty_special_character(
    e: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<Option<RunContent>> {
    let name = e.name();
    let character = if is_word_element(name.as_ref(), b"cr", prefixes) {
        SpecialCharacter::CarriageReturn
    } else if is_word_element(name.as_ref(), b"noBreakHyphen", prefixes) {
        SpecialCharacter::NoBreakHyphen
    } else if is_word_element(name.as_ref(), b"softHyphen", prefixes) {
        SpecialCharacter::SoftHyphen
    } else {
        return Ok(None);
    };
    for attribute in e.attributes() {
        if !attribute?.key.as_ref().starts_with(b"xmlns") {
            return Ok(None);
        }
    }
    Ok(Some(RunContent::SpecialCharacter(character)))
}

/// Read the run children F-265 typed, or `None` to keep the element raw.
fn parse_typed_special_run_child(
    e: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<Option<RunContent>> {
    if is_word_element(e.name().as_ref(), b"sym", prefixes) {
        parse_symbol(e, prefixes)
    } else if is_word_element(e.name().as_ref(), b"ptab", prefixes) {
        parse_positional_tab(e, prefixes)
    } else {
        parse_empty_special_character(e, prefixes)
    }
}

fn write_special_character<W: std::io::Write>(
    writer: &mut Writer<W>,
    character: SpecialCharacter,
) -> Result<()> {
    match character {
        SpecialCharacter::CarriageReturn => {
            writer.write_event(Event::Empty(BytesStart::new("w:cr")))?;
        }
        SpecialCharacter::NoBreakHyphen => {
            writer.write_event(Event::Empty(BytesStart::new("w:noBreakHyphen")))?;
        }
        SpecialCharacter::SoftHyphen => {
            writer.write_event(Event::Empty(BytesStart::new("w:softHyphen")))?;
        }
        SpecialCharacter::PositionalTab {
            alignment,
            relative_to,
            leader,
        } => {
            let mut e = BytesStart::new("w:ptab");
            e.push_attribute(("w:alignment", alignment.as_str()));
            e.push_attribute(("w:relativeTo", relative_to.as_str()));
            e.push_attribute(("w:leader", leader.as_str()));
            writer.write_event(Event::Empty(e))?;
        }
    }
    Ok(())
}

/// A typed comment range boundary at a run insertion point.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CommentRangeMarker {
    Start {
        id: i32,
        run_index: usize,
        /// Number of raw children at this run boundary that precede the marker.
        raw_before: usize,
        /// Whether the source marker contained a child element or visible text.
        has_child_content: bool,
    },
    End {
        id: i32,
        run_index: usize,
        /// Number of raw children at this run boundary that precede the marker.
        raw_before: usize,
        /// Whether the source marker contained a child element or visible text.
        has_child_content: bool,
    },
}

/// Read projection of a bookmark marker retained at a run boundary.
#[derive(Clone)]
pub struct BookmarkMarker {
    start: bool,
    id: Option<i32>,
    name: Option<String>,
    run_index: usize,
    raw_before: usize,
    projected_run_index: usize,
    tracked_run_index: usize,
    word_prefixes: Vec<String>,
    has_child_content: bool,
}

impl PartialEq for BookmarkMarker {
    fn eq(&self, other: &Self) -> bool {
        self.start == other.start
            && self.id == other.id
            && self.name == other.name
            && self.run_index == other.run_index
            && self.raw_before == other.raw_before
            && self.projected_run_index == other.projected_run_index
            && self.tracked_run_index == other.tracked_run_index
            && self.has_child_content == other.has_child_content
    }
}

impl Eq for BookmarkMarker {}

impl std::fmt::Debug for BookmarkMarker {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("BookmarkMarker")
            .field("start", &self.start)
            .field("id", &self.id)
            .field("name", &self.name)
            .field("run_index", &self.run_index)
            .field("raw_before", &self.raw_before)
            .field("projected_run_index", &self.projected_run_index)
            .field("tracked_run_index", &self.tracked_run_index)
            .field("has_child_content", &self.has_child_content)
            .finish()
    }
}

impl BookmarkMarker {
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

    /// Whether the source marker contained a child element or visible text.
    pub fn has_child_content(&self) -> bool {
        self.has_child_content
    }

    /// Accepted-view run boundary represented by this marker.
    #[doc(hidden)]
    pub fn projected_run_index(&self) -> usize {
        self.projected_run_index
    }

    /// Tracked-view run boundary represented by this marker.
    #[doc(hidden)]
    pub fn tracked_run_index(&self) -> usize {
        self.tracked_run_index
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
const RAW_ALTERNATE_CONTENT_DRAWING_FLAG: usize = 1usize << (usize::BITS - 2);
const RAW_ROOT_ATTRIBUTES_FLAG: usize = 1usize << (usize::BITS - 3);
const RAW_COMMENT_REFERENCE_FLAG: usize = 1usize << (usize::BITS - 4);
const RAW_CHILD_FLAGS: usize = RAW_LEGACY_HORIZONTAL_RULE_FLAG
    | RAW_ALTERNATE_CONTENT_DRAWING_FLAG
    | RAW_ROOT_ATTRIBUTES_FLAG
    | RAW_COMMENT_REFERENCE_FLAG;
const RAW_CHILD_POSITION_MASK: usize = !RAW_CHILD_FLAGS;
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

    pub(crate) fn from_empty_root(root: &BytesStart<'_>, word_prefixes: &[String]) -> Result<Self> {
        let mut run = Self {
            properties: None,
            content: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        };
        if let Some(record) = capture_root_attribute_record(root, word_prefixes)? {
            run.extra_xml.push(record);
            run.extra_xml_positions.push(RAW_ROOT_ATTRIBUTES_FLAG);
        }
        Ok(run)
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

    /// Report whether a raw carrier retains attributes from the run root.
    #[doc(hidden)]
    pub fn raw_child_is_root_attributes(encoded: usize) -> bool {
        encoded & RAW_ROOT_ATTRIBUTES_FLAG != 0
    }

    /// Report whether this raw slot has one selected drawing projection.
    #[doc(hidden)]
    pub fn raw_child_has_alternate_content_drawing(encoded: usize) -> bool {
        encoded & RAW_ALTERNATE_CONTENT_DRAWING_FLAG != 0
    }

    /// Replace a raw-child boundary without changing its parsed classification.
    #[doc(hidden)]
    pub fn set_raw_child_position(encoded: &mut usize, position: usize) {
        debug_assert_eq!(position & RAW_CHILD_FLAGS, 0);
        *encoded = (position & RAW_CHILD_POSITION_MASK) | (*encoded & RAW_CHILD_FLAGS);
    }

    fn encode_raw_child_position(
        position: usize,
        legacy_horizontal_rule: bool,
        alternate_content_drawing: bool,
    ) -> usize {
        debug_assert_eq!(position & RAW_CHILD_FLAGS, 0);
        position
            | if legacy_horizontal_rule {
                RAW_LEGACY_HORIZONTAL_RULE_FLAG
            } else {
                0
            }
            | if alternate_content_drawing {
                RAW_ALTERNATE_CONTENT_DRAWING_FLAG
            } else {
                0
            }
    }

    /// Get the combined text of all text content in this run.
    pub fn text(&self) -> String {
        self.content.iter().map(Self::content_text).collect()
    }

    /// Whether this run shows anything: text other than white space, or a
    /// child that draws without text, such as a drawing, a picture, an
    /// object, a field, a symbol, a tab, a break or a note reference.
    /// Deleted text, comment references, soft hyphens, field characters and
    /// the last rendered page break show nothing.
    fn has_visible_content(&self) -> bool {
        let visible_item = |content: &RunContent| match content {
            RunContent::Text(text) => !text.text.trim().is_empty(),
            RunContent::DeletedText(_)
            | RunContent::CommentReference { .. }
            | RunContent::SpecialCharacter(SpecialCharacter::SoftHyphen) => false,
            _ => true,
        };
        let visible_raw = |(raw, position): (&Vec<u8>, &usize)| {
            *position & (RAW_ROOT_ATTRIBUTES_FLAG | RAW_COMMENT_REFERENCE_FLAG) == 0
                && !is_xml_whitespace(raw)
                && !raw_root_is_one_of(
                    raw,
                    &[
                        b"fldChar",
                        b"instrText",
                        b"delInstrText",
                        b"lastRenderedPageBreak",
                    ],
                )
        };
        self.content.iter().any(visible_item)
            || self
                .extra_xml
                .iter()
                .zip(&self.extra_xml_positions)
                .any(visible_raw)
    }

    /// The text one content item contributes to [`Self::text`].
    fn content_text(content: &RunContent) -> &str {
        match content {
            RunContent::Text(t) | RunContent::DeletedText(t) => &t.text,
            RunContent::Tab => "\t",
            RunContent::Break(_) => "\n",
            RunContent::Drawing(_) => "", // Drawings have no text content
            RunContent::Field(field) => field.projected_text().unwrap_or(""),
            RunContent::FootnoteRef { custom_mark, .. }
            | RunContent::EndnoteRef { custom_mark, .. } => custom_mark.as_deref().unwrap_or(""),
            RunContent::CommentReference { .. } => "",
            // A symbol is font-encoded rather than Unicode, and the special
            // characters carry no text, so neither contributes to extraction.
            RunContent::Symbol { .. } => "",
            RunContent::SpecialCharacter(SpecialCharacter::CarriageReturn) => "\n",
            RunContent::SpecialCharacter(_) => "",
        }
    }

    /// The literal text that contributes to a public split offset.
    fn literal_text(content: &RunContent) -> &str {
        match content {
            RunContent::Text(text) | RunContent::DeletedText(text) => &text.text,
            RunContent::Tab
            | RunContent::Break(_)
            | RunContent::Drawing(_)
            | RunContent::Field(_)
            | RunContent::FootnoteRef { .. }
            | RunContent::EndnoteRef { .. }
            | RunContent::CommentReference { .. }
            | RunContent::Symbol { .. }
            | RunContent::SpecialCharacter(_) => "",
        }
    }

    fn literal_len(&self) -> usize {
        self.content
            .iter()
            .map(Self::literal_text)
            .map(str::chars)
            .map(Iterator::count)
            .sum()
    }

    /// Split this run at a Unicode scalar offset of its literal text.
    ///
    /// This run keeps the content before `offset`, and the returned run, with
    /// cloned properties, receives the rest. Text-free content and raw children
    /// at the split point stay in this run. Nothing changes on error.
    fn split_off_at(&mut self, offset: usize) -> std::result::Result<CT_R, RunSplitError> {
        let len = self.literal_len();
        if offset == 0 || offset >= len {
            return Err(RunSplitError::OffsetOutOfRange { offset, len });
        }
        // The first item whose text reaches past the offset, and how many of
        // its characters come before the offset.
        let mut end = 0;
        let (index, within, width) = self
            .content
            .iter()
            .enumerate()
            .find_map(|(index, content)| {
                let start = end;
                let width = Self::literal_text(content).chars().count();
                end += width;
                (width > 0 && end >= offset).then_some((index, offset - start, width))
            })
            .expect("an offset inside the run text falls inside one content item");
        let tail_content = if within == width {
            self.content.split_off(index + 1)
        } else {
            let rest = match &mut self.content[index] {
                RunContent::Text(text) => RunContent::Text(text.split_off_at(within)),
                RunContent::DeletedText(text) => RunContent::DeletedText(text.split_off_at(within)),
                _ => unreachable!("an interior literal offset falls inside literal text"),
            };
            let mut tail = self.content.split_off(index + 1);
            tail.insert(0, rest);
            tail
        };
        let removed_prefix_content = index + usize::from(within == width);

        let mut tail = CT_R {
            properties: self.properties.clone(),
            content: tail_content,
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        };
        // A raw child position counts the modeled children before it, so a
        // child after the first `index` content items moves to the tail.
        let head_boundary = usize::from(self.properties.is_some()) + index;
        if self.extra_xml_positions.len() == self.extra_xml.len() {
            let raw_children = std::mem::take(&mut self.extra_xml);
            let positions = std::mem::take(&mut self.extra_xml_positions);
            let mut alternate_drawings = std::mem::take(&mut self.alt_drawings).into_iter();
            for (raw, mut encoded) in raw_children.into_iter().zip(positions) {
                let position = Self::raw_child_position(encoded);
                let alternate_drawing = Self::raw_child_has_alternate_content_drawing(encoded)
                    .then(|| alternate_drawings.next())
                    .flatten();
                if position <= head_boundary {
                    self.extra_xml.push(raw);
                    self.extra_xml_positions.push(encoded);
                    self.alt_drawings.extend(alternate_drawing);
                } else {
                    Self::set_raw_child_position(&mut encoded, position - removed_prefix_content);
                    tail.extra_xml.push(raw);
                    tail.extra_xml_positions.push(encoded);
                    tail.alt_drawings.extend(alternate_drawing);
                }
            }
            self.alt_drawings.extend(alternate_drawings);
        }
        let head_raw_children = self.extra_xml.len();
        for content in &mut tail.content {
            if let RunContent::CommentReference { raw_before, .. } = content {
                *raw_before = raw_before.saturating_sub(head_raw_children);
            }
        }
        Ok(tail)
    }

    /// Replace typed run content while retaining every raw child boundary.
    #[doc(hidden)]
    pub fn replace_content(&mut self, content: Vec<RunContent>) {
        self.retain_comment_reference_sources(None);
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
        self.retain_comment_reference_sources(Some(removed));
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

    fn retain_comment_reference_sources(&mut self, removed: Option<&[bool]>) {
        if self.extra_xml.len() != self.extra_xml_positions.len() {
            return;
        }
        let property_boundary = usize::from(self.properties.is_some());
        let keep = self
            .extra_xml_positions
            .iter()
            .map(|encoded| {
                *encoded & RAW_COMMENT_REFERENCE_FLAG == 0
                    || removed.is_some_and(|removed| {
                        !removed
                            .get(
                                Self::raw_child_position(*encoded)
                                    .saturating_sub(property_boundary),
                            )
                            .copied()
                            .unwrap_or(true)
                    })
            })
            .collect::<Vec<_>>();
        let mut index = 0;
        self.extra_xml.retain(|_| {
            let retain = keep[index];
            index += 1;
            retain
        });
        let mut index = 0;
        self.extra_xml_positions.retain(|_| {
            let retain = keep[index];
            index += 1;
            retain
        });
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
        Self::from_xml_with_prefixes_and_root(reader, word_prefixes, None)
    }

    pub(crate) fn from_xml_with_prefixes_and_root(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        root: Option<&BytesStart<'_>>,
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
                        match content.last_mut() {
                            Some(RunContent::FootnoteRef {
                                custom_mark: Some(mark),
                                ..
                            })
                            | Some(RunContent::EndnoteRef {
                                custom_mark: Some(mark),
                                ..
                            }) if mark.is_empty() => *mark = text,
                            _ => content.push(RunContent::Text(CT_Text {
                                text,
                                preserve_space: preserve,
                            })),
                        }
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
                        let raw = capture_element(reader, e)?;
                        let mut source = Reader::from_reader(raw.as_slice());
                        source.read_event()?;
                        let has_payload = !matches!(source.read_event()?, Event::End(_));
                        if has_payload || comment_reference_has_payload_attributes(e, &prefixes) {
                            extra_xml.push(close_comment_reference_source(&raw, &prefixes)?);
                            extra_xml_positions.push(modeled_children | RAW_COMMENT_REFERENCE_FLAG);
                        }
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
                        let drawing = crate::drawing::parse_alternate_content(&raw, &prefixes);
                        let has_drawing = drawing.is_some();
                        if let Some(drawing) = drawing {
                            alt_drawings.push(drawing);
                        }
                        extra_xml.push(raw);
                        extra_xml_positions.push(Self::encode_raw_child_position(
                            modeled_children,
                            false,
                            has_drawing,
                        ));
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
                            false,
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
                        let custom_mark =
                            optional_word_attribute(e, b"customMarkFollows", &prefixes)
                                .filter(|value| value == "1" || value == "true")
                                .map(|_| String::new());
                        content.push(RunContent::FootnoteRef { id, custom_mark });
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"endnoteReference", &prefixes) {
                        let id = optional_word_attribute(e, b"id", &prefixes)
                            .and_then(|value| value.parse::<i32>().ok())
                            .unwrap_or(0);
                        let custom_mark =
                            optional_word_attribute(e, b"customMarkFollows", &prefixes)
                                .filter(|value| value == "1" || value == "true")
                                .map(|_| String::new());
                        content.push(RunContent::EndnoteRef { id, custom_mark });
                        modeled_children += 1;
                    } else if is_word_element(name.as_ref(), b"commentReference", &prefixes) {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        if comment_reference_has_payload_attributes(e, &prefixes) {
                            extra_xml.push(close_comment_reference_source(
                                &capture_empty_element(e)?,
                                &prefixes,
                            )?);
                            extra_xml_positions.push(modeled_children | RAW_COMMENT_REFERENCE_FLAG);
                        }
                        content.push(RunContent::CommentReference {
                            id,
                            raw_before: extra_xml.len(),
                        });
                        modeled_children += 1;
                    } else if let Some(typed) = parse_typed_special_run_child(e, &prefixes)? {
                        content.push(typed);
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
                Ok(event @ (Event::Comment(_) | Event::PI(_))) => {
                    extra_xml.push(capture_standalone_event(event.into_owned())?);
                    extra_xml_positions.push(modeled_children);
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

        if let Some(record) = root
            .map(|root| capture_root_attribute_record(root, word_prefixes))
            .transpose()?
            .flatten()
        {
            extra_xml.push(record);
            extra_xml_positions.push(RAW_ROOT_ATTRIBUTES_FLAG);
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
        for (position, raw) in self.extra_xml_positions.iter().zip(&self.extra_xml) {
            if Self::raw_child_is_root_attributes(*position) {
                push_root_attribute_record(&mut run, raw, None)?;
            }
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
                RunContent::FootnoteRef { id, custom_mark } => {
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:footnoteReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    if custom_mark.is_some() {
                        e.push_attribute(("w:customMarkFollows", "1"));
                    }
                    writer.write_event(Event::Empty(e))?;
                    if let Some(mark) = custom_mark {
                        let mut text = BytesStart::new("w:t");
                        text.push_attribute(("xml:space", "preserve"));
                        writer.write_event(Event::Start(text))?;
                        writer.write_event(Event::Text(quick_xml::events::BytesText::new(mark)))?;
                        writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                    }
                }
                RunContent::EndnoteRef { id, custom_mark } => {
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:endnoteReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    if custom_mark.is_some() {
                        e.push_attribute(("w:customMarkFollows", "1"));
                    }
                    writer.write_event(Event::Empty(e))?;
                    if let Some(mark) = custom_mark {
                        let mut text = BytesStart::new("w:t");
                        text.push_attribute(("xml:space", "preserve"));
                        writer.write_event(Event::Start(text))?;
                        writer.write_event(Event::Text(quick_xml::events::BytesText::new(mark)))?;
                        writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                    }
                }
                RunContent::Symbol { font, char_code } => {
                    let mut e = BytesStart::new("w:sym");
                    e.push_attribute(("w:font", font.as_str()));
                    e.push_attribute(("w:char", format!("{char_code:04X}").as_str()));
                    writer.write_event(Event::Empty(e))?;
                }
                RunContent::SpecialCharacter(character) => {
                    write_special_character(writer, *character)?;
                }
                RunContent::CommentReference { id, raw_before } => {
                    if !ordered_raw {
                        for (index, raw) in self
                            .extra_xml
                            .iter()
                            .enumerate()
                            .take((*raw_before).min(self.extra_xml.len()))
                            .skip(raw_written)
                        {
                            if !self
                                .extra_xml_positions
                                .get(index)
                                .is_some_and(|position| *position & RAW_COMMENT_REFERENCE_FLAG != 0)
                            {
                                write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
                            }
                            raw_written += 1;
                        }
                    }
                    if let Some(raw) = self
                        .extra_xml_positions
                        .iter()
                        .zip(&self.extra_xml)
                        .find_map(|(position, raw)| {
                            (*position & RAW_COMMENT_REFERENCE_FLAG != 0
                                && Self::raw_child_position(*position) == typed_boundary
                                && retained_comment_reference_id(raw) == Some(*id))
                            .then_some(raw)
                        })
                    {
                        write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
                        typed_boundary += 1;
                        if ordered_raw {
                            write_run_raw_boundary(
                                writer,
                                self,
                                typed_boundary,
                                foreign_word_namespace,
                            )?;
                        }
                        continue;
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
            for (index, raw) in self.extra_xml.iter().enumerate().skip(raw_written) {
                if is_root_attribute_record(raw)
                    || self
                        .extra_xml_positions
                        .get(index)
                        .is_some_and(|position| *position & RAW_COMMENT_REFERENCE_FLAG != 0)
                {
                    continue;
                }
                write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        Ok(())
    }
}

fn comment_reference_has_payload_attributes(element: &BytesStart<'_>, prefixes: &[String]) -> bool {
    let name = element.name();
    let own_binding = qualified_name_prefix(name.as_ref())
        .map_or_else(|| "xmlns".to_owned(), |prefix| format!("xmlns:{prefix}"));
    element.attributes().any(|attribute| {
        attribute.ok().is_none_or(|attribute| {
            if attribute_in_namespace(
                attribute.key.as_ref(),
                b"id",
                crate::namespace::W_NS,
                prefixes,
            ) {
                return false;
            }
            // The element's own Word alias declaration is structural, not opaque
            // payload. Every other declaration or attribute retains its source.
            !(attribute.key.as_ref() == own_binding.as_bytes()
                && attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .is_ok_and(|value| value == crate::namespace::W_NS))
        })
    })
}

fn close_comment_reference_source(raw: &[u8], prefixes: &[String]) -> Result<Vec<u8>> {
    let mut bindings = prefixes
        .iter()
        .filter_map(|binding| {
            binding
                .strip_prefix('\0')
                .and_then(|binding| binding.split_once('\0'))
                .map(|(prefix, uri)| (prefix.to_owned(), uri.to_owned()))
        })
        .collect::<Vec<_>>();
    for prefix in prefixes.iter().filter(|prefix| !prefix.starts_with('\0')) {
        if !bindings.iter().any(|(bound, _)| bound == prefix) {
            bindings.push((prefix.clone(), crate::namespace::W_NS.to_owned()));
        }
    }
    raw_with_external_bindings(raw, &bindings)
}

fn retained_comment_reference_id(raw: &[u8]) -> Option<i32> {
    let mut reader = quick_xml::reader::NsReader::from_reader(raw);
    let (ns, event) = reader.read_resolved_event().ok()?;
    if !matches!(ns, ResolveResult::Bound(namespace) if namespace.as_ref() == crate::namespace::W_NS.as_bytes())
    {
        return None;
    }
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return None,
    };
    if element.local_name().as_ref() != b"commentReference" {
        return None;
    }
    element.attributes().find_map(|attr| {
        let attr = attr.ok()?;
        let (ns, local) = reader.resolver().resolve_attribute(attr.key);
        (matches!(ns, ResolveResult::Bound(namespace) if namespace.as_ref() == crate::namespace::W_NS.as_bytes())
            && local.as_ref() == b"id").then(|| attr.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder()).ok()?.parse().ok()).flatten()
    })
}

fn write_run_raw_boundary<W: std::io::Write>(
    writer: &mut Writer<W>,
    run: &CT_R,
    boundary: usize,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    for (position, raw) in run.extra_xml_positions.iter().zip(&run.extra_xml) {
        if !CT_R::raw_child_is_root_attributes(*position)
            && *position & RAW_COMMENT_REFERENCE_FLAG == 0
            && CT_R::raw_child_position(*position) == boundary
        {
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
    bookmark_markers: Vec<BookmarkMarker>,
    projected_run_count: usize,
    tracked_run_count: usize,
}

#[derive(Default)]
struct AcceptedBookmarkProjection {
    markers: Vec<BookmarkMarker>,
    projected_run_count: usize,
    tracked_run_count: usize,
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

pub(crate) fn parse_run_raw(raw: &[u8], word_prefixes: &[String]) -> Result<CT_R> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"r") => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                return CT_R::from_xml_with_prefixes_and_root(&mut reader, &prefixes, Some(&start));
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

    let result_start = reader.buffer_position() as usize;
    let mut result_end = result_start;
    let mut cached_result = String::new();
    let mut cached_segments = Vec::new();
    let mut result_run_index = 0usize;
    loop {
        buffer.clear();
        let before = reader.buffer_position() as usize;
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
            Event::End(end) if matches_local_name(end.name().as_ref(), b"fldSimple") => {
                result_end = before;
                break;
            }
            Event::Eof => break,
            _ => {}
        }
    }

    let instruction = parse_field_instruction(&instruction);
    if instruction.name.is_empty() {
        return Ok(None);
    }
    let mut field = Field::parsed(
        instruction,
        cached_result,
        cached_segments,
        dirty,
        FieldForm::Simple,
        raw.to_vec(),
        prefixes.clone(),
    );
    let result_xml = &raw[result_start..result_end];
    // The typed pass still decides namespace and structure. This cheap absence
    // test avoids reparsing ordinary text-only caches.
    let possible_nested = result_xml.windows(7).any(|bytes| bytes == b"fldChar")
        || result_xml.windows(9).any(|bytes| bytes == b"fldSimple")
        || source_has_note_reference_name(result_xml);
    if possible_nested {
        // Reuse a Word prefix from the effective producer scope. Introducing
        // a canonical w declaration here would reclassify foreign w children.
        let prefix = prefixes
            .iter()
            .find(|prefix| !prefix.starts_with('\0'))
            .ok_or_else(|| OxmlError::InvalidValue("simple field has no Word prefix".to_owned()))?;
        let name = if prefix.is_empty() {
            "p".to_owned()
        } else {
            format!("{prefix}:p")
        };
        let mut cache_xml = format!("<{name}>").into_bytes();
        cache_xml.extend_from_slice(&raw[result_start..result_end]);
        cache_xml.extend_from_slice(format!("</{name}>").as_bytes());
        let mut cache_reader = Reader::from_reader(cache_xml.as_slice());
        let mut cache_buffer = Vec::new();
        let Event::Start(cache_root) = cache_reader.read_event_into(&mut cache_buffer)? else {
            unreachable!("synthetic cache root");
        };
        let cache_prefixes = word_prefixes_at(&cache_root, &prefixes)?;
        let paragraph = CT_P::from_xml_with_prefixes_and_root(
            &mut cache_reader,
            &cache_prefixes,
            Some(&cache_root),
        )?;
        let runs = paragraph.runs().into_iter().cloned().collect::<Vec<_>>();
        if runs.iter().flat_map(|run| &run.content).any(|content| {
            matches!(
                content,
                RunContent::Field(_)
                    | RunContent::FootnoteRef { .. }
                    | RunContent::EndnoteRef { .. }
                    | RunContent::CommentReference { .. }
            )
        }) {
            field.cached_result = runs.iter().map(cached_run_display).collect();
            if let FieldSource::Parsed {
                original_cached_result,
                ..
            } = &mut field.source
            {
                *original_cached_result = field.cached_result.clone();
            }
            field.parsed_cached_comment_ranges = paragraph.comment_ranges.clone();
            field.simple_cached_runs = Some(runs);
        }
    }
    Ok(Some(field))
}

/// Whether a preserved simple field is admitted by the complete typed parser.
#[doc(hidden)]
pub fn story_simple_field_is_typed(raw: &[u8], word_prefixes: &[String]) -> bool {
    parse_simple_field(raw, word_prefixes)
        .ok()
        .flatten()
        .is_some()
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

#[derive(Default)]
struct LegacyFormProjection {
    ff_data_last_slot: Option<usize>,
    kind_last_slot: Option<usize>,
    name: Option<String>,
    name_seen: bool,
    enabled: Option<bool>,
    enabled_seen: bool,
    calculate_on_exit: Option<bool>,
    calculate_on_exit_seen: bool,
    entry_macro_seen: bool,
    exit_macro_seen: bool,
    help_text_seen: bool,
    status_text_seen: bool,
    kind: Option<LegacyFormFieldKind>,
    text_type_seen: bool,
    text_format_seen: bool,
    checkbox_size_seen: bool,
    default_text: Option<String>,
    default_text_seen: bool,
    checked: Option<bool>,
    checked_seen: bool,
    check_default: Option<bool>,
    check_default_seen: bool,
    selected_index: Option<usize>,
    selected_index_seen: bool,
    selected_default: Option<usize>,
    selected_default_seen: bool,
    choices: Vec<String>,
    max_length: Option<usize>,
    max_length_seen: bool,
}

fn parse_legacy_form_data(
    raw: &[u8],
    word_prefixes: &[String],
    instruction_name: &str,
    cached_result: &str,
) -> Result<Option<LegacyFormFieldData>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<Option<String>>::new();
    let mut prefixes = word_prefixes.to_vec();
    let mut prefix_scopes = Vec::new();
    let mut begin_depth = None;
    let mut ff_data_depth = None;
    let mut ff_data_count = 0usize;
    let mut kind_depth = None;
    let mut leaf_depth = None;
    let mut projection = LegacyFormProjection::default();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                if leaf_depth.is_some() {
                    return Err(OxmlError::InvalidValue(
                        "legacy form leaf property contains an element".to_owned(),
                    ));
                }
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let word = is_word_element(
                    element.name().as_ref(),
                    local_name(element.name().as_ref()),
                    &local_prefixes,
                );
                let local = word.then(|| {
                    String::from_utf8_lossy(local_name(element.name().as_ref())).into_owned()
                });
                stack.push(local.clone());
                prefix_scopes.push(std::mem::replace(&mut prefixes, local_prefixes));
                let depth = stack.len() - 1;
                if local.as_deref() == Some("fldChar")
                    && begin_depth.is_none()
                    && optional_word_attribute(&element, b"fldCharType", &prefixes).as_deref()
                        == Some("begin")
                {
                    begin_depth = Some(depth);
                } else if local.as_deref() == Some("ffData") && begin_depth == depth.checked_sub(1)
                {
                    ff_data_count += 1;
                    if ff_data_count > 1 {
                        return Err(OxmlError::InvalidValue(
                            "duplicate direct w:ffData owner".to_owned(),
                        ));
                    }
                    ff_data_depth = Some(depth);
                } else if let Some(ff_depth) = ff_data_depth {
                    leaf_depth = project_legacy_form_element(
                        &element,
                        local.as_deref(),
                        depth,
                        ff_depth,
                        &prefixes,
                        &mut kind_depth,
                        &mut projection,
                    )?
                    .then_some(depth);
                }
            }
            Event::Empty(element) => {
                if leaf_depth.is_some() {
                    return Err(OxmlError::InvalidValue(
                        "legacy form leaf property contains an element".to_owned(),
                    ));
                }
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let word = is_word_element(
                    element.name().as_ref(),
                    local_name(element.name().as_ref()),
                    &local_prefixes,
                );
                let local = word.then(|| {
                    String::from_utf8_lossy(local_name(element.name().as_ref())).into_owned()
                });
                let depth = stack.len();
                if local.as_deref() == Some("fldChar")
                    && begin_depth.is_none()
                    && optional_word_attribute(&element, b"fldCharType", &local_prefixes).as_deref()
                        == Some("begin")
                {
                    begin_depth = Some(depth);
                } else if local.as_deref() == Some("ffData") && begin_depth == depth.checked_sub(1)
                {
                    ff_data_count += 1;
                    if ff_data_count > 1 {
                        return Err(OxmlError::InvalidValue(
                            "duplicate direct w:ffData owner".to_owned(),
                        ));
                    }
                } else if let Some(ff_depth) = ff_data_depth {
                    project_legacy_form_element(
                        &element,
                        local.as_deref(),
                        depth,
                        ff_depth,
                        &local_prefixes,
                        &mut kind_depth,
                        &mut projection,
                    )?;
                    if depth == ff_depth + 1
                        && matches!(local.as_deref(), Some("checkBox" | "ddList" | "textInput"))
                    {
                        kind_depth = None;
                    }
                }
            }
            Event::End(_) => {
                let depth = stack.len().saturating_sub(1);
                if leaf_depth == Some(depth) {
                    leaf_depth = None;
                }
                if kind_depth == Some(depth) {
                    kind_depth = None;
                }
                if ff_data_depth == Some(depth) {
                    ff_data_depth = None;
                }
                let closes_begin = begin_depth == Some(depth);
                stack.pop();
                prefixes = prefix_scopes
                    .pop()
                    .unwrap_or_else(|| word_prefixes.to_vec());
                if closes_begin {
                    break;
                }
            }
            Event::Text(text)
                if (leaf_depth.is_some()
                    || is_legacy_form_element_only_container(
                        stack.len(),
                        ff_data_depth,
                        kind_depth,
                    ))
                    && text.iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(OxmlError::InvalidValue(
                    "legacy form element-only content contains character data".to_owned(),
                ));
            }
            Event::CData(text)
                if leaf_depth.is_some()
                    || (is_legacy_form_element_only_container(
                        stack.len(),
                        ff_data_depth,
                        kind_depth,
                    ) && text.iter().any(|byte| !byte.is_ascii_whitespace())) =>
            {
                return Err(OxmlError::InvalidValue(
                    "legacy form element-only content contains character data".to_owned(),
                ));
            }
            Event::GeneralRef(reference) => {
                let in_element_only_container =
                    is_legacy_form_element_only_container(stack.len(), ff_data_depth, kind_depth);
                let non_whitespace = match reference.resolve_char_ref() {
                    Ok(Some(character)) => !character.is_ascii_whitespace(),
                    Ok(None) | Err(_) => true,
                };
                if leaf_depth.is_some() || (in_element_only_container && non_whitespace) {
                    return Err(OxmlError::InvalidValue(
                        "legacy form element-only content contains character data".to_owned(),
                    ));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let Some(kind) = projection.kind else {
        if ff_data_count == 0 {
            return Ok(None);
        }
        return Err(OxmlError::MissingElement("w:ffData form kind".to_owned()));
    };
    let matches_instruction = matches!(
        (kind, instruction_name),
        (LegacyFormFieldKind::TextInput, "FORMTEXT")
            | (LegacyFormFieldKind::CheckBox, "FORMCHECKBOX")
            | (LegacyFormFieldKind::DropDownList, "FORMDROPDOWN")
    );
    if !matches_instruction {
        return Err(OxmlError::InvalidValue(
            "w:ffData kind contradicts its field instruction".to_owned(),
        ));
    }
    if kind == LegacyFormFieldKind::CheckBox && !projection.checkbox_size_seen {
        return Err(OxmlError::MissingElement(
            "w:checkBox requires w:size or w:sizeAuto".to_owned(),
        ));
    }
    let value = match kind {
        LegacyFormFieldKind::TextInput => LegacyFormFieldValue::Text(
            projection
                .default_text
                .unwrap_or_else(|| cached_result.to_owned()),
        ),
        LegacyFormFieldKind::CheckBox => LegacyFormFieldValue::Checked(
            projection
                .checked
                .or(projection.check_default)
                .unwrap_or(false),
        ),
        LegacyFormFieldKind::DropDownList => {
            let index = projection
                .selected_index
                .or(projection.selected_default)
                .unwrap_or(0);
            if projection.choices.get(index).is_none() {
                return Err(OxmlError::InvalidValue(
                    "w:ddList selection is outside its entry list".to_owned(),
                ));
            }
            LegacyFormFieldValue::SelectedIndex(index)
        }
    };
    Ok(Some(LegacyFormFieldData {
        name: projection.name,
        enabled: projection.enabled.unwrap_or(true),
        calculate_on_exit: projection.calculate_on_exit.unwrap_or(true),
        kind,
        value,
        choices: projection.choices,
        max_length: projection.max_length,
    }))
}

fn project_legacy_form_element(
    element: &BytesStart<'_>,
    local: Option<&str>,
    depth: usize,
    ff_depth: usize,
    word_prefixes: &[String],
    kind_depth: &mut Option<usize>,
    projection: &mut LegacyFormProjection,
) -> Result<bool> {
    if let Some(local) = local
        && is_known_legacy_form_vocabulary(local)
    {
        let valid_position = if depth == ff_depth + 1 {
            legacy_form_ff_data_slot(Some(local)).is_some()
        } else if kind_depth.is_some_and(|kind| depth == kind + 1) {
            projection
                .kind
                .is_some_and(|kind| legacy_form_kind_child_slot(kind, local).is_some())
        } else {
            false
        };
        if !valid_position {
            return Err(OxmlError::InvalidValue(format!(
                "w:{local} is at the wrong legacy form container level"
            )));
        }
    }
    let is_leaf = (depth == ff_depth + 1
        && matches!(
            local,
            Some(
                "name"
                    | "enabled"
                    | "calcOnExit"
                    | "entryMacro"
                    | "exitMacro"
                    | "helpText"
                    | "statusText"
            )
        ))
        || (kind_depth.is_some_and(|kind| depth == kind + 1)
            && matches!(
                local,
                Some(
                    "type"
                        | "default"
                        | "maxLength"
                        | "format"
                        | "size"
                        | "sizeAuto"
                        | "checked"
                        | "result"
                        | "listEntry"
                )
            ));
    if depth == ff_depth + 1 {
        if let Some(slot) = legacy_form_ff_data_slot(local) {
            validate_legacy_form_child_order(&mut projection.ff_data_last_slot, slot, "w:ffData")?;
        }
        match local {
            Some("name") => {
                mark_legacy_form_singleton(&mut projection.name_seen, "w:name")?;
                projection.name = Some(required_bounded_legacy_form_value(
                    element,
                    word_prefixes,
                    "w:name",
                    20,
                )?);
            }
            Some("enabled") => {
                mark_legacy_form_singleton(&mut projection.enabled_seen, "w:enabled")?;
                projection.enabled = Some(legacy_form_boolean_value(
                    element,
                    word_prefixes,
                    "w:enabled",
                )?);
            }
            Some("calcOnExit") => {
                mark_legacy_form_singleton(&mut projection.calculate_on_exit_seen, "w:calcOnExit")?;
                projection.calculate_on_exit = Some(legacy_form_boolean_value(
                    element,
                    word_prefixes,
                    "w:calcOnExit",
                )?);
            }
            Some("entryMacro") => {
                mark_legacy_form_singleton(&mut projection.entry_macro_seen, "w:entryMacro")?;
                required_bounded_legacy_form_value(element, word_prefixes, "w:entryMacro", 33)?;
            }
            Some("exitMacro") => {
                mark_legacy_form_singleton(&mut projection.exit_macro_seen, "w:exitMacro")?;
                required_bounded_legacy_form_value(element, word_prefixes, "w:exitMacro", 33)?;
            }
            Some("helpText") => {
                mark_legacy_form_singleton(&mut projection.help_text_seen, "w:helpText")?;
                validate_legacy_info_text(element, word_prefixes, "w:helpText", 255)?;
            }
            Some("statusText") => {
                mark_legacy_form_singleton(&mut projection.status_text_seen, "w:statusText")?;
                validate_legacy_info_text(element, word_prefixes, "w:statusText", 140)?;
            }
            Some("textInput") => set_legacy_form_kind(
                projection,
                LegacyFormFieldKind::TextInput,
                depth,
                kind_depth,
            )?,
            Some("checkBox") => {
                set_legacy_form_kind(projection, LegacyFormFieldKind::CheckBox, depth, kind_depth)?
            }
            Some("ddList") => set_legacy_form_kind(
                projection,
                LegacyFormFieldKind::DropDownList,
                depth,
                kind_depth,
            )?,
            _ => {}
        }
    } else if kind_depth.is_some_and(|kind| depth == kind + 1) {
        if let (Some(kind), Some(local)) = (projection.kind, local) {
            if let Some(slot) = legacy_form_kind_child_slot(kind, local) {
                validate_legacy_form_child_order(
                    &mut projection.kind_last_slot,
                    slot,
                    "legacy form kind",
                )?;
            } else if is_known_legacy_form_kind_child(local) {
                return Err(OxmlError::InvalidValue(format!(
                    "w:{local} is not valid for this legacy form kind"
                )));
            }
        }
        match (projection.kind, local) {
            (Some(LegacyFormFieldKind::TextInput), Some("type")) => {
                mark_legacy_form_singleton(&mut projection.text_type_seen, "w:textInput/w:type")?;
                let value =
                    required_legacy_form_value(element, word_prefixes, "w:textInput/w:type")?;
                if !matches!(
                    value.as_str(),
                    "regular" | "number" | "date" | "currentTime" | "currentDate" | "calculated"
                ) {
                    return Err(OxmlError::InvalidValue(
                        "invalid w:textInput/w:type token".to_owned(),
                    ));
                }
            }
            (Some(LegacyFormFieldKind::TextInput), Some("default")) => {
                mark_legacy_form_singleton(
                    &mut projection.default_text_seen,
                    "w:textInput/w:default",
                )?;
                projection.default_text = Some(required_bounded_legacy_form_value(
                    element,
                    word_prefixes,
                    "w:textInput/w:default",
                    255,
                )?);
            }
            (Some(LegacyFormFieldKind::TextInput), Some("maxLength")) => {
                mark_legacy_form_singleton(
                    &mut projection.max_length_seen,
                    "w:textInput/w:maxLength",
                )?;
                let value =
                    legacy_form_usize_value(element, word_prefixes, "w:textInput/w:maxLength")?;
                if !(1..=32_767).contains(&value) {
                    return Err(OxmlError::InvalidValue(
                        "w:textInput/w:maxLength must be between 1 and 32767".to_owned(),
                    ));
                }
                projection.max_length = Some(value);
            }
            (Some(LegacyFormFieldKind::TextInput), Some("format")) => {
                mark_legacy_form_singleton(
                    &mut projection.text_format_seen,
                    "w:textInput/w:format",
                )?;
                required_bounded_legacy_form_value(
                    element,
                    word_prefixes,
                    "w:textInput/w:format",
                    64,
                )?;
            }
            (Some(LegacyFormFieldKind::CheckBox), Some("size")) => {
                mark_legacy_form_singleton(
                    &mut projection.checkbox_size_seen,
                    "w:checkBox size choice",
                )?;
                legacy_form_hps_measure_value(element, word_prefixes, "w:checkBox/w:size")?;
            }
            (Some(LegacyFormFieldKind::CheckBox), Some("sizeAuto")) => {
                mark_legacy_form_singleton(
                    &mut projection.checkbox_size_seen,
                    "w:checkBox size choice",
                )?;
                legacy_form_boolean_value(element, word_prefixes, "w:checkBox/w:sizeAuto")?;
            }
            (Some(LegacyFormFieldKind::CheckBox), Some("checked")) => {
                mark_legacy_form_singleton(&mut projection.checked_seen, "w:checkBox/w:checked")?;
                projection.checked = Some(legacy_form_boolean_value(
                    element,
                    word_prefixes,
                    "w:checkBox/w:checked",
                )?);
            }
            (Some(LegacyFormFieldKind::CheckBox), Some("default")) => {
                mark_legacy_form_singleton(
                    &mut projection.check_default_seen,
                    "w:checkBox/w:default",
                )?;
                projection.check_default = Some(legacy_form_boolean_value(
                    element,
                    word_prefixes,
                    "w:checkBox/w:default",
                )?);
            }
            (Some(LegacyFormFieldKind::DropDownList), Some("result")) => {
                mark_legacy_form_singleton(
                    &mut projection.selected_index_seen,
                    "w:ddList/w:result",
                )?;
                projection.selected_index = Some(legacy_form_usize_value(
                    element,
                    word_prefixes,
                    "w:ddList/w:result",
                )?);
            }
            (Some(LegacyFormFieldKind::DropDownList), Some("default")) => {
                mark_legacy_form_singleton(
                    &mut projection.selected_default_seen,
                    "w:ddList/w:default",
                )?;
                let value = legacy_form_usize_value(element, word_prefixes, "w:ddList/w:default")?;
                if value > 24 {
                    return Err(OxmlError::InvalidValue(
                        "w:ddList/w:default must be between 0 and 24".to_owned(),
                    ));
                }
                projection.selected_default = Some(value);
            }
            (Some(LegacyFormFieldKind::DropDownList), Some("listEntry")) => {
                if projection.choices.len() == 25 {
                    return Err(OxmlError::InvalidValue(
                        "w:ddList contains more than 25 list entries".to_owned(),
                    ));
                }
                projection.choices.push(required_bounded_legacy_form_value(
                    element,
                    word_prefixes,
                    "w:ddList/w:listEntry",
                    255,
                )?);
            }
            _ => {}
        }
    }
    Ok(is_leaf)
}

fn legacy_form_ff_data_slot(local: Option<&str>) -> Option<usize> {
    match local? {
        "name" => Some(0),
        "enabled" => Some(1),
        "calcOnExit" => Some(2),
        "entryMacro" => Some(3),
        "exitMacro" => Some(4),
        "helpText" => Some(5),
        "statusText" => Some(6),
        "checkBox" | "ddList" | "textInput" => Some(7),
        _ => None,
    }
}

fn legacy_form_kind_child_slot(kind: LegacyFormFieldKind, local: &str) -> Option<usize> {
    match kind {
        LegacyFormFieldKind::TextInput => match local {
            "type" => Some(0),
            "default" => Some(1),
            "maxLength" => Some(2),
            "format" => Some(3),
            _ => None,
        },
        LegacyFormFieldKind::CheckBox => match local {
            "size" | "sizeAuto" => Some(0),
            "default" => Some(1),
            "checked" => Some(2),
            _ => None,
        },
        LegacyFormFieldKind::DropDownList => match local {
            "result" => Some(0),
            "default" => Some(1),
            "listEntry" => Some(2),
            _ => None,
        },
    }
}

fn is_known_legacy_form_kind_child(local: &str) -> bool {
    matches!(
        local,
        "type"
            | "default"
            | "maxLength"
            | "format"
            | "size"
            | "sizeAuto"
            | "checked"
            | "result"
            | "listEntry"
    )
}

fn is_known_legacy_form_vocabulary(local: &str) -> bool {
    local == "ffData"
        || legacy_form_ff_data_slot(Some(local)).is_some()
        || is_known_legacy_form_kind_child(local)
}

fn is_legacy_form_element_only_container(
    stack_len: usize,
    ff_data_depth: Option<usize>,
    kind_depth: Option<usize>,
) -> bool {
    ff_data_depth.is_some_and(|depth| stack_len == depth + 1)
        || kind_depth.is_some_and(|depth| stack_len == depth + 1)
}

fn validate_legacy_form_child_order(
    last_slot: &mut Option<usize>,
    slot: usize,
    owner: &str,
) -> Result<()> {
    if last_slot.is_some_and(|last| slot < last) {
        return Err(OxmlError::InvalidValue(format!(
            "out-of-order modeled child in {owner}"
        )));
    }
    *last_slot = Some(slot);
    Ok(())
}

fn mark_legacy_form_singleton(seen: &mut bool, name: &str) -> Result<()> {
    if std::mem::replace(seen, true) {
        return Err(OxmlError::InvalidValue(format!(
            "duplicate legacy form singleton {name}"
        )));
    }
    Ok(())
}

fn required_legacy_form_value(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    name: &str,
) -> Result<String> {
    legacy_form_value_attribute(element, word_prefixes)?
        .ok_or_else(|| OxmlError::MissingElement(format!("{name} w:val attribute")))
}

fn required_bounded_legacy_form_value(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    name: &str,
    max_characters: usize,
) -> Result<String> {
    let value = required_legacy_form_value(element, word_prefixes, name)?;
    if value.chars().count() > max_characters {
        return Err(OxmlError::InvalidValue(format!(
            "{name} exceeds {max_characters} characters"
        )));
    }
    Ok(value)
}

fn legacy_form_boolean_value(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    name: &str,
) -> Result<bool> {
    let Some(value) = legacy_form_value_attribute(element, word_prefixes)? else {
        return Ok(true);
    };
    parse_field_bool(&value)
        .ok_or_else(|| OxmlError::InvalidValue(format!("invalid {name} boolean token")))
}

fn legacy_form_value_attribute(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
) -> Result<Option<String>> {
    legacy_form_word_attribute(element, word_prefixes, b"val", "w:val")
}

fn legacy_form_word_attribute(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    wanted: &[u8],
    name: &str,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let Some(separator) = key.iter().position(|byte| *byte == b':') else {
            continue;
        };
        if key.get(separator + 1..) == Some(wanted)
            && word_prefixes
                .iter()
                .any(|prefix| prefix.as_bytes() == &key[..separator])
        {
            if value.is_some() {
                return Err(OxmlError::InvalidValue(format!(
                    "duplicate legacy form {name} attribute"
                )));
            }
            let raw = std::str::from_utf8(&attribute.value)?;
            value = Some(
                quick_xml::escape::unescape(raw)
                    .map_err(|error| {
                        OxmlError::InvalidValue(format!(
                            "invalid legacy form {name} attribute: {error}"
                        ))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn validate_legacy_info_text(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    name: &str,
    max_characters: usize,
) -> Result<()> {
    if legacy_form_word_attribute(element, word_prefixes, b"type", "w:type")?
        .is_some_and(|value| !matches!(value.as_str(), "text" | "autoText"))
    {
        return Err(OxmlError::InvalidValue(format!(
            "invalid {name} w:type token"
        )));
    }
    if legacy_form_word_attribute(element, word_prefixes, b"val", "w:val")?
        .is_some_and(|value| value.chars().count() > max_characters)
    {
        return Err(OxmlError::InvalidValue(format!(
            "{name} w:val exceeds {max_characters} characters"
        )));
    }
    Ok(())
}

fn legacy_form_usize_value(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    name: &str,
) -> Result<usize> {
    required_legacy_form_value(element, word_prefixes, name)?
        .parse()
        .map_err(|_| OxmlError::InvalidValue(format!("invalid {name} numeric token")))
}

fn legacy_form_hps_measure_value(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    name: &str,
) -> Result<()> {
    let value = required_legacy_form_value(element, word_prefixes, name)?;
    if value.parse::<u64>().is_ok() {
        return Ok(());
    }
    let Some(unit) = ["mm", "cm", "in", "pt", "pc", "pi"]
        .into_iter()
        .find(|unit| value.ends_with(unit))
    else {
        return Err(OxmlError::InvalidValue(format!(
            "invalid {name} measure token"
        )));
    };
    let number = &value[..value.len() - unit.len()];
    let valid_number = number.split_once('.').map_or_else(
        || !number.is_empty() && number.bytes().all(|byte| byte.is_ascii_digit()),
        |(integer, fraction)| {
            !integer.is_empty()
                && !fraction.is_empty()
                && integer.bytes().all(|byte| byte.is_ascii_digit())
                && fraction.bytes().all(|byte| byte.is_ascii_digit())
        },
    );
    if !valid_number {
        return Err(OxmlError::InvalidValue(format!(
            "invalid {name} measure token"
        )));
    }
    Ok(())
}

fn set_legacy_form_kind(
    projection: &mut LegacyFormProjection,
    kind: LegacyFormFieldKind,
    depth: usize,
    kind_depth: &mut Option<usize>,
) -> Result<()> {
    if projection.kind.replace(kind).is_some() {
        return Err(OxmlError::InvalidValue(
            "w:ffData contains more than one form kind".to_owned(),
        ));
    }
    *kind_depth = Some(depth);
    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
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
    cached_fields: Vec<(std::ops::Range<usize>, Field)>,
    valid: bool,
}

struct ComplexFieldProjection<'a> {
    runs: &'a mut Vec<CT_R>,
    run_sources: &'a mut Vec<Option<Vec<u8>>>,
    extra_xml: &'a mut Vec<(usize, Vec<u8>)>,
    comment_ranges: &'a mut Vec<CommentRangeMarker>,
    comment_sources: &'a [(CommentRangeMarker, Vec<u8>)],
    bookmark_markers: &'a mut [BookmarkMarker],
    content_controls: &'a mut [(usize, usize, usize, CT_Sdt)],
    revisions: &'a mut [(usize, usize, CT_Revision)],
    hyperlinks: &'a mut [HyperlinkSpan],
    rubies: &'a mut [CT_Ruby],
    word_prefixes: &'a [String],
}

fn project_complex_fields(projection: ComplexFieldProjection<'_>) -> Result<()> {
    let ComplexFieldProjection {
        runs,
        run_sources,
        extra_xml,
        comment_ranges,
        comment_sources,
        bookmark_markers,
        content_controls,
        revisions,
        hyperlinks,
        rubies,
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
        if let Some(child) = runs[run_index]
            .content
            .iter()
            .find_map(|content| match content {
                RunContent::Field(field) if field.form() == FieldForm::Simple => Some(field),
                _ => None,
            })
        {
            for parent in &mut stack {
                if parent.separate_run.is_some() {
                    parent.cached_result.push_str(&child.cached_result);
                    for (text, properties) in child.cached_display_segments() {
                        push_cached_display_segment(
                            &mut parent.cached_segments,
                            run_index,
                            text,
                            properties,
                        );
                    }
                }
            }
            if let Some(parent) = stack.last_mut()
                && parent.separate_run.is_some()
            {
                let end = parent.cached_result.len();
                parent
                    .cached_fields
                    .push((end - child.cached_result.len()..end, child.clone()));
            }
            continue;
        }

        for event in complex_field_events(raw, word_prefixes)? {
            match event {
                ComplexFieldEvent::Begin(dirty) => stack.push(ComplexFieldBuilder {
                    start_run: run_index,
                    separate_run: None,
                    dirty,
                    instruction: Vec::new(),
                    cached_result: String::new(),
                    cached_segments: Vec::new(),
                    cached_fields: Vec::new(),
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
                        comment_sources,
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
                    parsed.cached_fields = field.cached_fields;
                    if field.separate_run.is_some()
                        && let FieldSource::Parsed { raw_xml, .. } = &parsed.source
                        && source_has_note_reference_name(raw_xml)
                        && let Ok((cache_runs, cache_ranges)) =
                            parsed_complex_cache_runs(raw_xml, word_prefixes)
                        && cache_runs
                            .iter()
                            .flat_map(|run| &run.content)
                            .any(|content| {
                                matches!(
                                    content,
                                    RunContent::FootnoteRef { .. }
                                        | RunContent::EndnoteRef { .. }
                                        | RunContent::CommentReference { .. }
                                )
                            })
                    {
                        parsed.parsed_cached_runs = Some(cache_runs);
                        parsed.parsed_cached_comment_ranges = cache_ranges;
                    }
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
                        } else {
                            let end = parent.cached_result.len();
                            let start = end.saturating_sub(parsed.cached_result.len());
                            parent.cached_fields.push((start..end, parsed));
                        }
                    } else if valid {
                        completed.push((field.start_run, run_index, parsed, None));
                    }
                }
            }
        }
    }

    let mut grouped = Vec::<(usize, usize, Vec<(Field, Option<CT_RPr>)>)>::new();
    for (start, end, field, properties) in completed {
        // Fields that share a physical run form one span, so the splice of
        // one never removes the run that holds the next.
        if let Some((_, group_end, fields)) = grouped.last_mut()
            && start <= *group_end
        {
            *group_end = end.max(*group_end);
            fields.push((field, properties));
        } else {
            grouped.push((start, end, vec![(field, properties)]));
        }
    }
    for (start, end, mut fields) in grouped.into_iter().rev() {
        let mut owned_comment_ids = std::collections::BTreeSet::new();
        let mut cache_owners = fields.iter().map(|(field, _)| field).collect::<Vec<_>>();
        while let Some(field) = cache_owners.pop() {
            cache_owners.extend(field.all_nested_fields_in_source_order());
            let mut ranges = field.parsed_cached_comment_ranges.clone();
            for marker in &mut ranges {
                match marker {
                    CommentRangeMarker::Start { raw_before, .. }
                    | CommentRangeMarker::End { raw_before, .. } => *raw_before = 0,
                }
            }
            if !ranges.is_empty()
                && field
                    .cached_result_runs()
                    .is_some_and(|runs| validate_cached_comment_ranges(runs, &ranges).is_ok())
            {
                for marker in ranges {
                    match marker {
                        CommentRangeMarker::Start { id, .. }
                        | CommentRangeMarker::End { id, .. } => {
                            owned_comment_ids.insert(id);
                        }
                    }
                }
            }
        }
        let unrelated_comments = comment_ranges
            .iter()
            .filter(|marker| match marker {
                CommentRangeMarker::Start { id, .. } | CommentRangeMarker::End { id, .. } => {
                    !owned_comment_ids.contains(id)
                }
            })
            .cloned()
            .collect::<Vec<_>>();
        if has_typed_boundary_inside(
            start,
            end,
            &unrelated_comments,
            bookmark_markers,
            content_controls,
            revisions,
            hyperlinks,
        ) {
            continue;
        }
        let raw_xml = complex_field_source(
            start,
            end,
            run_sources,
            extra_xml,
            hyperlinks,
            comment_sources,
        );
        let owner_id = fields.iter().find_map(|(field, _)| match field.source {
            FieldSource::Parsed { source_id, .. } => Some(source_id),
            FieldSource::New { .. } => None,
        });
        for (field, _) in &mut fields {
            if let FieldSource::Parsed {
                raw_xml: source,
                owner_id: field_owner_id,
                ..
            } = &mut field.source
            {
                *source = raw_xml.clone();
                if let Some(owner_id) = owner_id {
                    *field_owner_id = owner_id;
                }
            }
        }
        comment_ranges.retain(|marker| match marker {
            CommentRangeMarker::Start { id, .. } | CommentRangeMarker::End { id, .. } => {
                !owned_comment_ids.contains(id)
                    || !(marker.run_index() > start && marker.run_index() <= end)
            }
        });
        extra_xml.retain(|(at, _)| !(*at > start && *at <= end));
        for hyperlink in hyperlinks.iter_mut() {
            let hyperlink_start = hyperlink.run_start;
            hyperlink.extra_xml.retain(|(boundary, _, _)| {
                let absolute = hyperlink_start + *boundary;
                !(absolute > start && absolute <= end)
            });
        }

        let replacement = field_span_runs(&raw_xml, fields, word_prefixes);
        let replacement_count = replacement.len();
        runs.splice(start..=end, replacement);
        run_sources.splice(start..=end, std::iter::repeat_n(None, replacement_count));
        remap_complex_field_boundaries(
            start,
            end,
            replacement_count,
            ComplexFieldBoundariesMut {
                extra_xml,
                comment_ranges,
                bookmark_markers,
                content_controls,
                revisions,
                hyperlinks,
                rubies,
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
    comment_sources: &[(CommentRangeMarker, Vec<u8>)],
) -> Vec<u8> {
    let mut source = Vec::new();
    for (run_index, raw_source) in run_sources.iter().enumerate().take(end + 1).skip(start) {
        if let Some(raw) = raw_source.as_deref() {
            source.extend_from_slice(raw);
        }
        if run_index < end {
            let extras = extra_xml
                .iter()
                .filter(|(at, _)| *at == run_index + 1)
                .collect::<Vec<_>>();
            for slot in 0..=extras.len() {
                for (marker, raw) in comment_sources.iter().filter(|(marker, _)| {
                    marker.run_index() == run_index + 1 && marker.raw_before() == slot
                }) {
                    let _ = marker;
                    source.extend_from_slice(raw);
                }
                if let Some((_, raw)) = extras.get(slot) {
                    source.extend_from_slice(raw);
                }
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

/// One part of a physical field span: a whole field with its bytes, or the
/// run read from content outside any field.
enum FieldSpanPart {
    Field(Vec<u8>),
    Outside(CT_R),
}

/// Split the exact bytes of a complex field span into its fields and the run
/// content outside them, in order.
///
/// Each part is written as its own runs, so a part taken from a shared run
/// repeats that run's start tag and `w:rPr`. Returns `None` for a shape this
/// does not model, which leaves the span as it was read.
fn field_span_parts(raw: &[u8], word_prefixes: &[String]) -> Result<Option<Vec<FieldSpanPart>>> {
    use std::ops::Range;

    // A run's start tag, `w:rPr` and end tag.
    let mut runs = Vec::<(Range<usize>, Option<Range<usize>>, Range<usize>)>::new();
    // Each part: whether it is a field, and its children (run index, bytes).
    let mut parts = Vec::<(bool, Vec<(Option<usize>, Range<usize>)>)>::new();
    let mut current: Option<usize> = None;
    let mut depth = 0usize;
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        let (run_start, run_prefixes) = match event {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                if is_word_element(start.name().as_ref(), b"r", &prefixes) {
                    (before..reader.buffer_position() as usize, prefixes)
                } else {
                    reader.read_to_end_into(start.name(), &mut Vec::new())?;
                    let range = before..reader.buffer_position() as usize;
                    let Some(part) = current.filter(|_| depth > 0) else {
                        return Ok(None);
                    };
                    parts[part].1.push((None, range));
                    buffer.clear();
                    continue;
                }
            }
            Event::Eof => break,
            _ => {
                let range = before..reader.buffer_position() as usize;
                let Some(part) = current.filter(|_| depth > 0) else {
                    return Ok(None);
                };
                parts[part].1.push((None, range));
                buffer.clear();
                continue;
            }
        };
        let run = runs.len();
        runs.push((run_start, None, 0..0));
        let mut children = 0usize;
        loop {
            buffer.clear();
            let child_start = reader.buffer_position() as usize;
            let kind = match reader.read_event_into(&mut buffer)? {
                Event::Start(child) => {
                    let prefixes = word_prefixes_at(&child, &run_prefixes)?;
                    let kind = field_span_child_kind(&child, &prefixes);
                    reader.read_to_end_into(child.name(), &mut Vec::new())?;
                    kind
                }
                Event::Empty(child) => {
                    let prefixes = word_prefixes_at(&child, &run_prefixes)?;
                    field_span_child_kind(&child, &prefixes)
                }
                Event::End(_) => {
                    let end = reader.buffer_position() as usize;
                    if children == 0 {
                        let Some(part) = current.filter(|_| depth > 0) else {
                            return Ok(None);
                        };
                        parts[part].1.push((Some(run), child_start..child_start));
                    }
                    runs[run].2 = child_start..end;
                    break;
                }
                Event::Eof => return Ok(None),
                _ => FieldSpanChild::Other,
            };
            let range = child_start..reader.buffer_position() as usize;
            if kind == FieldSpanChild::Properties {
                runs[run].1 = Some(range);
                continue;
            }
            children += 1;
            let inside = match kind {
                FieldSpanChild::Begin => {
                    if depth == 0 {
                        parts.push((true, Vec::new()));
                        current = Some(parts.len() - 1);
                    }
                    depth += 1;
                    true
                }
                FieldSpanChild::End => {
                    if depth == 0 {
                        return Ok(None);
                    }
                    depth -= 1;
                    true
                }
                FieldSpanChild::FieldMarker => {
                    if depth == 0 {
                        return Ok(None);
                    }
                    true
                }
                FieldSpanChild::Other | FieldSpanChild::Properties => depth > 0,
            };
            if !inside && current.is_none_or(|part| parts[part].0) {
                parts.push((false, Vec::new()));
                current = Some(parts.len() - 1);
            }
            let part = current.expect("a part is open");
            parts[part].1.push((Some(run), range));
            if kind == FieldSpanChild::End && depth == 0 {
                current = None;
            }
        }
        buffer.clear();
    }
    if depth != 0 {
        return Ok(None);
    }

    let materialize = |children: &[(Option<usize>, Range<usize>)]| {
        let mut bytes = Vec::new();
        let mut open = None;
        for (run, range) in children {
            if *run != open {
                if let Some(open) = open {
                    bytes.extend_from_slice(&raw[runs[open].2.clone()]);
                }
                if let Some(run) = *run {
                    bytes.extend_from_slice(&raw[runs[run].0.clone()]);
                    if let Some(properties) = runs[run].1.clone() {
                        bytes.extend_from_slice(&raw[properties]);
                    }
                }
                open = *run;
            }
            bytes.extend_from_slice(&raw[range.clone()]);
        }
        if let Some(open) = open {
            bytes.extend_from_slice(&raw[runs[open].2.clone()]);
        }
        bytes
    };

    // Content outside every field that no reader sees, such as whitespace or
    // an unknown element, stays with its neighbouring field as before.
    let mut outside = Vec::with_capacity(parts.len());
    let mut index = 0;
    while index < parts.len() {
        if parts[index].0 {
            outside.push(None);
            index += 1;
            continue;
        }
        if let Ok(run) = parse_run_raw(&materialize(&parts[index].1), word_prefixes)
            && !run.content.is_empty()
        {
            outside.push(Some(run));
            index += 1;
            continue;
        }
        let (_, children) = parts.remove(index);
        if let Some(next) = parts.get_mut(index) {
            next.1.splice(0..0, children);
        } else if let Some(previous) = index.checked_sub(1).and_then(|at| parts.get_mut(at)) {
            previous.1.extend(children);
        }
    }
    Ok(Some(
        parts
            .iter()
            .zip(outside)
            .map(|((_, children), run)| match run {
                Some(run) => FieldSpanPart::Outside(run),
                None => FieldSpanPart::Field(materialize(children)),
            })
            .collect(),
    ))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FieldSpanChild {
    Properties,
    Begin,
    End,
    /// A separator or instruction text, which only ever sits inside a field.
    FieldMarker,
    Other,
}

fn field_span_child_kind(element: &BytesStart<'_>, word_prefixes: &[String]) -> FieldSpanChild {
    let name = element.name();
    if is_word_element(name.as_ref(), b"rPr", word_prefixes) {
        return FieldSpanChild::Properties;
    }
    if is_word_element(name.as_ref(), b"instrText", word_prefixes) {
        return FieldSpanChild::FieldMarker;
    }
    if !is_word_element(name.as_ref(), b"fldChar", word_prefixes) {
        return FieldSpanChild::Other;
    }
    match optional_word_attribute(element, b"fldCharType", word_prefixes).as_deref() {
        Some("begin") => FieldSpanChild::Begin,
        Some("end") => FieldSpanChild::End,
        _ => FieldSpanChild::FieldMarker,
    }
}

/// Whether one field's span starts with its begin marker and ends with its
/// end marker, the shape Word writes, so no run content sits outside it.
///
/// Only the first and last run children are read. Any other shape answers
/// `false` and takes the full split.
fn field_span_is_bare(raw: &[u8], word_prefixes: &[String]) -> bool {
    let marker = |element: &BytesStart<'_>, prefixes: &[String], kind: &str| {
        is_word_element(element.name().as_ref(), b"fldChar", prefixes)
            && optional_word_attribute(element, b"fldCharType", prefixes).as_deref() == Some(kind)
    };
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut run_prefixes = None;
    let begins = loop {
        let Ok(event) = reader.read_event_into(&mut buffer) else {
            return false;
        };
        match event {
            Event::Start(element) => {
                let Ok(prefixes) = word_prefixes_at(&element, word_prefixes) else {
                    return false;
                };
                if run_prefixes.is_none() {
                    if !is_word_element(element.name().as_ref(), b"r", &prefixes) {
                        return false;
                    }
                    run_prefixes = Some(prefixes);
                } else if is_word_element(element.name().as_ref(), b"rPr", &prefixes) {
                    if reader
                        .read_to_end_into(element.name(), &mut Vec::new())
                        .is_err()
                    {
                        return false;
                    }
                } else {
                    return false;
                }
            }
            Event::Empty(element) => {
                let Some(inherited) = run_prefixes.as_deref() else {
                    return false;
                };
                let Ok(prefixes) = word_prefixes_at(&element, inherited) else {
                    return false;
                };
                break marker(&element, &prefixes, "begin");
            }
            _ => return false,
        }
        buffer.clear();
    };
    if !begins {
        return false;
    }
    // The last run child is the empty element before the closing run tag.
    let body = raw.trim_ascii_end();
    let Some(close) = body.iter().rposition(|byte| *byte == b'<') else {
        return false;
    };
    let Some(last) = body[..close].iter().rposition(|byte| *byte == b'<') else {
        return false;
    };
    let mut reader = Reader::from_reader(&body[last..close]);
    let mut buffer = Vec::new();
    let Some(prefixes) = run_prefixes else {
        return false;
    };
    matches!(
        reader.read_event_into(&mut buffer),
        Ok(Event::Empty(element))
            if word_prefixes_at(&element, &prefixes)
                .is_ok_and(|prefixes| marker(&element, &prefixes, "end"))
    ) && body[last..close].trim_ascii_end().ends_with(b"/>")
}

/// Turn a projected field span into model runs, splitting out the run
/// content that sits outside its fields.
///
/// Without such content the span stays one run per field, as before.
fn field_span_runs(
    raw: &[u8],
    fields: Vec<(Field, Option<CT_RPr>)>,
    word_prefixes: &[String],
) -> Vec<CT_R> {
    let mut fields = fields;
    if source_has_note_reference_name(raw)
        && let Ok(Some(parts)) = field_span_parts(raw, word_prefixes)
    {
        for ((field, _), source) in
            fields
                .iter_mut()
                .zip(parts.into_iter().filter_map(|part| match part {
                    FieldSpanPart::Field(source) => Some(source),
                    FieldSpanPart::Outside(_) => None,
                }))
        {
            if let Ok((runs, ranges)) = parsed_complex_cache_runs(&source, word_prefixes)
                && runs.iter().flat_map(|run| &run.content).any(|content| {
                    matches!(
                        content,
                        RunContent::FootnoteRef { .. }
                            | RunContent::EndnoteRef { .. }
                            | RunContent::CommentReference { .. }
                    )
                })
            {
                field.parsed_cached_runs = Some(runs);
                field.parsed_cached_comment_ranges = ranges;
            }
        }
    }
    let one_run_per_field = |fields: Vec<(Field, Option<CT_RPr>)>| {
        fields
            .into_iter()
            .map(|(field, properties)| field_run(field, properties))
            .collect()
    };
    if fields.len() == 1 && field_span_is_bare(raw, word_prefixes) {
        return one_run_per_field(fields);
    }
    let parts = match field_span_parts(raw, word_prefixes) {
        Ok(Some(parts))
            if (fields.len() > 1
                || parts
                    .iter()
                    .any(|part| matches!(part, FieldSpanPart::Outside(_))))
                && parts
                    .iter()
                    .filter(|part| matches!(part, FieldSpanPart::Field(_)))
                    .count()
                    == fields.len() =>
        {
            parts
        }
        _ => return one_run_per_field(fields),
    };
    let shape = parts
        .into_iter()
        .map(|part| match part {
            FieldSpanPart::Field(_) => None,
            FieldSpanPart::Outside(run) => Some(run),
        })
        .collect::<Vec<_>>();
    let mut fields = fields.into_iter().enumerate();
    shape
        .iter()
        .enumerate()
        .map(|(position, run)| match run {
            Some(run) => run.clone(),
            None => {
                let (field_index, (mut field, properties)) =
                    fields.next().expect("one field per field part");
                field.span = Some(Box::new(FieldRunSpan {
                    runs: shape.clone(),
                    position,
                    field_index,
                }));
                field_run(field, properties)
            }
        })
        .collect()
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
    rubies: &'a mut [CT_Ruby],
}

fn remap_complex_field_boundaries(
    start: usize,
    end: usize,
    replacement_count: usize,
    boundaries: ComplexFieldBoundariesMut<'_>,
) {
    let ComplexFieldBoundariesMut {
        extra_xml,
        comment_ranges,
        bookmark_markers,
        content_controls,
        revisions,
        hyperlinks,
        rubies,
    } = boundaries;
    let source_count = end - start + 1;
    let remap = |at: &mut usize| {
        if *at > end {
            if replacement_count >= source_count {
                *at += replacement_count - source_count;
            } else {
                *at -= source_count - replacement_count;
            }
        } else if *at > start {
            *at = start + replacement_count;
        }
    };
    for (at, _) in extra_xml {
        if *at != ROOT_ATTRIBUTES_POSITION {
            remap(at);
        }
    }
    for marker in comment_ranges {
        match marker {
            CommentRangeMarker::Start { run_index, .. }
            | CommentRangeMarker::End { run_index, .. } => remap(run_index),
        }
    }
    for marker in bookmark_markers {
        if marker.run_index > end {
            if replacement_count >= source_count {
                let inserted = replacement_count - source_count;
                marker.projected_run_index += inserted;
                marker.tracked_run_index += inserted;
            } else {
                let removed = source_count - replacement_count;
                marker.projected_run_index = marker.projected_run_index.saturating_sub(removed);
                marker.tracked_run_index = marker.tracked_run_index.saturating_sub(removed);
            }
        }
        remap(&mut marker.run_index);
    }
    for (at, _, _, _) in content_controls {
        remap(at);
    }
    for (at, _, _) in revisions {
        remap(at);
    }
    for ruby in rubies {
        remap(&mut ruby.base_start);
        remap(&mut ruby.base_end);
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

/// Why [`CT_P::split_run`] could not split a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum RunSplitError {
    #[error("run index {run_index} is out of range for a paragraph with {run_count} runs")]
    RunOutOfRange { run_index: usize, run_count: usize },
    #[error("split offset {offset} exceeds a run's literal text length of {len} characters")]
    OffsetOutOfRange { offset: usize, len: usize },
    #[error("bookmark markers could not be projected after the split")]
    BookmarkProjection,
}

/// The range kind that [`CT_P::anchor_accepted_range`] writes.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RangeAnchor<'a> {
    /// Comment range markers, with the reference run right after the end.
    Comment(i32),
    /// Bookmark markers.
    Bookmark { id: i32, name: &'a str },
    /// Permission range with one selected editor or editor group.
    Permission {
        id: i32,
        editor: Option<&'a str>,
        group: Option<&'a str>,
    },
    /// Spelling or grammar proofing range, which has no numeric identifier.
    Proofing { kind: &'a str },
}

/// Why [`CT_P::anchor_accepted_range`] could not place a range exactly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum RangeAnchorError {
    #[error("run boundary {boundary} exceeds the paragraph run count {run_count}")]
    OutOfRange { boundary: usize, run_count: usize },
    #[error("run range {start}..{end} ends before it starts")]
    Reversed { start: usize, end: usize },
    #[error("run range {start}..{end} crosses the edge of an inline content control")]
    CrossesControl { start: usize, end: usize },
    #[error(
        "run boundary {boundary} falls inside an inline content control, where a range that continues into another paragraph cannot start or end"
    )]
    InsideControl { boundary: usize },
    #[error("run boundary {boundary} falls inside a tracked insertion or move")]
    InsideRevision { boundary: usize },
    #[error("run boundary {boundary} sits next to a tracked change inside a hyperlink")]
    HyperlinkRevision { boundary: usize },
    #[error("range markers could not be written into the paragraph")]
    Write,
}

/// One physical place for a range marker.
#[derive(Debug, Clone, PartialEq, Eq)]
enum MarkerSite {
    /// A position among the paragraph-level children at a direct run boundary.
    Paragraph { boundary: usize, position: usize },
    /// A `w:sdtContent` child index in the inline control that `controls`
    /// reaches: a paragraph control index, then nested content indexes.
    Control { controls: Vec<usize>, index: usize },
}

/// A direct run boundary and a position among its paragraph-level children.
type ChildPosition = (usize, usize);

/// One paragraph-level child at a direct run boundary, in serialization order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BoundaryItem {
    Control(usize),
    Marker(usize),
    Raw(usize),
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
    /// Ruby phonetic guides over half-open spans of [`Self::runs`].
    ///
    /// The base runs are ordinary paragraph runs, so text extraction, search
    /// and redaction see them without knowing about ruby. The phonetic runs
    /// stay inside the annotation, which is what keeps them out of every
    /// text projection.
    pub rubies: Vec<CT_Ruby>,
}

#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum AcceptedOwnerOrder {
    BeforeRaw,
    Raw(usize),
    AfterRaw,
}

/// One recursive step to a visible accepted-view run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AcceptedRunPathSegment {
    /// A direct run index in the current paragraph or content control.
    Run(usize),
    /// A content-control index in the current paragraph or content control.
    ContentControl(usize),
    /// A revision index in the current paragraph or content control.
    Revision(usize),
}

/// A checked recursive address for one visible accepted-view run.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct AcceptedRunPath {
    pub(crate) segments: Vec<AcceptedRunPathSegment>,
}

impl AcceptedRunPath {
    /// Return the recursive source steps for this run.
    pub fn segments(&self) -> &[AcceptedRunPathSegment] {
        &self.segments
    }
}

/// A content control or a revision wrapper at a run boundary of a paragraph.
#[derive(Clone, Copy)]
pub(crate) enum BoundaryOwner {
    /// An index into [`CT_P::content_controls`].
    ContentControl(usize),
    /// An index into [`CT_P::revisions`].
    Revision(usize),
    /// An index into [`CT_P::extra_xml`] of a smart tag or an inline custom
    /// XML element, see [`run_wrapper_paragraph`].
    Wrapper(usize),
}

/// The content controls, revision wrappers, smart tags and inline custom
/// XML elements at run `boundary` of `paragraph`, in document order.
pub(crate) fn boundary_owners(paragraph: &CT_P, boundary: usize) -> Vec<BoundaryOwner> {
    let mut owners = paragraph
        .content_controls
        .iter()
        .enumerate()
        .filter(|(_, (at, _, _, _))| *at == boundary)
        .map(|(index, (_, raw_before, _, _))| {
            (
                AcceptedOwnerOrder::Raw(*raw_before),
                BoundaryOwner::ContentControl(index),
            )
        })
        .chain(
            paragraph
                .revisions
                .iter()
                .enumerate()
                .filter(|(_, (at, _, _))| *at == boundary)
                .map(|(index, (_, slot, _))| {
                    let order = if let Some(hyperlink_index) = hyperlink_revision_index(*slot) {
                        if let Some(raw_before) = paragraph
                            .hyperlinks
                            .get(hyperlink_index)
                            .and_then(|hyperlink| hyperlink.preserved_raw_before)
                        {
                            AcceptedOwnerOrder::Raw(raw_before)
                        } else if paragraph
                            .hyperlinks
                            .get(hyperlink_index)
                            .is_some_and(|hyperlink| boundary == hyperlink.run_end)
                        {
                            AcceptedOwnerOrder::BeforeRaw
                        } else {
                            AcceptedOwnerOrder::AfterRaw
                        }
                    } else {
                        AcceptedOwnerOrder::Raw(*slot)
                    };
                    (order, BoundaryOwner::Revision(index))
                }),
        )
        .chain(
            paragraph
                .extra_xml
                .iter()
                .enumerate()
                .filter(|(_, (at, _))| *at == boundary)
                .enumerate()
                .filter(|(_, (_, (_, raw)))| is_run_wrapper(raw))
                .map(|(slot, (index, _))| {
                    (AcceptedOwnerOrder::Raw(slot), BoundaryOwner::Wrapper(index))
                }),
        )
        .collect::<Vec<_>>();
    // A control goes before the raw child of the same slot.
    owners
        .sort_by_key(|(order, owner)| (*order, !matches!(owner, BoundaryOwner::ContentControl(_))));
    owners.into_iter().map(|(_, owner)| owner).collect()
}

/// Whether `raw` is a smart tag or an inline custom XML element with
/// content, read with the conventional `w` prefix and those it declares.
fn is_run_wrapper(raw: &[u8]) -> bool {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let Ok(Event::Start(start)) = reader.read_event_into(&mut buffer) else {
        return false;
    };
    let Ok(prefixes) = word_prefixes_at(&start, &["w".to_owned()]) else {
        return false;
    };
    is_word_element(start.name().as_ref(), b"smartTag", &prefixes)
        || is_word_element(start.name().as_ref(), b"customXml", &prefixes)
}

/// The content of a smart tag, an inline custom XML element or a simple
/// field, parsed as a paragraph from its preserved source `raw` in the scope
/// of the Word prefixes `inherited`. Its runs are the text a reader sees
/// inside the wrapper. None for any other element, and for a wrapper whose
/// content [`with_run_wrapper_content`] could not write back, because `w`
/// does not name WordprocessingML there.
pub(crate) fn run_wrapper_paragraph(raw: &[u8], inherited: &[String]) -> Option<CT_P> {
    wrapper_content_paragraph(raw, inherited, &["smartTag", "customXml", "fldSimple"])
}

/// The content of `raw`, a Word element named by one of `wrappers`, parsed
/// as a paragraph, see [`run_wrapper_paragraph`].
fn wrapper_content_paragraph(raw: &[u8], inherited: &[String], wrappers: &[&str]) -> Option<CT_P> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let Ok(Event::Start(start)) = reader.read_event_into(&mut buffer) else {
        return None;
    };
    let prefixes = word_prefixes_at(&start, inherited).ok()?;
    if !wrappers
        .iter()
        .any(|local| is_word_element(start.name().as_ref(), local.as_bytes(), &prefixes))
        || !prefixes.iter().any(|prefix| prefix == "w")
    {
        return None;
    }
    let (content_start, content_end) = run_wrapper_content_bounds(raw)?;
    let mut paragraph_xml = b"<w:p>".to_vec();
    paragraph_xml.extend_from_slice(&raw[content_start..content_end]);
    paragraph_xml.extend_from_slice(b"</w:p>");
    let mut paragraph_reader = Reader::from_reader(paragraph_xml.as_slice());
    let mut paragraph_buffer = Vec::new();
    let Ok(Event::Start(paragraph_start)) = paragraph_reader.read_event_into(&mut paragraph_buffer)
    else {
        return None;
    };
    CT_P::from_xml_with_prefixes_and_root(&mut paragraph_reader, &prefixes, Some(&paragraph_start))
        .ok()
}

/// The source `raw` of a wrapper that [`run_wrapper_paragraph`] read, with
/// `paragraph` written back as its content between its own start and end
/// tags.
pub(crate) fn with_run_wrapper_content(raw: &[u8], paragraph: &CT_P) -> Result<Vec<u8>> {
    let missing = || OxmlError::MissingElement("wrapper content".to_owned());
    let mut writer = Writer::new(Vec::new());
    paragraph.to_xml(&mut writer)?;
    let paragraph_xml = writer.into_inner();
    // A paragraph without any content is written as `<w:p/>`.
    let content = match paragraph_xml.iter().position(|byte| *byte == b'>') {
        Some(end) if paragraph_xml[..end].ends_with(b"/") => &[][..],
        Some(end) => paragraph_xml[end + 1..]
            .strip_suffix(b"</w:p>")
            .ok_or_else(missing)?,
        None => return Err(missing()),
    };
    let (content_start, content_end) = run_wrapper_content_bounds(raw).ok_or_else(missing)?;
    let mut updated = raw[..content_start].to_vec();
    updated.extend_from_slice(content);
    updated.extend_from_slice(&raw[content_end..]);
    Ok(updated)
}

/// The byte range of the content of the element `raw`, between the end of
/// its start tag and the start of its end tag.
fn run_wrapper_content_bounds(raw: &[u8]) -> Option<(usize, usize)> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let Ok(Event::Start(_)) = reader.read_event_into(&mut buffer) else {
        return None;
    };
    let content_start = reader.buffer_position() as usize;
    let content_end = raw.iter().rposition(|byte| *byte == b'<')?;
    (content_start <= content_end).then_some((content_start, content_end))
}

/// The preserved source of a run that holds only an unchanged simple field,
/// with the Word prefixes it was read with.
pub(crate) fn simple_field_source(run: &CT_R) -> Option<(&[u8], &[String])> {
    let [RunContent::Field(field)] = run.content.as_slice() else {
        return None;
    };
    match &field.source {
        FieldSource::Parsed {
            form: FieldForm::Simple,
            raw_xml,
            word_prefixes,
            ..
        } if field.is_unchanged() => Some((raw_xml, word_prefixes)),
        _ => None,
    }
}

/// Read the simple field of `run` again from the source `raw`, which
/// [`with_run_wrapper_content`] wrote. Returns false, and changes nothing,
/// when `run` holds no unchanged simple field or `raw` is not one.
pub(crate) fn set_simple_field_source(run: &mut CT_R, raw: &[u8]) -> Result<bool> {
    let Some((_, prefixes)) = simple_field_source(run) else {
        return Ok(false);
    };
    let Some(field) = parse_simple_field(raw, prefixes)? else {
        return Ok(false);
    };
    run.content = vec![RunContent::Field(field)];
    Ok(true)
}

/// Append the text of the accepted view of `paragraph` to `output`, see
/// [`CT_P::accepted_text`].
fn append_accepted_text(paragraph: &CT_P, output: &mut String) {
    let wrapper_text = |raw: &[u8], prefixes: &[String], output: &mut String| {
        if let Some(content) = run_wrapper_paragraph(raw, prefixes) {
            append_accepted_text(&content, output);
        }
    };
    for boundary in 0..=paragraph.runs.len() {
        for owner in boundary_owners(paragraph, boundary) {
            let mut runs = Vec::new();
            match owner {
                BoundaryOwner::ContentControl(index) => {
                    append_accepted_control_runs(&paragraph.content_controls[index].3, &mut runs);
                }
                BoundaryOwner::Revision(index) => {
                    append_accepted_revision_runs(&paragraph.revisions[index].2, &mut runs);
                }
                BoundaryOwner::Wrapper(index) => {
                    wrapper_text(&paragraph.extra_xml[index].1, &["w".to_owned()], output);
                }
            }
            output.extend(runs.iter().map(|run| run.text()));
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            match simple_field_source(run) {
                Some((raw, prefixes)) => wrapper_text(raw, prefixes, output),
                None => output.push_str(&run.text()),
            }
        }
    }
}

struct CommentTextProjection {
    text: String,
    open: bool,
    start: Option<usize>,
    end: Option<usize>,
    start_offset: Option<usize>,
    end_offset: Option<usize>,
    markers: Vec<std::ops::Range<usize>>,
}

fn comment_projection_paragraph(
    xml: &[u8],
    runs: &[std::ops::Range<usize>],
    fields: &[std::ops::Range<usize>],
    window: std::ops::Range<usize>,
) -> Result<CT_P> {
    let kept = |span: &std::ops::Range<usize>| window.start <= span.start && span.end <= window.end;
    let mut removed = runs
        .iter()
        .filter(|span| !kept(span))
        .cloned()
        .collect::<Vec<_>>();
    removed.extend(
        fields
            .iter()
            .filter(|field| {
                !kept(field)
                    && !runs
                        .iter()
                        .any(|run| field.start <= run.start && run.end <= field.end && kept(run))
            })
            .cloned(),
    );
    removed.sort_by_key(|span| (span.start, std::cmp::Reverse(span.end)));
    let mut outer = Vec::<std::ops::Range<usize>>::new();
    for span in removed {
        if outer
            .last()
            .is_some_and(|previous| span.end <= previous.end)
        {
            continue;
        }
        outer.push(span);
    }
    let mut projected = xml.to_vec();
    for span in outer.into_iter().rev() {
        projected.drain(span);
    }
    CT_P::from_xml_fragment(&projected)
}

fn project_comment_text(xml: &[u8], id: i32, incoming: bool) -> Result<CommentTextProjection> {
    let invalid = |reason: &str| OxmlError::InvalidValue(format!("comment {id}: {reason}"));
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<(Vec<u8>, bool, usize, Option<usize>)>::new();
    let mut runs = Vec::new();
    let mut fields = Vec::new();
    let mut markers = Vec::<(bool, std::ops::Range<usize>)>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        let word = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == crate::namespace::W_NS.as_bytes());
        drop(namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let mut marker = None;
                if word && matches!(local.as_slice(), b"commentRangeStart" | b"commentRangeEnd") {
                    let mut selected_id = None;
                    for attribute in element.attributes() {
                        let attribute = attribute?;
                        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
                        if name.as_ref() == b"id"
                            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == crate::namespace::W_NS.as_bytes())
                        {
                            if selected_id.is_some() {
                                return Err(invalid("duplicate identity attribute"));
                            }
                            selected_id = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )?
                                    .parse::<i32>()
                                    .map_err(|_| invalid("invalid marker id"))?,
                            );
                        }
                    }
                    if selected_id == Some(id)
                        && !stack
                            .iter()
                            .any(|(name, word, _, _)| *word && name == b"txbxContent")
                    {
                        if stack.iter().any(|(name, word, _, _)| {
                            *word
                                && matches!(name.as_slice(), b"del" | b"moveFrom" | b"txbxContent")
                        }) {
                            return Err(invalid(
                                "selected source is outside the accepted paragraph view",
                            ));
                        }
                        if stack.iter().any(|(name, word, _, _)| *word && name == b"r") {
                            return Err(invalid(
                                "existing accepted run axis cannot faithfully represent a mid-run source boundary",
                            ));
                        }
                        if stack.iter().any(|(name, word, _, _)| {
                            !*word
                                || !matches!(
                                    name.as_slice(),
                                    b"p" | b"sdt"
                                        | b"sdtContent"
                                        | b"hyperlink"
                                        | b"ins"
                                        | b"moveTo"
                                        | b"smartTag"
                                        | b"customXml"
                                        | b"fldSimple"
                                )
                        }) {
                            return Err(invalid(
                                "selected source has an unsupported opaque owner outside the accepted paragraph view",
                            ));
                        }
                        markers.push((local == b"commentRangeStart", before..after));
                        marker = Some(markers.len() - 1);
                    }
                }
                if matches!(event, Event::Start(_)) {
                    stack.push((local, word, before, marker));
                } else if word && local == b"r" {
                    runs.push(before..after);
                } else if word && local == b"fldSimple" {
                    fields.push(before..after);
                }
            }
            Event::End(_) => {
                if let Some((name, word, start, marker)) = stack.pop() {
                    if let Some(index) = marker {
                        if !xml[markers[index].1.end..before]
                            .iter()
                            .all(u8::is_ascii_whitespace)
                        {
                            return Err(invalid(
                                "selected range marker has unsupported child content",
                            ));
                        }
                        markers[index].1.end = after;
                    }
                    if word && name == b"r" {
                        runs.push(start..after);
                    }
                    if word && name == b"fldSimple" {
                        fields.push(start..after);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut open = incoming;
    let mut start_span = None;
    let mut end_span = None;
    for (start, span) in &markers {
        if *start {
            if open || start_span.is_some() {
                return Err(invalid("duplicate or reversed selected start"));
            }
            start_span = Some(span.clone());
            open = true;
        } else {
            if !open || end_span.is_some() {
                return Err(invalid("unmatched or reversed selected end"));
            }
            end_span = Some(span.clone());
            open = false;
        }
    }
    let prefix = |at: usize| -> Result<(usize, usize)> {
        let p = comment_projection_paragraph(xml, &runs, &fields, 0..at)?;
        Ok((p.accepted_run_paths().len(), p.accepted_text().len()))
    };
    let start = start_span
        .as_ref()
        .map(|span| prefix(span.start))
        .transpose()?;
    let end = end_span
        .as_ref()
        .map(|span| prefix(span.start))
        .transpose()?;
    let text = if incoming || start_span.is_some() {
        comment_projection_paragraph(
            xml,
            &runs,
            &fields,
            start_span.as_ref().map_or(0, |span| span.end)
                ..end_span.as_ref().map_or(xml.len(), |span| span.start),
        )?
        .accepted_text()
    } else {
        String::new()
    };
    Ok(CommentTextProjection {
        text,
        open,
        start: start.map(|value| value.0),
        end: end.map(|value| value.0),
        start_offset: start.map(|value| value.1),
        end_offset: end.map(|value| value.1),
        markers: markers.into_iter().map(|(_, span)| span).collect(),
    })
}

fn accepted_paragraph_run_paths(paragraph: &CT_P) -> Vec<AcceptedRunPath> {
    let mut output = Vec::new();
    let mut prefix = Vec::new();
    append_accepted_paragraph_run_paths(paragraph, &mut prefix, &mut output);
    output
}

pub(crate) fn append_accepted_paragraph_run_paths(
    paragraph: &CT_P,
    prefix: &mut Vec<AcceptedRunPathSegment>,
    output: &mut Vec<AcceptedRunPath>,
) {
    for boundary in 0..=paragraph.runs.len() {
        for owner in boundary_owners(paragraph, boundary) {
            match owner {
                BoundaryOwner::ContentControl(index) => {
                    prefix.push(AcceptedRunPathSegment::ContentControl(index));
                    paragraph.content_controls[index]
                        .3
                        .append_accepted_run_paths(prefix, output);
                }
                BoundaryOwner::Revision(index) => {
                    prefix.push(AcceptedRunPathSegment::Revision(index));
                    paragraph.revisions[index]
                        .2
                        .append_accepted_run_paths(prefix, output);
                }
                BoundaryOwner::Wrapper(_) => continue,
            }
            prefix.pop();
        }
        if paragraph.runs.get(boundary).is_some() {
            prefix.push(AcceptedRunPathSegment::Run(boundary));
            output.push(AcceptedRunPath {
                segments: prefix.clone(),
            });
            prefix.pop();
        }
    }
}

fn accepted_paragraph_runs(paragraph: &CT_P) -> Vec<&CT_R> {
    let mut output = Vec::new();
    for boundary in 0..=paragraph.runs.len() {
        for owner in boundary_owners(paragraph, boundary) {
            match owner {
                BoundaryOwner::ContentControl(index) => {
                    append_accepted_control_runs(&paragraph.content_controls[index].3, &mut output);
                }
                BoundaryOwner::Revision(index) => {
                    append_accepted_revision_runs(&paragraph.revisions[index].2, &mut output);
                }
                BoundaryOwner::Wrapper(_) => {}
            }
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            output.push(run);
        }
    }
    output
}

fn append_accepted_control_runs<'a>(control: &'a CT_Sdt, output: &mut Vec<&'a CT_R>) {
    for boundary in 0..=control.content.len() {
        for (_, revision) in control.revisions().iter().filter(|(at, _)| *at == boundary) {
            append_accepted_revision_runs(revision, output);
        }
        let Some(content) = control.content.get(boundary) else {
            continue;
        };
        match content {
            SdtContent::Run(run) => output.push(run),
            SdtContent::ContentControl(control) => append_accepted_control_runs(control, output),
            SdtContent::Paragraph(paragraph) => output.extend(accepted_paragraph_runs(paragraph)),
            SdtContent::Table(table) => append_accepted_table_runs(table, output),
            SdtContent::Row(row) => append_accepted_row_runs(row, output),
            SdtContent::Cell(cell) => append_accepted_cell_runs(cell, output),
            SdtContent::RawXml(_) => {}
        }
    }
}

fn append_accepted_table_runs<'a>(table: &'a CT_Tbl, output: &mut Vec<&'a CT_R>) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_accepted_control_runs(control, output);
        }
        if let Some(row) = table.rows.get(boundary) {
            append_accepted_row_runs(row, output);
        }
    }
}

fn append_accepted_row_runs<'a>(row: &'a CT_Row, output: &mut Vec<&'a CT_R>) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_accepted_control_runs(control, output);
        }
        if let Some(cell) = row.cells.get(boundary) {
            append_accepted_cell_runs(cell, output);
        }
    }
}

fn append_accepted_cell_runs<'a>(cell: &'a CT_Tc, output: &mut Vec<&'a CT_R>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => output.extend(accepted_paragraph_runs(paragraph)),
            CellContent::Table(table) => append_accepted_table_runs(table, output),
            CellContent::ContentControl(control) => append_accepted_control_runs(control, output),
        }
    }
}

fn tracked_paragraph_runs(paragraph: &CT_P) -> Vec<&CT_R> {
    let mut output = Vec::new();
    for boundary in 0..=paragraph.runs.len() {
        for owner in boundary_owners(paragraph, boundary) {
            match owner {
                BoundaryOwner::ContentControl(index) => {
                    append_tracked_control_runs(&paragraph.content_controls[index].3, &mut output);
                }
                BoundaryOwner::Revision(index) => {
                    append_tracked_revision_runs(&paragraph.revisions[index].2, &mut output);
                }
                BoundaryOwner::Wrapper(_) => {}
            }
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            output.push(run);
        }
    }
    output
}

fn append_tracked_control_runs<'a>(control: &'a CT_Sdt, output: &mut Vec<&'a CT_R>) {
    for boundary in 0..=control.content.len() {
        for (_, revision) in control.revisions().iter().filter(|(at, _)| *at == boundary) {
            append_tracked_revision_runs(revision, output);
        }
        let Some(content) = control.content.get(boundary) else {
            continue;
        };
        match content {
            SdtContent::Run(run) => output.push(run),
            SdtContent::ContentControl(control) => append_tracked_control_runs(control, output),
            SdtContent::Paragraph(paragraph) => output.extend(tracked_paragraph_runs(paragraph)),
            SdtContent::Table(table) => append_tracked_table_runs(table, output),
            SdtContent::Row(row) => append_tracked_row_runs(row, output),
            SdtContent::Cell(cell) => append_tracked_cell_runs(cell, output),
            SdtContent::RawXml(_) => {}
        }
    }
}

fn append_tracked_revision_runs<'a>(revision: &'a CT_Revision, output: &mut Vec<&'a CT_R>) {
    if let Some(paragraph) = revision.content_paragraph() {
        output.extend(tracked_paragraph_runs(paragraph));
        return;
    }
    let crate::revision::RevisionContent::Runs(runs) = revision.content() else {
        return;
    };
    for boundary in 0..=runs.len() {
        for (_, nested) in revision
            .nested_revisions()
            .iter()
            .filter(|(at, _)| *at == boundary)
        {
            append_tracked_revision_runs(nested, output);
        }
        if let Some(run) = runs.get(boundary) {
            output.push(run);
        }
    }
}

fn append_tracked_table_runs<'a>(table: &'a CT_Tbl, output: &mut Vec<&'a CT_R>) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_tracked_control_runs(control, output);
        }
        if let Some(row) = table.rows.get(boundary) {
            append_tracked_row_runs(row, output);
        }
    }
}

fn append_tracked_row_runs<'a>(row: &'a CT_Row, output: &mut Vec<&'a CT_R>) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_tracked_control_runs(control, output);
        }
        if let Some(cell) = row.cells.get(boundary) {
            append_tracked_cell_runs(cell, output);
        }
    }
}

fn append_tracked_cell_runs<'a>(cell: &'a CT_Tc, output: &mut Vec<&'a CT_R>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => output.extend(tracked_paragraph_runs(paragraph)),
            CellContent::Table(table) => append_tracked_table_runs(table, output),
            CellContent::ContentControl(control) => append_tracked_control_runs(control, output),
        }
    }
}

fn append_accepted_revision_runs<'a>(revision: &'a CT_Revision, output: &mut Vec<&'a CT_R>) {
    if !matches!(
        revision.kind(),
        RevisionKind::Insertion | RevisionKind::MoveTo
    ) {
        return;
    }
    if let Some(paragraph) = revision.content_paragraph() {
        output.extend(accepted_paragraph_runs(paragraph));
    }
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
            rubies: Vec::new(),
        }
    }

    /// Report whether one raw paragraph carrier retains root attributes.
    #[doc(hidden)]
    pub fn raw_is_root_attributes(position: usize, raw: &[u8]) -> bool {
        position == ROOT_ATTRIBUTES_POSITION && is_root_attribute_record(raw)
    }

    pub(crate) fn from_empty_root(root: &BytesStart<'_>, word_prefixes: &[String]) -> Result<Self> {
        let mut paragraph = Self::new();
        if let Some(record) = capture_root_attribute_record(root, word_prefixes)? {
            paragraph.extra_xml.push((ROOT_ATTRIBUTES_POSITION, record));
        }
        Ok(paragraph)
    }

    /// Get the combined text of all runs in this paragraph.
    pub fn text(&self) -> String {
        self.runs().iter().map(|run| run.text()).collect()
    }

    /// Whether the accepted view of this paragraph shows anything, so that
    /// a paragraph without it reads as empty: text other than white space,
    /// an equation, or a run that draws without text, such as a drawing, a
    /// picture, an object, a field, a symbol, a tab, a break or a note
    /// reference. The runs inside content controls, insertions, smart tags,
    /// inline custom XML elements and bidirectional embeddings or overrides
    /// count, deleted runs do not.
    pub fn has_visible_content(&self) -> bool {
        !self.equations.is_empty()
            || self
                .accepted_bookmark_runs()
                .into_iter()
                .any(CT_R::has_visible_content)
            || self.extra_xml.iter().any(|(_, raw)| {
                wrapper_content_paragraph(
                    raw,
                    &["w".to_owned()],
                    &["smartTag", "customXml", "bdo", "dir"],
                )
                .is_some_and(|content| content.has_visible_content())
            })
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

    /// Borrow typed source runs in physical order, including revision owners.
    /// Deleted runs reserve identity but selection remains the caller's job.
    /// Opaque wrapper XML does not become a typed source.
    #[doc(hidden)]
    pub fn source_runs(&self) -> Vec<&CT_R> {
        tracked_paragraph_runs(self)
    }

    /// Return the text of the accepted view, as `Paragraph::text` reads it:
    /// the text of [`Self::accepted_bookmark_runs`], and that of the runs
    /// inside smart tags, inline custom XML elements and simple fields, in
    /// document order, nested ones included. The runs of a wrapper inside a
    /// content control, a revision or a hyperlink are not read.
    #[doc(hidden)]
    pub fn accepted_text(&self) -> String {
        let mut text = String::new();
        append_accepted_text(self, &mut text);
        text
    }

    /// Project one selected comment span through the existing accepted text and run axis.
    /// Returns display, continuation state and optional start/end run boundaries.
    /// Original source is untouched. Unrepresentable source endpoints are errors.
    #[doc(hidden)]
    pub fn accepted_comment_projection(
        &self,
        id: i32,
        open_at_start: bool,
    ) -> Result<(String, bool, Option<usize>, Option<usize>)> {
        if self.comment_ranges.iter().any(|marker| match marker {
            CommentRangeMarker::Start {
                id: marker_id,
                has_child_content,
                ..
            }
            | CommentRangeMarker::End {
                id: marker_id,
                has_child_content,
                ..
            } => *marker_id == id && *has_child_content,
        }) {
            return Err(OxmlError::InvalidValue(format!(
                "comment {id}: selected range marker has unsupported child content"
            )));
        }
        let mut raw = Vec::new();
        self.to_xml(&mut Writer::new(&mut raw))?;
        Self::accepted_comment_source_projection(&raw, id, open_at_start)
    }

    /// Project a comment from its namespace-closed original paragraph source.
    /// Retain its exact inherited bindings through the private fidelity probe.
    /// The accepted text, endpoint axis and unsupported-source checks are the
    /// same as [`Self::accepted_comment_projection`].
    #[doc(hidden)]
    #[allow(clippy::type_complexity)] // Concrete text, open state and two accepted endpoints.
    pub fn accepted_comment_source_projection(
        xml: &[u8],
        id: i32,
        open_at_start: bool,
    ) -> Result<(String, bool, Option<usize>, Option<usize>)> {
        let raw = raw_with_external_bindings(
            xml,
            &[("w".to_owned(), crate::namespace::W_NS.to_owned())],
        )?;
        let mut root_reader = Reader::from_reader(raw.as_slice());
        let bindings = loop {
            match root_reader.read_event()? {
                Event::Start(element) | Event::Empty(element) => {
                    break raw_namespace_declarations(&element)?;
                }
                Event::Eof => {
                    return Err(OxmlError::MissingElement("comment paragraph root".into()));
                }
                _ => {}
            }
        };
        let selected = project_comment_text(&raw, id, open_at_start)?;
        if !selected.markers.is_empty() {
            // Remove only this selected pair before asking the existing checked
            // writer to place the same endpoints. No temporary identity is allocated.
            let mut unmarked = raw;
            for span in selected.markers.iter().rev() {
                unmarked.drain(span.clone());
            }
            let mut probe = CT_P::from_xml_fragment(&unmarked)?;
            probe.anchor_accepted_range(selected.start, selected.end, RangeAnchor::Comment(id))
                .map_err(|error| OxmlError::InvalidValue(format!("comment {id}: existing accepted run axis cannot faithfully represent selected source: {error}")))?;
            let mut raw = Vec::new();
            probe.to_xml(&mut Writer::new(&mut raw))?;
            let raw = raw_with_external_bindings(&raw, &bindings)?;
            let raw = raw_with_external_bindings(
                &raw,
                &[("w".to_owned(), crate::namespace::W_NS.to_owned())],
            )?;
            let placed = project_comment_text(&raw, id, open_at_start)?;
            if selected.text != placed.text
                || selected.start_offset != placed.start_offset
                || selected.end_offset != placed.end_offset
            {
                return Err(OxmlError::InvalidValue(format!(
                    "comment {id}: existing accepted run axis cannot faithfully represent selected source"
                )));
            }
        }
        Ok((selected.text, selected.open, selected.start, selected.end))
    }

    /// Return the accepted view as a paragraph of plain runs, for the
    /// exporters: the runs whose text [`Self::accepted_text`] reads, in
    /// document order. Those of content controls, tracked insertions and
    /// moves in, smart tags, inline custom XML elements and simple fields
    /// are included, deleted and moved-away runs are left out, and the
    /// hyperlink spans cover the same runs as in the source. Content
    /// controls, revisions, wrappers and markers are not kept. Borrowed when
    /// the paragraph holds none of them, since its runs are then the view.
    #[doc(hidden)]
    pub fn accepted_view(&self) -> Cow<'_, CT_P> {
        if self.content_controls.is_empty()
            && self.revisions.is_empty()
            && !self.extra_xml.iter().any(|(_, raw)| is_run_wrapper(raw))
            && !self
                .runs
                .iter()
                .any(|run| simple_field_source(run).is_some())
        {
            return Cow::Borrowed(self);
        }
        let mut runs = Vec::new();
        // The index into `self.hyperlinks` of the hyperlink each run is in.
        let mut links = Vec::new();
        for boundary in 0..=self.runs.len() {
            for owner in boundary_owners(self, boundary) {
                let mut owned = Vec::new();
                let link = match owner {
                    BoundaryOwner::ContentControl(index) => {
                        append_accepted_control_runs(&self.content_controls[index].3, &mut owned);
                        None
                    }
                    BoundaryOwner::Revision(index) => {
                        let (_, slot, revision) = &self.revisions[index];
                        append_accepted_revision_runs(revision, &mut owned);
                        hyperlink_revision_index(*slot)
                    }
                    BoundaryOwner::Wrapper(index) => {
                        if let Some(content) =
                            run_wrapper_paragraph(&self.extra_xml[index].1, &["w".to_owned()])
                        {
                            let view = content.accepted_view().into_owned();
                            links.extend(std::iter::repeat_n(None, view.runs.len()));
                            runs.extend(view.runs);
                        }
                        continue;
                    }
                };
                links.extend(std::iter::repeat_n(link, owned.len()));
                runs.extend(owned.into_iter().cloned());
            }
            if let Some(run) = self.runs.get(boundary) {
                let link = self.hyperlinks.iter().position(|hyperlink| {
                    hyperlink.run_start <= boundary && boundary < hyperlink.run_end
                });
                // A simple field reads as the runs of its result, as in
                // `Self::accepted_text`, so a revision inside it is resolved.
                let field = simple_field_source(run)
                    .and_then(|(raw, prefixes)| run_wrapper_paragraph(raw, prefixes));
                match field {
                    Some(content) => {
                        let view = content.accepted_view().into_owned();
                        links.extend(std::iter::repeat_n(link, view.runs.len()));
                        runs.extend(view.runs);
                    }
                    None => {
                        runs.push(run.clone());
                        links.push(link);
                    }
                }
            }
        }
        let mut hyperlinks = Vec::new();
        let mut start = 0;
        while start < runs.len() {
            let mut end = start + 1;
            while end < runs.len() && links[end] == links[start] {
                end += 1;
            }
            if let Some(link) = links[start] {
                hyperlinks.push(HyperlinkSpan {
                    run_start: start,
                    run_end: end,
                    ..self.hyperlinks[link].clone()
                });
            }
            start = end;
        }
        Cow::Owned(CT_P {
            properties: self.properties.clone(),
            runs,
            hyperlinks,
            ..CT_P::new()
        })
    }

    /// Return the content of a preserved smart tag or inline custom XML
    /// element, which [`Self::accepted_view`] reads in place. None for any
    /// other raw paragraph child.
    #[doc(hidden)]
    pub fn raw_run_wrapper_content(raw: &[u8]) -> Option<CT_P> {
        if !is_run_wrapper(raw) {
            return None;
        }
        run_wrapper_paragraph(raw, &["w".to_owned()])
    }

    /// Return accepted-view runs in the same order as bookmark projections.
    #[doc(hidden)]
    pub fn accepted_bookmark_runs(&self) -> Vec<&CT_R> {
        accepted_paragraph_runs(self)
    }

    /// Return recursive addresses for every accepted-view run.
    #[doc(hidden)]
    pub fn accepted_run_paths(&self) -> Vec<AcceptedRunPath> {
        accepted_paragraph_run_paths(self)
    }

    /// Resolve one recursive accepted-view run address.
    #[doc(hidden)]
    pub fn accepted_run(&self, path: &AcceptedRunPath) -> Option<&CT_R> {
        self.accepted_run_segments(path.segments())
    }

    pub(crate) fn accepted_run_segments(&self, path: &[AcceptedRunPathSegment]) -> Option<&CT_R> {
        let (first, rest) = path.split_first()?;
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => self.runs.get(index),
            AcceptedRunPathSegment::ContentControl(index) => self
                .content_controls
                .get(index)?
                .3
                .accepted_run_segments(rest),
            AcceptedRunPathSegment::Revision(index) => {
                self.revisions.get(index)?.2.accepted_run_segments(rest)
            }
            AcceptedRunPathSegment::Run(_) => None,
        }
    }

    /// Replace one accepted-view run while retaining its recursive owner.
    #[doc(hidden)]
    pub fn replace_accepted_run(
        &mut self,
        path: &AcceptedRunPath,
        replacement: CT_R,
    ) -> Result<bool> {
        self.replace_accepted_run_segments(path.segments(), replacement)
    }

    pub(crate) fn replace_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
        replacement: CT_R,
    ) -> Result<bool> {
        let Some((first, rest)) = path.split_first() else {
            return Ok(false);
        };
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => {
                let Some(run) = self.runs.get_mut(index) else {
                    return Ok(false);
                };
                *run = replacement;
                Ok(true)
            }
            AcceptedRunPathSegment::ContentControl(index) => {
                let Some((_, _, _, control)) = self.content_controls.get_mut(index) else {
                    return Ok(false);
                };
                control.replace_accepted_run_segments(rest, replacement)
            }
            AcceptedRunPathSegment::Revision(index) => {
                let Some((_, _, revision)) = self.revisions.get_mut(index) else {
                    return Ok(false);
                };
                revision.replace_accepted_run_segments(rest, replacement)
            }
            AcceptedRunPathSegment::Run(_) => Ok(false),
        }
    }

    /// Split one accepted-view run and return the accepted boundary.
    #[doc(hidden)]
    pub fn split_accepted_run(
        &mut self,
        path: &AcceptedRunPath,
        accepted_index: usize,
        offset: usize,
    ) -> Result<usize> {
        let run = self
            .accepted_run(path)
            .ok_or_else(|| OxmlError::InvalidValue("accepted run path is stale".to_owned()))?;
        let literal_len = run.literal_len();
        if offset > literal_len {
            return Err(OxmlError::InvalidValue(format!(
                "split offset {offset} exceeds a run's literal text length of {literal_len} characters"
            )));
        }
        if offset == 0 {
            return Ok(accepted_index);
        }
        if offset == literal_len {
            return Ok(accepted_index + 1);
        }
        if self.split_accepted_run_segments(path.segments(), offset)? {
            Ok(accepted_index + 1)
        } else {
            Err(OxmlError::InvalidValue(
                "accepted run path is stale".to_owned(),
            ))
        }
    }

    /// Return the literal text of the accepted-view runs, the text that
    /// split offsets count. Tabs, breaks and other non-text content have no
    /// width.
    #[doc(hidden)]
    pub fn accepted_literal_text(&self) -> String {
        accepted_paragraph_runs(self)
            .into_iter()
            .flat_map(|run| run.content.iter().map(CT_R::literal_text))
            .collect()
    }

    /// Return the text of every `t` element, in any namespace, between the
    /// range markers of comment `id`, or `None` when this paragraph does not
    /// hold both markers.
    ///
    /// Unlike [`Self::accepted_literal_text`], this includes the text of
    /// preserved children that the accepted view leaves out, such as a
    /// `w:fldSimple` result, so it is the text the commented range shows.
    #[doc(hidden)]
    pub fn comment_range_text(&self, id: i32) -> Option<String> {
        let mut xml = Vec::new();
        self.to_xml(&mut Writer::new(&mut xml)).ok()?;
        let is_marker = |element: &BytesStart<'_>, local: &[u8]| {
            matches_local_name(element.name().as_ref(), local)
                && element
                    .attributes()
                    .flatten()
                    .find(|attribute| matches_local_name(attribute.key.as_ref(), b"id"))
                    .and_then(|attribute| std::str::from_utf8(&attribute.value).ok()?.parse().ok())
                    == Some(id)
        };
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        let mut text = None::<String>;
        loop {
            match reader.read_event_into(&mut buffer).ok()? {
                Event::Start(element) | Event::Empty(element)
                    if is_marker(&element, b"commentRangeStart") =>
                {
                    text = Some(String::new());
                }
                Event::Start(element) | Event::Empty(element)
                    if is_marker(&element, b"commentRangeEnd") =>
                {
                    return text;
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"t") => {
                    let content = crate::xml_text::read_element_text(&mut reader, element.name());
                    if let Some(text) = text.as_mut() {
                        text.push_str(&content);
                    }
                }
                Event::Eof => return None,
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Read the literal range safeguard from its namespace-qualified source.
    /// Only qualified Word markers, identity attributes and `t` elements count.
    /// Tabs and breaks have zero width, while field-result text remains visible.
    /// This does not change the richer accepted anchor projection.
    #[doc(hidden)]
    pub fn comment_range_text_from_source(xml: &[u8], id: i32) -> Result<Option<String>> {
        // Validate the complete source separately. Text extraction below skips
        // `t` subtrees and ends at the selected marker, neither of which may
        // hide a missing namespace declaration elsewhere in the candidate.
        let mut validation = NsReader::from_reader(xml);
        loop {
            let (namespace, event) = validation.read_resolved_event()?;
            if let ResolveResult::Unknown(prefix) = namespace {
                return Err(OxmlError::InvalidValue(format!(
                    "comment literal source has unresolved element namespace prefix {:?}",
                    String::from_utf8_lossy(&prefix)
                )));
            }
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    for attribute in element.attributes() {
                        let attribute = attribute?;
                        attribute.decoded_and_normalized_value(
                            XmlVersion::Implicit1_0,
                            validation.decoder(),
                        )?;
                        if !namespace_declaration(attribute.key.as_ref())
                            && let (ResolveResult::Unknown(prefix), _) =
                                validation.resolver().resolve_attribute(attribute.key)
                        {
                            return Err(OxmlError::InvalidValue(format!(
                                "comment literal source has unresolved attribute namespace prefix {:?}",
                                String::from_utf8_lossy(&prefix)
                            )));
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        let mut reader = NsReader::from_reader(xml);
        let mut text = None::<String>;
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            let word = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == crate::namespace::W_NS.as_bytes());
            drop(namespace);
            match event {
                Event::Start(element) | Event::Empty(element)
                    if word
                        && matches!(
                            element.local_name().as_ref(),
                            b"commentRangeStart" | b"commentRangeEnd"
                        ) =>
                {
                    let mut marker_id = None;
                    for attribute in element.attributes() {
                        let attribute = attribute?;
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        if local.as_ref() == b"id"
                            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == crate::namespace::W_NS.as_bytes())
                        {
                            marker_id = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )?
                                    .parse::<i32>()
                                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?,
                            );
                        }
                    }
                    if marker_id == Some(id) {
                        if element.local_name().as_ref() == b"commentRangeEnd" {
                            return Ok(text);
                        }
                        text = Some(String::new());
                    }
                }
                Event::Start(element) if word && element.local_name().as_ref() == b"t" => {
                    let encoded = reader.read_text(element.name())?;
                    if let Some(text) = text.as_mut() {
                        text.push_str(&crate::xml_text::decode_escaped(&encoded));
                    }
                }
                Event::Eof => return Ok(None),
                _ => {}
            }
        }
    }

    /// Split accepted-view runs so that the non-empty literal text span
    /// `[start, end)`, in Unicode scalar values of
    /// [`Self::accepted_literal_text`], covers whole runs, and return the run
    /// boundaries around it.
    #[doc(hidden)]
    pub fn split_accepted_literal_span(
        &mut self,
        start: usize,
        end: usize,
    ) -> Result<(usize, usize)> {
        let lengths = accepted_paragraph_runs(self)
            .into_iter()
            .map(CT_R::literal_len)
            .collect::<Vec<_>>();
        // The run holding one literal character and its offset in that run.
        let locate = |character: usize| {
            let mut run_start = 0;
            lengths.iter().enumerate().find_map(|(index, len)| {
                let found = (character < run_start + len).then_some((index, character - run_start));
                run_start += len;
                found
            })
        };
        let outside = || {
            OxmlError::InvalidValue(format!(
                "literal span {start}..{end} is not inside the paragraph text"
            ))
        };
        let ((start_run, start_offset), (end_run, end_offset)) = match end.checked_sub(1) {
            Some(last) if start < end => (
                locate(start).ok_or_else(outside)?,
                locate(last).ok_or_else(outside)?,
            ),
            _ => return Err(outside()),
        };
        // The end goes first, so the start run keeps its index.
        let path = self.accepted_run_paths()[end_run].clone();
        let end_boundary = self.split_accepted_run(&path, end_run, end_offset + 1)?;
        let path = self.accepted_run_paths()[start_run].clone();
        let start_boundary = self.split_accepted_run(&path, start_run, start_offset)?;
        // A split inside the start run adds one run before the end boundary.
        Ok((start_boundary, end_boundary + usize::from(start_offset > 0)))
    }

    pub(crate) fn split_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
        offset: usize,
    ) -> Result<bool> {
        let Some((first, rest)) = path.split_first() else {
            return Ok(false);
        };
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => self
                .split_run(index, offset)
                .map(|_| true)
                .map_err(|error| OxmlError::InvalidValue(error.to_string())),
            AcceptedRunPathSegment::ContentControl(index) => {
                let Some((_, _, _, control)) = self.content_controls.get_mut(index) else {
                    return Ok(false);
                };
                control.split_accepted_run_segments(rest, offset)
            }
            AcceptedRunPathSegment::Revision(index) => {
                let Some((_, _, revision)) = self.revisions.get_mut(index) else {
                    return Ok(false);
                };
                revision.split_accepted_run_segments(rest, offset)
            }
            AcceptedRunPathSegment::Run(_) => Ok(false),
        }
    }

    /// Remove one accepted-view run.
    ///
    /// The comment, bookmark and permission markers and the other preserved
    /// children around the run stay where they are. A run inside a
    /// hyperlink, an inline content control or a tracked insertion is removed
    /// inside it. A hyperlink or a tracked insertion left with nothing in it
    /// goes too, while a content control stays with its properties, as Word
    /// keeps an emptied control to show its placeholder. A run holding a
    /// comment, footnote or endnote reference, part of a complex field whose
    /// other parts are in other runs, or part of a tracked move destination
    /// is refused, since the note, the field or the move would lose its
    /// balance. On
    /// error the paragraph is unchanged.
    #[doc(hidden)]
    pub fn remove_accepted_run(&mut self, path: &AcceptedRunPath) -> Result<()> {
        let stale = || OxmlError::InvalidValue("accepted run path is stale".to_owned());
        let run = self.accepted_run(path).ok_or_else(stale)?;
        if let Some(id) = run.content.iter().find_map(|content| match content {
            RunContent::CommentReference { id, .. } => Some(*id),
            _ => None,
        }) {
            return Err(OxmlError::InvalidValue(format!(
                "it holds the reference of comment {id}, which removing the comment removes"
            )));
        }
        if let Some((kind, id)) = run.content.iter().find_map(|content| match content {
            RunContent::FootnoteRef { id, .. } => Some(("footnote", *id)),
            RunContent::EndnoteRef { id, .. } => Some(("endnote", *id)),
            _ => None,
        }) {
            return Err(OxmlError::InvalidValue(format!(
                "it holds the reference of {kind} {id}, which would be left without one"
            )));
        }
        if run
            .extra_xml
            .iter()
            .any(|raw| raw_is_complex_field_part(raw))
        {
            return Err(OxmlError::InvalidValue(
                "it holds part of a complex field whose other parts are in other runs".to_owned(),
            ));
        }
        let mut paragraph = self.clone();
        if !paragraph.remove_accepted_run_segments(path.segments())? {
            return Err(stale());
        }
        // Accepted run indexes after the run moved down by one.
        let _ = paragraph.refresh_bookmark_projection();
        *self = paragraph;
        Ok(())
    }

    /// Remove the accepted-view run at `path`, and a tracked insertion left
    /// with nothing in it. Returns false for a stale path.
    pub(crate) fn remove_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
    ) -> Result<bool> {
        let Some((first, rest)) = path.split_first() else {
            return Ok(false);
        };
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => {
                if index >= self.runs.len() {
                    return Ok(false);
                }
                let mut removed = vec![false; self.runs.len()];
                removed[index] = true;
                self.remove_runs(&removed);
                Ok(true)
            }
            AcceptedRunPathSegment::ContentControl(index) => {
                let Some((_, _, _, control)) = self.content_controls.get_mut(index) else {
                    return Ok(false);
                };
                control.remove_accepted_run_segments(rest)
            }
            AcceptedRunPathSegment::Revision(index) => {
                let Some((_, _, revision)) = self.revisions.get_mut(index) else {
                    return Ok(false);
                };
                match revision.remove_accepted_run_segments(rest)? {
                    None => Ok(false),
                    Some(emptied) => {
                        if emptied {
                            self.remove_revision(index);
                        }
                        Ok(true)
                    }
                }
            }
            AcceptedRunPathSegment::Run(_) => Ok(false),
        }
    }

    /// Whether the paragraph has no child at all, attributes aside.
    pub(crate) fn has_no_children(&self) -> bool {
        self.runs.is_empty()
            && self.hyperlinks.is_empty()
            && self.comment_ranges.is_empty()
            && self.bookmark_markers.is_empty()
            && self.content_controls.is_empty()
            && self.revisions.is_empty()
            && self.equations.is_empty()
            && self.rubies.is_empty()
            && self
                .extra_xml
                .iter()
                .all(|(position, raw)| Self::raw_is_root_attributes(*position, raw))
    }

    /// Remove the revision wrapper at `index` with the placeholder that holds
    /// its place among the paragraph children, or its place inside a
    /// hyperlink, and the hyperlink when nothing else is left in it.
    fn remove_revision(&mut self, index: usize) {
        let (boundary, slot, _) = self.revisions.remove(index);
        let Some(hyperlink_index) = hyperlink_revision_index(slot) else {
            self.remove_boundary_raw(boundary, slot);
            return;
        };
        // The children of a hyperlink count the revisions written before them.
        let position = self.revisions[..index]
            .iter()
            .filter(|(at, other, _)| *at == boundary && *other == slot)
            .count();
        let Some(hyperlink) = self.hyperlinks.get_mut(hyperlink_index) else {
            return;
        };
        let relative = boundary.saturating_sub(hyperlink.run_start);
        for (at, revisions_before, _) in &mut hyperlink.extra_xml {
            if *at == relative && *revisions_before > position {
                *revisions_before -= 1;
            }
        }
        let emptied = hyperlink.run_start == hyperlink.run_end
            && hyperlink.extra_xml.is_empty()
            && !self.revisions.iter().any(|(_, other, _)| *other == slot);
        if !emptied {
            return;
        }
        let hyperlink = self.hyperlinks.remove(hyperlink_index);
        for (_, other, _) in &mut self.revisions {
            if let Some(later) = hyperlink_revision_index(*other)
                && later > hyperlink_index
            {
                *other = hyperlink_revision_slot(later - 1);
            }
        }
        if let Some(raw_before) = hyperlink.preserved_raw_before {
            self.remove_boundary_raw(hyperlink.run_start, raw_before);
        }
    }

    /// Remove the preserved child at `slot` among those at run `boundary`,
    /// and move every projection that counts the children before it. The
    /// children on both sides of it then share one slot, where those that
    /// came before it stay first.
    fn remove_boundary_raw(&mut self, boundary: usize, slot: usize) {
        let Some(position) = self
            .extra_xml
            .iter()
            .enumerate()
            .filter(|(_, (at, _))| *at == boundary)
            .nth(slot)
            .map(|(position, _)| position)
        else {
            return;
        };
        self.extra_xml.remove(position);
        self.comment_ranges
            .sort_by_key(|marker| (marker.run_index(), marker.raw_before()));
        let merged_markers = self
            .comment_ranges
            .iter()
            .filter(|marker| marker.run_index() == boundary && marker.raw_before() == slot)
            .count();
        let shift = |before: &mut usize| {
            if *before > slot {
                *before -= 1;
            }
        };
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
                } if *run_index == boundary => shift(raw_before),
                _ => {}
            }
        }
        for (at, raw_before, markers_before, _) in &mut self.content_controls {
            if *at == boundary {
                if *raw_before == slot + 1 {
                    *markers_before += merged_markers;
                }
                shift(raw_before);
            }
        }
        for marker in &mut self.bookmark_markers {
            if marker.run_index == boundary {
                shift(&mut marker.raw_before);
            }
        }
        for (at, other, _) in &mut self.revisions {
            if *at == boundary && hyperlink_revision_index(*other).is_none() {
                shift(other);
            }
        }
        for (at, other, _) in &mut self.equations {
            if *at == boundary {
                shift(other);
            }
        }
        for hyperlink in &mut self.hyperlinks {
            if hyperlink.run_start == boundary
                && let Some(raw_before) = hyperlink.preserved_raw_before.as_mut()
            {
                shift(raw_before);
            }
        }
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
        self.shift_run_boundaries_from(run_index);

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
        self.shift_ruby_spans_for_insert(run_index);
        self.runs.insert(run_index, run);
        self.refresh_bookmark_projection()
    }

    /// Split the direct run at `run_index` at a Unicode scalar offset of its
    /// literal text.
    ///
    /// The run keeps the text before `offset`. A new direct run at
    /// `run_index + 1` receives the rest with cloned properties, inside the
    /// same hyperlink. Zero selects the boundary before the run and the literal
    /// text length selects the boundary after it without changing the paragraph.
    /// Tabs, breaks, fields, drawings, references, and raw children are
    /// zero-width for this operation.
    /// Text-free content and raw children at the split point stay in the first
    /// run, and boundary markers after the original run stay after the second.
    /// Only [`RunSplitError::BookmarkProjection`] leaves the paragraph changed.
    pub fn split_run(
        &mut self,
        run_index: usize,
        offset: usize,
    ) -> std::result::Result<usize, RunSplitError> {
        let run_count = self.runs.len();
        let literal_len = self
            .runs
            .get(run_index)
            .ok_or(RunSplitError::RunOutOfRange {
                run_index,
                run_count,
            })?
            .literal_len();
        if offset > literal_len {
            return Err(RunSplitError::OffsetOutOfRange {
                offset,
                len: literal_len,
            });
        }
        if offset == 0 {
            return Ok(run_index);
        }
        if offset == literal_len {
            return Ok(run_index + 1);
        }
        let tail = self.runs[run_index].split_off_at(offset)?;
        let inserted = run_index + 1;
        self.shift_run_boundaries_from(inserted);
        for hyperlink in &mut self.hyperlinks {
            if hyperlink.run_start >= inserted {
                hyperlink.run_start += 1;
                hyperlink.run_end += 1;
            } else if hyperlink.run_end >= inserted {
                // The hyperlink contains the split run, so it grows by one run.
                let boundary_after_run = inserted - hyperlink.run_start;
                for (boundary, _, _) in &mut hyperlink.extra_xml {
                    if *boundary >= boundary_after_run {
                        *boundary += 1;
                    }
                }
                hyperlink.run_end += 1;
            }
        }
        self.shift_ruby_spans_for_insert(inserted);
        self.runs.insert(inserted, tail);
        if self.refresh_bookmark_projection() {
            Ok(inserted)
        } else {
            Err(RunSplitError::BookmarkProjection)
        }
    }

    /// Move every run-boundary projection at or after `run_index` one run later.
    ///
    /// Hyperlink spans are left to the caller, which decides whether the new
    /// run joins or splits them.
    /// Move every ruby span across one run inserted at `inserted`.
    ///
    /// A span that starts at or after the insertion point moves whole, and a
    /// span that already contains the point grows by one run, which is the
    /// rule `w:hyperlink` spans follow beside this call.
    fn shift_ruby_spans_for_insert(&mut self, inserted: usize) {
        for ruby in &mut self.rubies {
            if ruby.base_start >= inserted {
                ruby.base_start += 1;
                ruby.base_end += 1;
            } else if ruby.base_end >= inserted {
                ruby.base_end += 1;
            }
        }
    }

    fn shift_run_boundaries_from(&mut self, run_index: usize) {
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
        for (position, _) in &mut self.extra_xml {
            if *position != ROOT_ATTRIBUTES_POSITION && *position >= run_index {
                *position += 1;
            }
        }
        for (position, _, _) in &mut self.equations {
            if *position >= run_index {
                *position += 1;
            }
        }
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
        self.remove_runs(&removed);
    }

    /// Remove the direct runs flagged in `removed` and move every run-boundary
    /// projection onto the boundary that remains. Raw children, comment
    /// markers, bookmarks and controls of the boundaries that collapse into
    /// one keep their order, and a hyperlink or a ruby annotation left
    /// without runs is dropped. A hyperlink that still holds a marker or a
    /// tracked change stays where it was, after the children at its start.
    pub(crate) fn remove_runs(&mut self, removed: &[bool]) {
        let removed_run_addresses = self
            .runs
            .iter()
            .zip(removed)
            .filter_map(|(run, remove)| remove.then_some(std::ptr::from_ref(run)))
            .collect::<Vec<_>>();
        let removed_projected_indices = accepted_paragraph_runs(self)
            .into_iter()
            .enumerate()
            .filter_map(|(index, run)| {
                removed_run_addresses
                    .contains(&std::ptr::from_ref(run))
                    .then_some(index)
            })
            .collect::<Vec<_>>();

        let old_run_count = self.runs.len();
        // A hyperlink keeping no run is written in place of a raw child at its
        // boundary. Its placeholder goes after the children at its start, so
        // the children at its end stay after it.
        let collapsed = (0..self.hyperlinks.len())
            .filter(|&index| {
                let hyperlink = &self.hyperlinks[index];
                hyperlink.preserved_raw_before.is_none()
                    && hyperlink.run_start < hyperlink.run_end
                    && hyperlink.run_end <= old_run_count
                    && removed[hyperlink.run_start..hyperlink.run_end]
                        .iter()
                        .all(|remove| *remove)
                    && (!hyperlink.extra_xml.is_empty()
                        || self
                            .revisions
                            .iter()
                            .any(|(_, slot, _)| hyperlink_revision_index(*slot) == Some(index)))
            })
            .collect::<Vec<_>>();
        for index in &collapsed {
            let start = self.hyperlinks[*index].run_start;
            let slot = self
                .extra_xml
                .iter()
                .filter(|(position, _)| *position == start)
                .count();
            self.extra_xml.push((start, Vec::new()));
            self.hyperlinks[*index].preserved_raw_before = Some(slot);
        }
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
        for (position, raw) in &self.extra_xml {
            if Self::raw_is_root_attributes(*position, raw) {
                continue;
            }
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
            marker.projected_run_index -= removed_projected_indices
                .iter()
                .filter(|index| **index < marker.projected_run_index)
                .count();
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
        for (position, raw) in &mut self.extra_xml {
            if Self::raw_is_root_attributes(*position, raw) {
                continue;
            }
            *position = boundary_map[(*position).min(old_run_count)];
        }
        for (position, raw_before, _) in &mut self.equations {
            let old_boundary = (*position).min(old_run_count);
            *position = boundary_map[old_boundary];
            *raw_before = raw_prefixes[old_boundary] + (*raw_before).min(raw_counts[old_boundary]);
        }
        for ruby in &mut self.rubies {
            ruby.base_start = boundary_map[ruby.base_start.min(old_run_count)];
            ruby.base_end = boundary_map[ruby.base_end.min(old_run_count)];
        }
        // An annotation over no base run is not written, so drop it.
        self.rubies.retain(|ruby| ruby.base_start < ruby.base_end);
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
            .filter_map(|(run, remove)| (!*remove).then_some(run))
            .collect();
        for old_index in collapsed {
            let Some(index) = hyperlink_map[old_index] else {
                continue;
            };
            let hyperlink = &self.hyperlinks[index];
            let (boundary, slot) = (
                hyperlink.run_start,
                hyperlink.preserved_raw_before.unwrap_or_default(),
            );
            let mut raw = Vec::new();
            let mut writer = Writer::new(&mut raw);
            let written = write_hyperlink_start(&mut writer, hyperlink)
                .and_then(|()| {
                    write_hyperlink_boundary(
                        &mut writer,
                        &self.revisions,
                        index,
                        hyperlink,
                        boundary,
                    )
                })
                .and_then(|()| {
                    writer
                        .write_event(Event::End(BytesEnd::new(hyperlink_qname(hyperlink))))
                        .map_err(Into::into)
                });
            if written.is_ok()
                && let Some((_, placeholder)) = self
                    .extra_xml
                    .iter_mut()
                    .filter(|(position, _)| *position == boundary)
                    .nth(slot)
            {
                *placeholder = raw;
            }
        }
        let _ = self.refresh_bookmark_projection();
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
        self.extra_xml.push((run_index, raw));
        let projected = self.refresh_bookmark_projection();
        if !projected {
            self.extra_xml.pop();
        }
        projected
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
        self.extra_xml.push((run_index, raw));
        let projected = self.refresh_bookmark_projection();
        if !projected {
            self.extra_xml.pop();
        }
        projected
    }

    /// Write range markers at accepted-view run boundaries, the run index
    /// space of [`Self::accepted_run_paths`].
    ///
    /// The markers go inside `w:sdtContent` when the range starts or ends
    /// between two runs of an inline content control, and around the whole
    /// control when the range covers it. A missing side means the range
    /// continues into another paragraph, so that side must not fall inside a
    /// control. A comment reference run follows the comment end marker. A
    /// range that cannot be written exactly is refused and the paragraph is
    /// unchanged.
    #[doc(hidden)]
    pub fn anchor_accepted_range(
        &mut self,
        start: Option<usize>,
        end: Option<usize>,
        anchor: RangeAnchor<'_>,
    ) -> std::result::Result<(), RangeAnchorError> {
        let (start_site, end_site) = self.accepted_range_sites(start, end)?;
        let mut paragraph = self.clone();
        // The end goes first. A start site never follows it, so the end and
        // its reference run leave the start site where it was.
        if let Some(site) = end_site {
            paragraph.insert_range_marker(site, anchor, false)?;
        }
        if let Some(site) = start_site {
            paragraph.insert_range_marker(site, anchor, true)?;
        }
        if !paragraph.refresh_bookmark_projection() {
            return Err(RangeAnchorError::Write);
        }
        *self = paragraph;
        Ok(())
    }

    /// Insert one run at a checked accepted-view boundary atomically.
    #[doc(hidden)]
    pub fn insert_accepted_run(
        &mut self,
        boundary: usize,
        run: CT_R,
    ) -> std::result::Result<(), RangeAnchorError> {
        let (site, _) = self.accepted_range_sites(Some(boundary), Some(boundary))?;
        let mut paragraph = self.clone();
        let inserted = match site.ok_or(RangeAnchorError::Write)? {
            MarkerSite::Control { controls, index } => paragraph
                .control_at_mut(&controls)
                .is_some_and(|control| control.insert_content(index, vec![SdtContent::Run(run)])),
            MarkerSite::Paragraph { boundary, position } => {
                paragraph.insert_run_after_boundary_items(boundary, position, run)
            }
        };
        if !inserted || !paragraph.refresh_bookmark_projection() {
            return Err(RangeAnchorError::Write);
        }
        *self = paragraph;
        Ok(())
    }

    /// Resolve where the start and end markers of a range go.
    fn accepted_range_sites(
        &self,
        start: Option<usize>,
        end: Option<usize>,
    ) -> std::result::Result<(Option<MarkerSite>, Option<MarkerSite>), RangeAnchorError> {
        let paths = self.accepted_run_paths();
        let run_count = paths.len();
        for boundary in [start, end].into_iter().flatten() {
            if boundary > run_count {
                return Err(RangeAnchorError::OutOfRange {
                    boundary,
                    run_count,
                });
            }
        }
        // The recursive owner of a run is its path without the final run step.
        let owner = |index: usize| {
            let segments = paths[index].segments();
            &segments[..segments.len() - 1]
        };
        // The deepest owner holding the runs on both sides of a boundary.
        let shared = |boundary: usize| match boundary.checked_sub(1) {
            Some(left) if boundary < run_count => {
                &owner(left)[..common_prefix_len(owner(left), owner(boundary))]
            }
            _ => &[] as &[AcceptedRunPathSegment],
        };
        let (owner_path, blamed) = match (start, end) {
            (Some(start), Some(end)) if start > end => {
                return Err(RangeAnchorError::Reversed { start, end });
            }
            (Some(start), Some(end)) if start == end => (shared(start), start),
            (Some(start), Some(end)) => {
                // The outermost owner that holds the first and the last run
                // of the range and reaches both boundaries.
                let (start_depth, end_depth) = (shared(start).len(), shared(end).len());
                let depth = start_depth.max(end_depth);
                let first = owner(start);
                if depth > common_prefix_len(first, owner(end - 1)) {
                    return Err(RangeAnchorError::CrossesControl { start, end });
                }
                let blamed = if start_depth >= end_depth { start } else { end };
                (&first[..depth], blamed)
            }
            (Some(boundary), None) | (None, Some(boundary)) => {
                let shared = shared(boundary);
                if shared
                    .iter()
                    .any(|segment| matches!(segment, AcceptedRunPathSegment::Revision(_)))
                {
                    return Err(RangeAnchorError::InsideRevision { boundary });
                } else if !shared.is_empty() {
                    return Err(RangeAnchorError::InsideControl { boundary });
                }
                (shared, boundary)
            }
            (None, None) => return Ok((None, None)),
        };
        let mut controls = Vec::with_capacity(owner_path.len());
        for segment in owner_path {
            match *segment {
                AcceptedRunPathSegment::ContentControl(index) => controls.push(index),
                AcceptedRunPathSegment::Revision(_) | AcceptedRunPathSegment::Run(_) => {
                    return Err(RangeAnchorError::InsideRevision { boundary: blamed });
                }
            }
        }

        let end_site = end
            .map(|boundary| self.marker_site(&paths, &controls, boundary, false))
            .transpose()?;
        let start_site = match (start, end) {
            (Some(start), Some(end)) if start == end => end_site.clone(),
            _ => start
                .map(|boundary| self.marker_site(&paths, &controls, boundary, true))
                .transpose()?,
        };
        Ok((start_site, end_site))
    }

    /// Place a marker just before the run after `boundary` for a start, or
    /// just after the run before it for an end, inside the owner `controls`.
    fn marker_site(
        &self,
        paths: &[AcceptedRunPath],
        controls: &[usize],
        boundary: usize,
        start: bool,
    ) -> std::result::Result<MarkerSite, RangeAnchorError> {
        if !controls.is_empty() {
            let path = if start {
                paths.get(boundary)
            } else {
                boundary.checked_sub(1).map(|index| &paths[index])
            }
            .ok_or(RangeAnchorError::Write)?;
            let control = self.control_at(controls).ok_or(RangeAnchorError::Write)?;
            let index = match path.segments()[controls.len()] {
                AcceptedRunPathSegment::Run(index)
                | AcceptedRunPathSegment::ContentControl(index) => index,
                AcceptedRunPathSegment::Revision(index) => {
                    control
                        .revisions()
                        .get(index)
                        .ok_or(RangeAnchorError::Write)?
                        .0
                }
            };
            return Ok(MarkerSite::Control {
                controls: controls.to_vec(),
                index: index + usize::from(!start),
            });
        }
        // A paragraph-level marker must fall between the top-level owners of
        // both neighbouring runs, which a revision inside a hyperlink can
        // prevent.
        let before_right = paths
            .get(boundary)
            .map(|path| self.paragraph_owner_positions(path).0);
        let after_left = boundary
            .checked_sub(1)
            .map(|index| self.paragraph_owner_positions(&paths[index]).1);
        let (Some(before_right), Some(after_left)) = (
            before_right.unwrap_or(Some((
                self.runs.len(),
                self.boundary_items(self.runs.len()).len(),
            ))),
            after_left.unwrap_or(Some((0, 0))),
        ) else {
            return Err(RangeAnchorError::HyperlinkRevision { boundary });
        };
        let (boundary, position) = if start { before_right } else { after_left };
        Ok(MarkerSite::Paragraph { boundary, position })
    }

    /// Positions just before and just after the paragraph-level owner of one
    /// accepted run, or `None` where no paragraph-level child can go.
    fn paragraph_owner_positions(
        &self,
        path: &AcceptedRunPath,
    ) -> (Option<ChildPosition>, Option<ChildPosition>) {
        let around = |boundary: usize, item: BoundaryItem| {
            let items = self.boundary_items(boundary);
            match items.iter().position(|candidate| *candidate == item) {
                Some(position) => (Some((boundary, position)), Some((boundary, position + 1))),
                None => (None, None),
            }
        };
        let raw_at = |boundary: usize, slot: usize| {
            self.extra_xml
                .iter()
                .enumerate()
                .filter(|(_, (position, _))| *position == boundary)
                .nth(slot)
                .map(|(index, _)| BoundaryItem::Raw(index))
        };
        match path.segments()[0] {
            AcceptedRunPathSegment::Run(index) => (
                Some((index, self.boundary_items(index).len())),
                Some((index + 1, 0)),
            ),
            AcceptedRunPathSegment::ContentControl(index) => self
                .content_controls
                .get(index)
                .map_or((None, None), |(boundary, _, _, _)| {
                    around(*boundary, BoundaryItem::Control(index))
                }),
            AcceptedRunPathSegment::Revision(index) => {
                let Some((boundary, slot, _)) = self.revisions.get(index) else {
                    return (None, None);
                };
                let boundary = *boundary;
                let Some(hyperlink) = hyperlink_revision_index(*slot) else {
                    return raw_at(boundary, *slot)
                        .map_or((None, None), |item| around(boundary, item));
                };
                match self.hyperlinks.get(hyperlink) {
                    Some(hyperlink) if hyperlink.preserved_raw_before.is_some() => {
                        raw_at(boundary, hyperlink.preserved_raw_before.unwrap_or_default())
                            .map_or((None, None), |item| around(boundary, item))
                    }
                    // A closing hyperlink writes its revision before the
                    // paragraph children at its end boundary, and an open
                    // one writes it after them.
                    Some(hyperlink)
                        if hyperlink.run_start < hyperlink.run_end
                            && boundary == hyperlink.run_end =>
                    {
                        (None, Some((boundary, 0)))
                    }
                    Some(hyperlink) if hyperlink.run_start < hyperlink.run_end => {
                        (Some((boundary, self.boundary_items(boundary).len())), None)
                    }
                    _ => (None, None),
                }
            }
        }
    }

    /// Reach the inline control that `controls` names.
    fn control_at(&self, controls: &[usize]) -> Option<&CT_Sdt> {
        let (first, rest) = controls.split_first()?;
        let mut control = &self.content_controls.get(*first)?.3;
        for index in rest {
            let SdtContent::ContentControl(nested) = control.content.get(*index)? else {
                return None;
            };
            control = nested;
        }
        Some(control)
    }

    fn control_at_mut(&mut self, controls: &[usize]) -> Option<&mut CT_Sdt> {
        let (first, rest) = controls.split_first()?;
        let mut control = &mut self.content_controls.get_mut(*first)?.3;
        for index in rest {
            let SdtContent::ContentControl(nested) = control.content.get_mut(*index)? else {
                return None;
            };
            control = nested;
        }
        Some(control)
    }

    fn insert_range_marker(
        &mut self,
        site: MarkerSite,
        anchor: RangeAnchor<'_>,
        start: bool,
    ) -> std::result::Result<(), RangeAnchorError> {
        let marker_xml = range_marker_xml(anchor, start).ok_or(RangeAnchorError::Write)?;
        let reference = match anchor {
            RangeAnchor::Comment(id) if !start => Some(CT_R {
                properties: None,
                content: vec![RunContent::CommentReference { id, raw_before: 0 }],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            }),
            RangeAnchor::Comment(_)
            | RangeAnchor::Bookmark { .. }
            | RangeAnchor::Permission { .. }
            | RangeAnchor::Proofing { .. } => None,
        };
        let inserted = match site {
            MarkerSite::Control { controls, index } => {
                let mut children = vec![SdtContent::RawXml(marker_xml)];
                children.extend(reference.map(SdtContent::Run));
                self.control_at_mut(&controls)
                    .is_some_and(|control| control.insert_content(index, children))
            }
            MarkerSite::Paragraph { boundary, position } => {
                let mut items = self.boundary_items(boundary);
                let item = match anchor {
                    RangeAnchor::Comment(id) => {
                        self.comment_ranges.push(if start {
                            CommentRangeMarker::Start {
                                id,
                                run_index: boundary,
                                raw_before: 0,
                                has_child_content: false,
                            }
                        } else {
                            CommentRangeMarker::End {
                                id,
                                run_index: boundary,
                                raw_before: 0,
                                has_child_content: false,
                            }
                        });
                        BoundaryItem::Marker(self.comment_ranges.len() - 1)
                    }
                    RangeAnchor::Bookmark { .. }
                    | RangeAnchor::Permission { .. }
                    | RangeAnchor::Proofing { .. } => {
                        self.extra_xml.push((boundary, marker_xml));
                        BoundaryItem::Raw(self.extra_xml.len() - 1)
                    }
                };
                items.insert(position.min(items.len()), item);
                self.place_boundary_items(&[(boundary, items)]);
                match reference {
                    Some(run) => self.insert_run_after_boundary_items(boundary, position + 1, run),
                    None => true,
                }
            }
        };
        if inserted {
            Ok(())
        } else {
            Err(RangeAnchorError::Write)
        }
    }

    /// List the paragraph-level children at one direct run boundary in the
    /// order that serialization writes them.
    fn boundary_items(&self, boundary: usize) -> Vec<BoundaryItem> {
        let raws = self
            .extra_xml
            .iter()
            .enumerate()
            .filter(|(_, (position, _))| *position == boundary)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let mut items = Vec::new();
        for raw_slot in 0..=raws.len() {
            let markers = self
                .comment_ranges
                .iter()
                .enumerate()
                .filter(|(_, marker)| {
                    marker.run_index() == boundary
                        && marker.raw_before().min(raws.len()) == raw_slot
                })
                .map(|(index, _)| index)
                .collect::<Vec<_>>();
            for marker_slot in 0..=markers.len() {
                items.extend(
                    self.content_controls
                        .iter()
                        .enumerate()
                        .filter(|(_, (at, raw_before, markers_before, _))| {
                            *at == boundary
                                && (*raw_before).min(raws.len()) == raw_slot
                                && (*markers_before).min(markers.len()) == marker_slot
                        })
                        .map(|(index, _)| BoundaryItem::Control(index)),
                );
                if let Some(marker) = markers.get(marker_slot) {
                    items.push(BoundaryItem::Marker(*marker));
                }
            }
            if let Some(raw) = raws.get(raw_slot) {
                items.push(BoundaryItem::Raw(*raw));
            }
        }
        items
    }

    /// Give each listed direct run boundary exactly its listed children in
    /// order, and move every projection that names a moved raw child.
    fn place_boundary_items(&mut self, boundaries: &[(usize, Vec<BoundaryItem>)]) {
        let mut raw_moves = Vec::new();
        let mut placed_raws = Vec::new();
        let mut placed_markers = Vec::new();
        for (boundary, items) in boundaries {
            let (mut raws, mut markers) = (0, 0);
            for item in items {
                match *item {
                    BoundaryItem::Control(index) => {
                        let control = &mut self.content_controls[index];
                        (control.0, control.1, control.2) = (*boundary, raws, markers);
                    }
                    BoundaryItem::Marker(index) => {
                        match &mut self.comment_ranges[index] {
                            CommentRangeMarker::Start {
                                run_index,
                                raw_before,
                                ..
                            }
                            | CommentRangeMarker::End {
                                run_index,
                                raw_before,
                                ..
                            } => (*run_index, *raw_before) = (*boundary, raws),
                        }
                        markers += 1;
                        placed_markers.push(index);
                    }
                    BoundaryItem::Raw(index) => {
                        let old_boundary = self.extra_xml[index].0;
                        let old_slot = self.extra_xml[..index]
                            .iter()
                            .filter(|(position, _)| *position == old_boundary)
                            .count();
                        raw_moves.push(((old_boundary, old_slot), (*boundary, raws)));
                        placed_raws.push(index);
                        raws += 1;
                        markers = 0;
                    }
                }
            }
        }
        // Serialization reads the raw children and markers of one boundary in
        // vector order, so the placed ones move to the end in list order.
        let mut extra_xml = std::mem::take(&mut self.extra_xml)
            .into_iter()
            .map(Some)
            .collect::<Vec<_>>();
        let placed = placed_raws
            .iter()
            .zip(&raw_moves)
            .filter_map(|(index, (_, (boundary, _)))| {
                extra_xml[*index].take().map(|(_, raw)| (*boundary, raw))
            })
            .collect::<Vec<_>>();
        self.extra_xml = extra_xml.into_iter().flatten().chain(placed).collect();
        let mut comment_ranges = std::mem::take(&mut self.comment_ranges)
            .into_iter()
            .map(Some)
            .collect::<Vec<_>>();
        let placed = placed_markers
            .iter()
            .filter_map(|index| comment_ranges[*index].take())
            .collect::<Vec<_>>();
        self.comment_ranges = comment_ranges.into_iter().flatten().chain(placed).collect();

        let moved = |at: usize, slot: usize| {
            raw_moves
                .iter()
                .find(|(old, _)| *old == (at, slot))
                .map(|(_, new)| *new)
        };
        let mut moved_hyperlinks = Vec::new();
        for (hyperlink_index, hyperlink) in self.hyperlinks.iter_mut().enumerate() {
            if hyperlink.run_start == hyperlink.run_end
                && let Some(slot) = hyperlink.preserved_raw_before
                && let Some((boundary, slot)) = moved(hyperlink.run_start, slot)
            {
                moved_hyperlinks.push((hyperlink_index, hyperlink.run_start, boundary));
                (hyperlink.run_start, hyperlink.run_end) = (boundary, boundary);
                hyperlink.preserved_raw_before = Some(slot);
            }
        }
        for (at, slot, _) in &mut self.revisions {
            if let Some(hyperlink) = hyperlink_revision_index(*slot) {
                if let Some((_, _, boundary)) = moved_hyperlinks
                    .iter()
                    .find(|(index, old, _)| *index == hyperlink && *old == *at)
                {
                    *at = *boundary;
                }
            } else if let Some(new) = moved(*at, *slot) {
                (*at, *slot) = new;
            }
        }
        for (at, slot, _) in &mut self.equations {
            if let Some(new) = moved(*at, *slot) {
                (*at, *slot) = new;
            }
        }
    }

    /// Insert a direct run at `boundary` after the first `kept` children
    /// there, so the remaining children follow the new run.
    fn insert_run_after_boundary_items(&mut self, boundary: usize, kept: usize, run: CT_R) -> bool {
        let items = self.boundary_items(boundary);
        // The insertion moves every child of the boundary after the new run.
        if !self.insert_unwrapped_run(boundary, run) {
            return false;
        }
        let (before, after) = items.split_at(kept.min(items.len()));
        if !before.is_empty() {
            self.place_boundary_items(&[
                (boundary, before.to_vec()),
                (boundary + 1, after.to_vec()),
            ]);
        }
        true
    }

    /// Remap facade-authored bookmark marker ids without reserializing any
    /// unrelated preserved child XML.
    #[doc(hidden)]
    pub fn remap_authored_bookmark_ids(
        &mut self,
        remap: &std::collections::HashMap<i32, i32>,
    ) -> bool {
        for (_, raw) in &mut self.extra_xml {
            remap_authored_bookmark_marker(raw, remap);
        }
        // Markers anchored inside an inline control are its content children.
        for (_, _, _, control) in &mut self.content_controls {
            control.remap_authored_bookmark_ids(remap);
        }
        self.refresh_bookmark_projection()
    }

    fn refresh_bookmark_projection(&mut self) -> bool {
        let mut word_prefixes = vec![
            "w".to_owned(),
            format!("\0r\0{R_NS}"),
            format!("\0mc\0{}", crate::namespace::MC_NS),
        ];
        // Inline controls keep the source prefix of their preserved runs,
        // which the isolated paragraph no longer declares.
        for prefix in self
            .bookmark_markers
            .iter()
            .flat_map(|marker| marker.word_prefixes.iter())
            .chain(
                self.content_controls
                    .iter()
                    .flat_map(|(_, _, _, control)| control.word_prefixes()),
            )
        {
            if !word_prefixes.contains(prefix) {
                word_prefixes.push(prefix.clone());
            }
        }
        let mut raw = Vec::new();
        if self.to_xml(&mut Writer::new(&mut raw)).is_err() {
            return false;
        }
        let mut reader = Reader::from_reader(raw.as_slice());
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start)) if matches_local_name(start.name().as_ref(), b"p") => {
                    let Ok(parsed) = CT_P::from_xml_with_prefixes(&mut reader, &word_prefixes)
                    else {
                        return false;
                    };
                    self.bookmark_markers = parsed.bookmark_markers;
                    return true;
                }
                Ok(Event::Eof) | Err(_) => return false,
                Ok(_) => {}
            }
            buffer.clear();
        }
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

    /// Parse a standalone paragraph fragment with declarations on its root.
    #[doc(hidden)]
    pub fn from_xml_fragment(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(element) => {
                    let prefixes = word_prefixes_at(&element, &[])?;
                    if !is_word_element(element.name().as_ref(), b"p", &prefixes) {
                        return Err(OxmlError::MissingElement("w:p root".to_owned()));
                    }
                    return Self::from_xml_with_prefixes_and_root(
                        &mut reader,
                        &prefixes,
                        Some(&element),
                    );
                }
                Event::Empty(element) => {
                    let prefixes = word_prefixes_at(&element, &[])?;
                    if is_word_element(element.name().as_ref(), b"p", &prefixes) {
                        return Self::from_empty_root(&element, &prefixes);
                    }
                    return Err(OxmlError::MissingElement("w:p root".to_owned()));
                }
                Event::Eof => return Err(OxmlError::MissingElement("w:p root".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Return exact source bytes for complex fields admitted by the typed paragraph parser.
    #[doc(hidden)]
    pub fn story_complex_field_sources(
        xml: &[u8],
        inherited_word_prefixes: &[String],
    ) -> Result<Vec<Vec<u8>>> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let paragraph = loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(element) => {
                    let prefixes = word_prefixes_at(&element, inherited_word_prefixes)?;
                    if !is_word_element(element.name().as_ref(), b"p", &prefixes) {
                        return Err(OxmlError::MissingElement("w:p root".to_owned()));
                    }
                    break Self::from_xml_with_prefixes_and_root(
                        &mut reader,
                        &prefixes,
                        Some(&element),
                    )?;
                }
                Event::Empty(element) => {
                    let prefixes = word_prefixes_at(&element, inherited_word_prefixes)?;
                    if is_word_element(element.name().as_ref(), b"p", &prefixes) {
                        break Self::from_empty_root(&element, &prefixes)?;
                    }
                    return Err(OxmlError::MissingElement("w:p root".to_owned()));
                }
                Event::Eof => return Err(OxmlError::MissingElement("w:p root".to_owned())),
                _ => {}
            }
            buffer.clear();
        };
        let mut sources = Vec::new();
        collect_story_complex_field_sources(&paragraph, &mut sources);
        Ok(sources)
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        Self::from_xml_with_prefixes_and_root(reader, word_prefixes, None)
    }

    pub(crate) fn from_xml_with_prefixes_and_root(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        root: Option<&BytesStart<'_>>,
    ) -> Result<Self> {
        reader.config_mut().trim_text(false);
        let mut properties = None;
        let mut runs = Vec::new();
        let mut run_sources = Vec::new();
        let mut comment_sources = Vec::new();
        let mut hyperlinks = Vec::new();
        let mut comment_ranges = Vec::new();
        let mut bookmark_markers = Vec::new();
        let mut extra_xml = Vec::new();
        let mut content_controls = Vec::new();
        let mut revisions = Vec::new();
        let mut rubies = Vec::new();
        let mut projected_run_count = 0usize;
        let mut tracked_run_count = 0usize;
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
                        projected_run_count += 1;
                        tracked_run_count += 1;
                    } else if matches_local_name(name.as_ref(), b"r")
                        && !element_prefix_has_binding(name.as_ref(), &prefixes)
                    {
                        let prefixes = prefixes_with_assumed_word_owner(name.as_ref(), &prefixes)?;
                        let raw = capture_element(reader, e)?;
                        runs.push(parse_run_raw(&raw, &prefixes)?);
                        run_sources.push(Some(raw));
                        projected_run_count += 1;
                        tracked_run_count += 1;
                    } else if is_word_element(name.as_ref(), b"sdt", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        if let Some(sdt) = CT_Sdt::from_inline_raw(&raw, &prefixes) {
                            let raw_before = raw_xml_count_at(&extra_xml, runs.len());
                            let markers_before = comment_ranges
                                .iter()
                                .filter(|marker: &&CommentRangeMarker| {
                                    marker.run_index() == runs.len()
                                        && marker.raw_before() == raw_before
                                })
                                .count();
                            let projection = accepted_control_bookmark_projection(&sdt);
                            append_nested_bookmark_projection(
                                &mut bookmark_markers,
                                &mut projected_run_count,
                                &mut tracked_run_count,
                                projection,
                                runs.len(),
                                raw_before,
                            );
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
                            // A hyperlink with no run stays raw, and its
                            // bookmarks still count where it sits.
                            append_nested_bookmark_projection(
                                &mut bookmark_markers,
                                &mut projected_run_count,
                                &mut tracked_run_count,
                                AcceptedBookmarkProjection {
                                    markers: parsed.bookmark_markers,
                                    projected_run_count: 0,
                                    tracked_run_count: 0,
                                },
                                run_start,
                                raw_xml_count_at(&extra_xml, run_start),
                            );
                            extra_xml.push((run_start, raw));
                        } else {
                            append_nested_bookmark_projection(
                                &mut bookmark_markers,
                                &mut projected_run_count,
                                &mut tracked_run_count,
                                AcceptedBookmarkProjection {
                                    markers: parsed.bookmark_markers.clone(),
                                    projected_run_count: parsed.projected_run_count,
                                    tracked_run_count: parsed.tracked_run_count,
                                },
                                run_start,
                                0,
                            );
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
                    } else if is_word_element(name.as_ref(), b"ruby", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        if let Some((mut ruby, base_runs)) = CT_Ruby::from_raw(&raw, &prefixes)? {
                            ruby.base_start = runs.len();
                            ruby.base_end = runs.len() + base_runs.len();
                            projected_run_count += base_runs.len();
                            tracked_run_count += base_runs.len();
                            run_sources.extend(base_runs.iter().map(|_| None));
                            runs.extend(base_runs);
                            rubies.push(ruby);
                        } else {
                            extra_xml.push((runs.len(), raw));
                        }
                    } else if is_word_element(name.as_ref(), b"fldSimple", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        if let Some(field) = parse_simple_field(&raw, &prefixes)? {
                            runs.push(field_run(field, None));
                            run_sources.push(Some(raw));
                            projected_run_count += 1;
                            tracked_run_count += 1;
                        } else {
                            extra_xml.push((runs.len(), raw));
                        }
                    } else if is_word_element(name.as_ref(), b"commentRangeStart", &prefixes)
                        || is_word_element(name.as_ref(), b"commentRangeEnd", &prefixes)
                    {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        let raw = capture_element(reader, e)?;
                        push_comment_marker(
                            &mut comment_ranges,
                            &extra_xml,
                            runs.len(),
                            id,
                            is_word_element(name.as_ref(), b"commentRangeStart", &prefixes),
                            raw_element_has_child_content(&raw),
                        );
                        comment_sources
                            .push((comment_ranges.last().expect("captured marker").clone(), raw));
                    } else if is_word_element(name.as_ref(), b"bookmarkStart", &prefixes)
                        || is_word_element(name.as_ref(), b"bookmarkEnd", &prefixes)
                    {
                        let start = is_word_element(name.as_ref(), b"bookmarkStart", &prefixes);
                        let id = optional_word_attribute(e, b"id", &prefixes)
                            .and_then(|value| value.parse().ok());
                        let bookmark_name = optional_word_attribute(e, b"name", &prefixes);
                        let raw = capture_element(reader, e)?;
                        bookmark_markers.push(BookmarkMarker {
                            start,
                            id,
                            name: bookmark_name,
                            run_index: runs.len(),
                            raw_before: raw_xml_count_at(&extra_xml, runs.len()),
                            projected_run_index: projected_run_count,
                            tracked_run_index: tracked_run_count,
                            word_prefixes: prefixes.clone(),
                            has_child_content: raw_element_has_child_content(&raw),
                        });
                        extra_xml.push((runs.len(), raw));
                    } else if is_word_element(name.as_ref(), b"ins", &prefixes)
                        || is_word_element(name.as_ref(), b"del", &prefixes)
                        || is_word_element(name.as_ref(), b"moveFrom", &prefixes)
                        || is_word_element(name.as_ref(), b"moveTo", &prefixes)
                    {
                        let raw_before = raw_xml_count_at(&extra_xml, runs.len());
                        let raw = capture_element(reader, e)?;
                        if let Some(revision) =
                            CT_Revision::from_raw_content(raw.clone(), &prefixes)
                        {
                            let projection = accepted_revision_bookmark_projection(&revision);
                            append_nested_bookmark_projection(
                                &mut bookmark_markers,
                                &mut projected_run_count,
                                &mut tracked_run_count,
                                projection,
                                runs.len(),
                                raw_before,
                            );
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
                    if is_word_element(name.as_ref(), b"pPr", &prefixes)
                        && e.attributes().next().is_none()
                    {
                        properties.get_or_insert_default();
                    } else if is_word_element(name.as_ref(), b"commentRangeStart", &prefixes)
                        || is_word_element(name.as_ref(), b"commentRangeEnd", &prefixes)
                    {
                        let id = required_word_i32_attribute(e, b"id", &prefixes)?;
                        push_comment_marker(
                            &mut comment_ranges,
                            &extra_xml,
                            runs.len(),
                            id,
                            is_word_element(name.as_ref(), b"commentRangeStart", &prefixes),
                            false,
                        );
                        comment_sources.push((
                            comment_ranges.last().expect("captured marker").clone(),
                            capture_empty_element(e)?,
                        ));
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
                            projected_run_index: projected_run_count,
                            tracked_run_index: tracked_run_count,
                            word_prefixes: prefixes.clone(),
                            has_child_content: false,
                        });
                        extra_xml.push((runs.len(), capture_empty_element(e)?));
                    } else if is_word_element(name.as_ref(), b"fldSimple", &prefixes) {
                        let raw = capture_empty_element(e)?;
                        let field = optional_word_attribute(e, b"instr", &prefixes).and_then(
                            |instruction| {
                                let instruction = parse_field_instruction(&instruction);
                                (!instruction.name.is_empty()).then(|| {
                                    let dirty = optional_word_attribute(e, b"dirty", &prefixes)
                                        .and_then(|value| parse_field_bool(&value));
                                    Field::parsed(
                                        instruction,
                                        String::new(),
                                        Vec::new(),
                                        dirty,
                                        FieldForm::Simple,
                                        raw.clone(),
                                        prefixes.clone(),
                                    )
                                })
                            },
                        );
                        if let Some(field) = field {
                            runs.push(field_run(field, None));
                            run_sources.push(None);
                            projected_run_count += 1;
                            tracked_run_count += 1;
                        } else {
                            extra_xml.push((runs.len(), raw));
                        }
                    } else if !matches_local_name(name.as_ref(), b"p") {
                        let raw_before = raw_xml_count_at(&extra_xml, runs.len());
                        let raw = capture_empty_element(e)?;
                        if (is_word_element(name.as_ref(), b"ins", &prefixes)
                            || is_word_element(name.as_ref(), b"del", &prefixes)
                            || is_word_element(name.as_ref(), b"moveFrom", &prefixes)
                            || is_word_element(name.as_ref(), b"moveTo", &prefixes))
                            && let Some(revision) =
                                CT_Revision::from_raw_content(raw.clone(), &prefixes)
                        {
                            let projection = accepted_revision_bookmark_projection(&revision);
                            append_nested_bookmark_projection(
                                &mut bookmark_markers,
                                &mut projected_run_count,
                                &mut tracked_run_count,
                                projection,
                                runs.len(),
                                raw_before,
                            );
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
            comment_sources: &comment_sources,
            bookmark_markers: &mut bookmark_markers,
            content_controls: &mut content_controls,
            revisions: &mut revisions,
            hyperlinks: &mut hyperlinks,
            rubies: &mut rubies,
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

        if let Some(record) = root
            .map(|root| capture_root_attribute_record(root, word_prefixes))
            .transpose()?
            .flatten()
        {
            extra_xml.push((ROOT_ATTRIBUTES_POSITION, record));
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
            rubies,
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
        if para_id.is_none()
            && self.properties.is_none()
            && self.runs.is_empty()
            && self.hyperlinks.is_empty()
            && self.comment_ranges.is_empty()
            && self.bookmark_markers.is_empty()
            && self.extra_xml.is_empty()
            && self.content_controls.is_empty()
            && self.revisions.is_empty()
            && self.equations.is_empty()
        {
            writer.write_event(Event::Empty(BytesStart::new("w:p")))?;
            return Ok(());
        }

        let mut start = BytesStart::new("w:p");
        for (position, raw) in &self.extra_xml {
            if Self::raw_is_root_attributes(*position, raw) {
                push_root_attribute_record(&mut start, raw, para_id.map(|_| (W14_NS, "paraId")))?;
            }
        }
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
        let mut current_ruby: Option<usize> = None;
        let mut written_field_owner = None;
        let mut written_span_end = 0;
        for (run_idx, run) in self.runs.iter().enumerate() {
            if run_idx < written_span_end {
                continue;
            }
            let in_hl = hyperlink_runs.get(&run_idx).copied();

            // A ruby annotation wraps its base runs, so it closes before any
            // paragraph boundary content that sits between two runs.
            if let Some(ruby_index) = current_ruby
                && self.rubies[ruby_index].base_end == run_idx
            {
                self.rubies[ruby_index].write_end(writer)?;
                current_ruby = None;
            }

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

            let foreign_word_namespace = current_hyperlink
                .and_then(|index| shadowed_word_namespace(&self.hyperlinks[index]));
            if let Some((span_end, raw, true)) = self.field_span_at(run_idx) {
                write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
                written_field_owner = None;
                written_span_end = span_end;
                continue;
            }

            // Fields are paragraph children even though the facade stores them
            // in synthetic runs for the existing inline traversal contract.
            if run.content.len() == 1
                && let RunContent::Field(field) = &run.content[0]
            {
                if field.span.is_some() {
                    write_run_field(
                        writer,
                        field,
                        foreign_word_namespace,
                        run.properties.as_ref(),
                    )?;
                    written_field_owner = None;
                    continue;
                }
                let owner_id = field.source_owner_id();
                if owner_id.is_some() && owner_id == written_field_owner {
                    continue;
                }
                let field = if let Some(owner_id) = owner_id {
                    let owner_fields = self.runs[run_idx..]
                        .iter()
                        .map_while(|run| match run.content.as_slice() {
                            [RunContent::Field(field)]
                                if field.source_owner_id() == Some(owner_id) =>
                            {
                                Some(field)
                            }
                            _ => None,
                        })
                        .collect::<Vec<_>>();
                    let owner_changed = owner_fields.iter().any(|field| !field.is_unchanged());
                    if owner_fields.len() > 1 && owner_changed {
                        return Err(OxmlError::InvalidValue(
                            "a field sharing one physical run was changed".to_owned(),
                        ));
                    }
                    written_field_owner = Some(owner_id);
                    field
                } else {
                    written_field_owner = None;
                    field
                };
                write_field_with_result_properties(
                    writer,
                    field,
                    current_hyperlink
                        .and_then(|index| shadowed_word_namespace(&self.hyperlinks[index])),
                    run.properties.as_ref(),
                )?;
                continue;
            }

            if run
                .content
                .iter()
                .any(|content| matches!(content, RunContent::Field(_)))
            {
                for field in run.content.iter().filter_map(|content| match content {
                    RunContent::Field(field) => Some(field),
                    _ => None,
                }) {
                    let Some(owner_id) = field.source_owner_id() else {
                        continue;
                    };
                    let owner_fields = self
                        .runs
                        .iter()
                        .flat_map(|candidate| &candidate.content)
                        .filter(|content| {
                            matches!(content, RunContent::Field(candidate) if candidate.source_owner_id() == Some(owner_id))
                        })
                        .count();
                    if owner_fields > 1 {
                        return Err(OxmlError::InvalidValue(
                            "a field sharing one physical run was changed".to_owned(),
                        ));
                    }
                }
                written_field_owner = None;
                write_mixed_field_run(
                    writer,
                    run,
                    current_hyperlink
                        .and_then(|index| shadowed_word_namespace(&self.hyperlinks[index])),
                )?;
                continue;
            }

            written_field_owner = None;

            if current_ruby.is_none()
                && let Some(ruby_index) = self
                    .rubies
                    .iter()
                    .position(|ruby| ruby.base_start == run_idx && ruby.base_end > run_idx)
            {
                self.rubies[ruby_index].write_start(writer)?;
                current_ruby = Some(ruby_index);
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

        if let Some(ruby_index) = current_ruby {
            self.rubies[ruby_index].write_end(writer)?;
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

    /// Each untouched field span's source bytes, with the bytes that write its
    /// split runs as separate physical runs.
    #[doc(hidden)]
    pub fn detached_field_spans(&self) -> Result<Vec<(&[u8], Vec<u8>)>> {
        let mut spans = Vec::new();
        let mut run_idx = 0;
        while run_idx < self.runs.len() {
            let Some((end, raw, true)) = self.field_span_at(run_idx) else {
                run_idx += 1;
                continue;
            };
            spans.push((raw, self.detached_field_span(run_idx..end)?));
            run_idx = end;
        }
        Ok(spans)
    }

    /// Append the field source replacements that start at run `run_idx`, and
    /// return the number of runs they cover.
    ///
    /// A field read from a run it shared with text is replaced together with
    /// that text, as one physical span, so every field of the span can change.
    #[doc(hidden)]
    pub fn field_source_replacements_at(
        &self,
        run_idx: usize,
        output: &mut Vec<(Vec<u8>, Vec<u8>)>,
    ) -> Result<usize> {
        if let Some((end, raw, untouched)) = self.field_span_at(run_idx) {
            let replacement = if untouched {
                raw.to_vec()
            } else {
                self.detached_field_span(run_idx..end)?
            };
            output.push((raw.to_vec(), replacement));
            return Ok(end - run_idx);
        }
        for content in self.runs.get(run_idx).map_or(&[][..], |run| &run.content) {
            let RunContent::Field(field) = content else {
                continue;
            };
            let (source, replacement) = field.source_replacement()?.ok_or_else(|| {
                OxmlError::MissingElement("the source of a parsed field".to_owned())
            })?;
            output.push((source.to_vec(), replacement));
        }
        Ok(1)
    }

    /// Write the runs of one field span as separate physical runs.
    fn detached_field_span(&self, runs: std::ops::Range<usize>) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        for run in &self.runs[runs] {
            match run.content.as_slice() {
                [RunContent::Field(field)] => {
                    write_run_field(&mut writer, field, None, run.properties.as_ref())?
                }
                _ => run.to_xml(&mut writer)?,
            }
        }
        Ok(writer.into_inner())
    }

    /// The field span whose runs start at `run_idx`, while they are still the
    /// runs the reader split it into: its end, source bytes, and whether its
    /// fields are unchanged too.
    ///
    /// The reader splits a run shared by a field and text into sibling runs.
    /// While those runs, their fields and the boundaries between them are as
    /// read, the span is written back as its original bytes.
    fn field_span_at(&self, run_idx: usize) -> Option<(usize, &[u8], bool)> {
        for lead in 0..=1 {
            let [RunContent::Field(field)] = self.runs.get(run_idx + lead)?.content.as_slice()
            else {
                continue;
            };
            let Some(span) = field.span.as_deref() else {
                continue;
            };
            let FieldSource::Parsed {
                raw_xml, owner_id, ..
            } = &field.source
            else {
                continue;
            };
            if span.position != lead || span.field_index != 0 {
                continue;
            }
            let end = run_idx + span.runs.len();
            let runs = self.runs.get(run_idx..end)?;
            let in_place = runs.iter().zip(&span.runs).all(|(run, shape)| match shape {
                Some(shape) => run == shape,
                None => matches!(
                    run.content.as_slice(),
                    [RunContent::Field(candidate)]
                        if candidate.source_owner_id() == Some(*owner_id)
                ),
            });
            let last = end - 1;
            let inside = |at: usize| at > run_idx && at <= last;
            let bounded = has_typed_boundary_inside(
                run_idx,
                last,
                &self.comment_ranges,
                &self.bookmark_markers,
                &self.content_controls,
                &self.revisions,
                &self.hyperlinks,
            ) || self.extra_xml.iter().any(|(at, _)| inside(*at))
                || self.equations.iter().any(|(at, _, _)| inside(*at))
                || self
                    .rubies
                    .iter()
                    .any(|ruby| ruby.base_start < end && ruby.base_end > run_idx)
                || self.hyperlinks.iter().any(|hyperlink| {
                    hyperlink
                        .extra_xml
                        .iter()
                        .any(|(boundary, _, _)| inside(hyperlink.run_start + *boundary))
                });
            if !in_place || bounded {
                return None;
            }
            let unchanged = runs.iter().all(|run| match run.content.as_slice() {
                [RunContent::Field(field)] => field.is_unchanged(),
                _ => true,
            });
            return Some((end, raw_xml.as_slice(), unchanged));
        }
        None
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
            sdt.collect_controls(SdtOwner::Inline, controls);
        }
    }
}

/// Write one field, leaving out the run content the reader split off it.
///
/// A field read from a run it shared with text or another field keeps that
/// whole span as its source. Once the span is edited the text is written by
/// its own runs, so this field writes only its own part of the source, with
/// its own cache and dirty updates.
fn write_run_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
    result_properties: Option<&CT_RPr>,
) -> Result<()> {
    let Some(span) = &field.span else {
        return write_field_with_result_properties(
            writer,
            field,
            foreign_word_namespace,
            result_properties,
        );
    };
    let mut own = field.clone();
    own.span = None;
    if let FieldSource::Parsed {
        raw_xml,
        word_prefixes,
        ..
    } = &mut own.source
    {
        *raw_xml = field_span_parts(raw_xml, word_prefixes)?
            .and_then(|parts| {
                parts
                    .into_iter()
                    .filter_map(|part| match part {
                        FieldSpanPart::Field(bytes) => Some(bytes),
                        FieldSpanPart::Outside(_) => None,
                    })
                    .nth(span.field_index)
            })
            .ok_or_else(|| {
                OxmlError::InvalidValue(
                    "a field could not be separated from its physical run".to_owned(),
                )
            })?;
    }
    write_field_with_result_properties(writer, &own, foreign_word_namespace, result_properties)
}

fn write_mixed_field_run<W: std::io::Write>(
    writer: &mut Writer<W>,
    run: &CT_R,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let mut segment_start = 0usize;
    for (field_index, content) in run.content.iter().enumerate() {
        let RunContent::Field(field) = content else {
            continue;
        };
        write_run_content_segment(
            writer,
            run,
            segment_start,
            field_index,
            segment_start == 0,
            false,
            foreign_word_namespace,
        )?;
        write_run_field(
            writer,
            field,
            foreign_word_namespace,
            run.properties.as_ref(),
        )?;
        segment_start = field_index + 1;
    }
    write_run_content_segment(
        writer,
        run,
        segment_start,
        run.content.len(),
        segment_start == 0,
        true,
        foreign_word_namespace,
    )
}

fn write_run_content_segment<W: std::io::Write>(
    writer: &mut Writer<W>,
    run: &CT_R,
    start: usize,
    end: usize,
    first: bool,
    last: bool,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let property_boundary = usize::from(run.properties.is_some());
    let lower = if first { 0 } else { property_boundary + start };
    let upper = property_boundary + end;
    let mut extra_xml = Vec::new();
    let mut extra_xml_positions = Vec::new();
    if run.extra_xml_positions.len() == run.extra_xml.len() {
        for (position, raw) in run.extra_xml_positions.iter().zip(&run.extra_xml) {
            let boundary = CT_R::raw_child_position(*position);
            if boundary < lower
                || boundary > upper
                || (*position & RAW_COMMENT_REFERENCE_FLAG != 0 && boundary == upper)
            {
                continue;
            }
            let mut mapped = *position;
            let mapped_boundary = if boundary < property_boundary {
                boundary
            } else {
                boundary.saturating_sub(start)
            };
            CT_R::set_raw_child_position(&mut mapped, mapped_boundary);
            extra_xml.push(raw.clone());
            extra_xml_positions.push(mapped);
        }
    } else if last {
        extra_xml = run.extra_xml.clone();
    }
    if start == end && extra_xml.is_empty() {
        return Ok(());
    }
    CT_R {
        properties: run.properties.clone(),
        content: run.content[start..end].to_vec(),
        extra_xml,
        extra_xml_positions,
        alt_drawings: Vec::new(),
    }
    .to_xml_with_word_override(writer, foreign_word_namespace)
}

fn collect_story_complex_field_sources(paragraph: &CT_P, sources: &mut Vec<Vec<u8>>) {
    for run in paragraph.runs() {
        for content in &run.content {
            let RunContent::Field(field) = content else {
                continue;
            };
            collect_story_complex_field_source(field, sources);
        }
    }
    for (_, _, revision) in &paragraph.revisions {
        collect_story_revision_complex_field_sources(revision, sources);
    }
}

fn collect_story_complex_field_source(field: &Field, sources: &mut Vec<Vec<u8>>) {
    if let FieldSource::Parsed {
        form: FieldForm::Complex,
        raw_xml,
        ..
    } = &field.source
    {
        sources.push(raw_xml.clone());
    }
    for nested in field.nested_fields_in_source_order() {
        collect_story_complex_field_source(nested, sources);
    }
}

fn collect_story_revision_complex_field_sources(
    revision: &CT_Revision,
    sources: &mut Vec<Vec<u8>>,
) {
    if let Some(paragraph) = revision.content_paragraph() {
        collect_story_complex_field_sources(paragraph, sources);
    }
    for (_, nested) in revision.nested_revisions() {
        collect_story_revision_complex_field_sources(nested, sources);
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
    write_field_with_result_properties(writer, field, foreign_word_namespace, None)
}

fn write_field_with_result_properties<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
    result_properties: Option<&CT_RPr>,
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
        original_locked,
        original_legacy_form,
        word_prefixes,
        ..
    } = &field.source
        && instruction_structure_eq(&field.instruction, original_instruction)
        && instruction_source_identity_eq(&field.instruction, original_instruction)
        && field.instruction.raw == original_instruction.raw
    {
        let cached_changed = field.cached_result != *original_cached_result
            && field.nested_cached_projection().as_ref() != Some(&field.cached_result);
        let legacy_form_changed = field.legacy_form != *original_legacy_form;
        let updated = match form {
            FieldForm::Simple => {
                update_simple_field_source(field, raw_xml, word_prefixes, cached_changed)?
            }
            FieldForm::Complex => {
                let mut updated = update_nested_field_sources(field, raw_xml, word_prefixes)?;
                if legacy_form_changed {
                    let form = field.legacy_form.as_ref().ok_or_else(|| {
                        OxmlError::MissingElement("typed legacy form data".to_owned())
                    })?;
                    updated = rewrite_legacy_form_source(&updated, word_prefixes, form)?;
                }
                if cached_changed || field.dirty != *original_dirty {
                    update_complex_field_source(field, &updated, word_prefixes, cached_changed)?
                } else {
                    updated
                }
            }
        };
        let updated = if let Some(runs) = field.typed_cache() {
            replace_typed_field_cache(
                &updated,
                *form,
                word_prefixes,
                runs,
                &field.typed_cached_comment_ranges,
            )?
        } else {
            updated
        };
        let updated = if field.locked != *original_locked {
            rewrite_field_lock(&updated, *form, word_prefixes, field.locked)?
        } else {
            updated
        };
        return write_raw_with_word_override(writer, &updated, foreign_word_namespace);
    }

    let mut field = field.clone();
    field.instruction = instruction_for_write(&field);

    let form = match &field.source {
        FieldSource::Parsed { form, .. } => *form,
        FieldSource::New { form, .. } => *form,
    };
    let form = if instruction_contains_nested(&field.instruction) {
        FieldForm::Complex
    } else {
        form
    };
    match form {
        FieldForm::Simple => {
            write_simple_field(writer, &field, foreign_word_namespace, result_properties)
        }
        FieldForm::Complex => {
            write_complex_field(writer, &field, foreign_word_namespace, result_properties)
        }
    }
}

fn rewrite_legacy_form_source(
    raw: &[u8],
    word_prefixes: &[String],
    form: &LegacyFormFieldData,
) -> Result<Vec<u8>> {
    let (kind_name, value_name, value) = match (&form.kind, &form.value) {
        (LegacyFormFieldKind::TextInput, LegacyFormFieldValue::Text(value)) => {
            ("textInput", "default", value.clone())
        }
        (LegacyFormFieldKind::CheckBox, LegacyFormFieldValue::Checked(value)) => (
            "checkBox",
            "checked",
            if *value { "1" } else { "0" }.to_owned(),
        ),
        (LegacyFormFieldKind::DropDownList, LegacyFormFieldValue::SelectedIndex(value)) => {
            ("ddList", "result", value.to_string())
        }
        _ => {
            return Err(OxmlError::InvalidValue(
                "legacy form value kind does not match the field kind".to_owned(),
            ));
        }
    };
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(raw.len() + value.len()));
    let mut buffer = Vec::new();
    let mut prefixes = word_prefixes.to_vec();
    let mut prefix_scopes = Vec::new();
    let mut stack = Vec::<Option<String>>::new();
    let mut begin_depth = None;
    let mut begin_seen = false;
    let mut ff_depth = None;
    let mut kind_depth = None;
    let mut kind_prefix = "w".to_owned();
    let mut wrote_value = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let local = word_local_name(&element, &local_prefixes);
                let depth = stack.len();
                if kind_depth.is_some_and(|kind| depth == kind + 1)
                    && !wrote_value
                    && form.kind == LegacyFormFieldKind::TextInput
                    && matches!(local.as_deref(), Some("maxLength" | "format"))
                {
                    write_legacy_form_value_element(
                        writer.get_mut(),
                        &kind_prefix,
                        value_name,
                        &value,
                    );
                    wrote_value = true;
                }
                if kind_depth.is_some_and(|kind| depth == kind + 1)
                    && !wrote_value
                    && form.kind == LegacyFormFieldKind::DropDownList
                    && local.as_deref() != Some(value_name)
                {
                    write_legacy_form_value_element(
                        writer.get_mut(),
                        &kind_prefix,
                        value_name,
                        &value,
                    );
                    wrote_value = true;
                }
                if kind_depth.is_some_and(|kind| depth == kind + 1)
                    && local.as_deref() == Some(value_name)
                {
                    write_patched_value_event(
                        writer.get_mut(),
                        &element,
                        &local_prefixes,
                        &value,
                        false,
                    )?;
                    wrote_value = true;
                } else {
                    writer.write_event(Event::Start(element.clone()))?;
                }
                stack.push(local.clone());
                prefix_scopes.push(std::mem::replace(&mut prefixes, local_prefixes));
                if local.as_deref() == Some("fldChar")
                    && !begin_seen
                    && optional_word_attribute(&element, b"fldCharType", &prefixes).as_deref()
                        == Some("begin")
                {
                    begin_depth = Some(depth);
                    begin_seen = true;
                } else if local.as_deref() == Some("ffData") && begin_depth == depth.checked_sub(1)
                {
                    ff_depth = Some(depth);
                } else if ff_depth.is_some_and(|ff| depth == ff + 1)
                    && local.as_deref() == Some(kind_name)
                {
                    kind_depth = Some(depth);
                    kind_prefix = qname_prefix(element.name().as_ref())
                        .unwrap_or("w")
                        .to_owned();
                }
            }
            Event::Empty(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let local = word_local_name(&element, &local_prefixes);
                let depth = stack.len();
                if ff_depth.is_some_and(|ff| depth == ff + 1) && local.as_deref() == Some(kind_name)
                {
                    kind_prefix = qname_prefix(element.name().as_ref())
                        .unwrap_or("w")
                        .to_owned();
                    writer.write_event(Event::Start(element.clone()))?;
                    write_legacy_form_value_element(
                        writer.get_mut(),
                        &kind_prefix,
                        value_name,
                        &value,
                    );
                    writer.write_event(Event::End(element.to_end()))?;
                    wrote_value = true;
                    buffer.clear();
                    continue;
                }
                if kind_depth.is_some_and(|kind| depth == kind + 1)
                    && !wrote_value
                    && form.kind == LegacyFormFieldKind::TextInput
                    && matches!(local.as_deref(), Some("maxLength" | "format"))
                {
                    write_legacy_form_value_element(
                        writer.get_mut(),
                        &kind_prefix,
                        value_name,
                        &value,
                    );
                    wrote_value = true;
                }
                if kind_depth.is_some_and(|kind| depth == kind + 1)
                    && !wrote_value
                    && form.kind == LegacyFormFieldKind::DropDownList
                    && local.as_deref() != Some(value_name)
                {
                    write_legacy_form_value_element(
                        writer.get_mut(),
                        &kind_prefix,
                        value_name,
                        &value,
                    );
                    wrote_value = true;
                }
                if kind_depth.is_some_and(|kind| depth == kind + 1)
                    && local.as_deref() == Some(value_name)
                {
                    write_patched_value_event(
                        writer.get_mut(),
                        &element,
                        &local_prefixes,
                        &value,
                        true,
                    )?;
                    wrote_value = true;
                } else {
                    writer.write_event(Event::Empty(element))?;
                }
            }
            Event::End(element) => {
                let depth = stack.len().saturating_sub(1);
                if kind_depth == Some(depth) {
                    if !wrote_value {
                        write_legacy_form_value_element(
                            writer.get_mut(),
                            &kind_prefix,
                            value_name,
                            &value,
                        );
                        wrote_value = true;
                    }
                    kind_depth = None;
                }
                if ff_depth == Some(depth) {
                    ff_depth = None;
                }
                if begin_depth == Some(depth) {
                    begin_depth = None;
                }
                writer.write_event(Event::End(element))?;
                stack.pop();
                prefixes = prefix_scopes
                    .pop()
                    .unwrap_or_else(|| word_prefixes.to_vec());
            }
            Event::Eof => break,
            event => writer.write_event(event)?,
        }
        buffer.clear();
    }
    if !wrote_value {
        return Err(OxmlError::MissingElement(format!(
            "legacy form {kind_name} value"
        )));
    }
    Ok(writer.into_inner())
}

fn word_local_name(element: &BytesStart<'_>, word_prefixes: &[String]) -> Option<String> {
    let name = element.name();
    let local = local_name(name.as_ref());
    is_word_element(name.as_ref(), local, word_prefixes)
        .then(|| String::from_utf8_lossy(local).into_owned())
}

fn qname_prefix(name: &[u8]) -> Option<&str> {
    let separator = name.iter().position(|byte| *byte == b':')?;
    std::str::from_utf8(&name[..separator]).ok()
}

fn write_legacy_form_value_element(output: &mut Vec<u8>, prefix: &str, name: &str, value: &str) {
    let value = quick_xml::escape::escape(value);
    output.extend_from_slice(
        format!(
            "<{prefix}:{name} xmlns:{prefix}=\"{}\" {prefix}:val=\"{value}\"/>",
            crate::namespace::W_NS,
        )
        .as_bytes(),
    );
}

fn write_patched_value_event(
    output: &mut Vec<u8>,
    element: &BytesStart<'_>,
    word_prefixes: &[String],
    value: &str,
    empty: bool,
) -> Result<()> {
    let mut tag_writer = Writer::new(Vec::new());
    tag_writer.write_event(if empty {
        Event::Empty(element.clone())
    } else {
        Event::Start(element.clone())
    })?;
    let mut tag = tag_writer.into_inner();
    let escaped = quick_xml::escape::escape(value);
    for attribute in element.attributes() {
        let attribute = attribute?;
        if is_word_attribute(attribute.key.as_ref(), b"val", word_prefixes) {
            let key = attribute.key.as_ref();
            if let Some(range) = lexical_attribute_value_range(&tag, key) {
                tag.splice(range, escaped.as_bytes().iter().copied());
                output.extend_from_slice(&tag);
                return Ok(());
            }
        }
    }
    let name = element.name();
    let preferred = qname_prefix(name.as_ref()).unwrap_or("w");
    let (prefix, declare) = usable_qualified_word_prefix(word_prefixes, &tag, preferred);
    if declare {
        let insert_at = if empty { tag.len() - 2 } else { tag.len() - 1 };
        tag.splice(
            insert_at..insert_at,
            format!(r#" xmlns:{prefix}="{}""#, crate::namespace::W_NS).bytes(),
        );
    }
    let insert_at = if empty { tag.len() - 2 } else { tag.len() - 1 };
    tag.splice(
        insert_at..insert_at,
        format!(" {prefix}:val=\"{escaped}\"").bytes(),
    );
    output.extend_from_slice(&tag);
    Ok(())
}

fn usable_qualified_word_prefix(
    word_prefixes: &[String],
    tag: &[u8],
    preferred: &str,
) -> (String, bool) {
    if !preferred.is_empty() && word_prefixes.iter().any(|prefix| prefix == preferred) {
        return (preferred.to_owned(), false);
    }
    if let Some(prefix) = word_prefixes
        .iter()
        .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
    {
        return (prefix.clone(), false);
    }
    for index in 0usize.. {
        let candidate = if index == 0 {
            "w".to_owned()
        } else {
            format!("w{index}")
        };
        if !tag
            .windows(candidate.len() + 6)
            .any(|window| window.starts_with(b"xmlns:") && window[6..] == *candidate.as_bytes())
        {
            return (candidate, true);
        }
    }
    unreachable!()
}

fn lexical_attribute_value_range(tag: &[u8], wanted: &[u8]) -> Option<std::ops::Range<usize>> {
    let mut cursor = tag.iter().position(|byte| byte.is_ascii_whitespace())?;
    while cursor < tag.len() {
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if matches!(tag.get(cursor), Some(b'>' | b'/')) {
            return None;
        }
        let name_start = cursor;
        while tag
            .get(cursor)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'=' | b'>' | b'/'))
        {
            cursor += 1;
        }
        let name = &tag[name_start..cursor];
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return None;
        }
        cursor += 1;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        let quote = *tag.get(cursor)?;
        if !matches!(quote, b'\"' | b'\'') {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while tag.get(cursor) != Some(&quote) {
            cursor += 1;
            if cursor >= tag.len() {
                return None;
            }
        }
        let value_end = cursor;
        cursor += 1;
        if name == wanted {
            return Some(value_start..value_end);
        }
    }
    None
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

    let complex_children = field
        .cached_fields
        .iter()
        .filter(|(_, child)| child.form() == FieldForm::Complex)
        .collect::<Vec<_>>();
    let simple_children = field
        .cached_fields
        .iter()
        .filter(|(_, child)| child.form() == FieldForm::Simple)
        .collect::<Vec<_>>();
    if field.typed_cache().is_none()
        && (scan.result_nested.len() != complex_children.len()
            || scan.result_simple.len() != simple_children.len())
    {
        return Err(OxmlError::MissingElement(
            "nested cached field spans in owning complex field".to_owned(),
        ));
    }
    for ((_, nested), span) in complex_children.into_iter().zip(scan.result_nested) {
        if nested.is_unchanged() {
            continue;
        }
        let replacement =
            rewrite_isolated_nested_field(nested, raw, &span, &scan.runs, word_prefixes)?;
        edits.push((span.start, span.end, replacement));
    }
    for ((_, child), (start, end)) in simple_children.into_iter().zip(scan.result_simple) {
        if child.is_unchanged() {
            continue;
        }
        let mut writer = Writer::new(Vec::new());
        write_field(&mut writer, child, None)?;
        edits.push((start, end, writer.into_inner()));
    }
    edits.sort_by_key(|(start, _, _)| *start);
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
    result_nested: Vec<NestedComplexSpan>,
    result_simple: Vec<(usize, usize)>,
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
    let mut result_nested = Vec::new();
    let mut result_simple = Vec::new();
    loop {
        let run_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let run_prefixes = word_prefixes_at(&element, word_prefixes)?;
                if is_word_element(element.name().as_ref(), b"fldSimple", &run_prefixes) {
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                    if open_fields.len() == 1 && open_fields[0].separated {
                        result_simple.push((run_start, reader.buffer_position() as usize));
                    }
                    buffer.clear();
                    continue;
                }
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
                                    &mut result_nested,
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
                                    &mut result_nested,
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
            Event::Eof => {
                return Ok(ComplexSourceScan {
                    runs,
                    nested,
                    result_nested,
                    result_simple,
                });
            }
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
    result_nested: &mut Vec<NestedComplexSpan>,
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
            if open_fields.len() == 1 {
                let target = if open_fields[0].separated {
                    result_nested
                } else {
                    nested
                };
                target.push(NestedComplexSpan {
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
        original_legacy_form,
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
    if field.legacy_form != *original_legacy_form {
        let form = field
            .legacy_form
            .as_ref()
            .ok_or_else(|| OxmlError::MissingElement("typed legacy form data".to_owned()))?;
        updated = rewrite_legacy_form_source(&updated, word_prefixes, form)?;
    }
    let cached_changed = field.cached_result != *original_cached_result
        && field.nested_cached_projection().as_ref() != Some(&field.cached_result);
    if cached_changed || field.dirty != *original_dirty {
        updated = update_complex_field_source(field, &updated, word_prefixes, cached_changed)?;
    }
    if let Some(runs) = field.typed_cache() {
        updated = replace_typed_field_cache(
            &updated,
            FieldForm::Complex,
            word_prefixes,
            runs,
            &field.typed_cached_comment_ranges,
        )?;
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

fn find_cached_field_source(
    raw: &[u8],
    source: &[u8],
    from: usize,
    inherited: &[String],
) -> Result<Option<usize>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut scopes = Vec::<Vec<String>>::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let prefixes =
                    word_prefixes_at(&element, scopes.last().map_or(inherited, Vec::as_slice))?;
                let field_source = is_word_element(element.name().as_ref(), b"r", &prefixes)
                    || is_word_element(element.name().as_ref(), b"fldSimple", &prefixes);
                if !scopes.is_empty()
                    && field_source
                    && start >= from
                    && raw.get(start..start.saturating_add(source.len())) == Some(source)
                {
                    return Ok(Some(start));
                }
                let wrapper = [
                    b"fldSimple".as_slice(),
                    b"hyperlink",
                    b"sdt",
                    b"sdtContent",
                    b"ins",
                    b"del",
                    b"moveFrom",
                    b"moveTo",
                    b"smartTag",
                    b"customXml",
                ]
                .iter()
                .any(|local| is_word_element(element.name().as_ref(), local, &prefixes));
                if wrapper {
                    scopes.push(prefixes);
                } else {
                    reader.read_to_end_into(element.name(), &mut Vec::new())?;
                }
            }
            Event::Empty(element) => {
                let prefixes =
                    word_prefixes_at(&element, scopes.last().map_or(inherited, Vec::as_slice))?;
                if !scopes.is_empty()
                    && is_word_element(element.name().as_ref(), b"fldSimple", &prefixes)
                    && start >= from
                    && raw.get(start..start.saturating_add(source.len())) == Some(source)
                {
                    return Ok(Some(start));
                }
            }
            Event::End(_) => {
                scopes.pop();
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn update_simple_field_source(
    field: &Field,
    raw: &[u8],
    word_prefixes: &[String],
    cached_changed: bool,
) -> Result<Vec<u8>> {
    let nested_source;
    let raw = if let Some(runs) = &field.simple_cached_runs {
        if field
            .cached_fields_in_source_order()
            .iter()
            .any(|child| matches!(child.source, FieldSource::New { .. }))
        {
            return Err(OxmlError::MissingElement(
                "replacement cached field has no original simple-cache source identity".to_owned(),
            ));
        }
        let mut paragraph = CT_P::new();
        paragraph.runs = runs.clone();
        let mut replacements = Vec::new();
        let mut boundary = 0;
        while boundary < paragraph.runs.len() {
            boundary += paragraph.field_source_replacements_at(boundary, &mut replacements)?;
        }
        let mut edits = Vec::new();
        let mut cursor = 0;
        for (source, replacement) in replacements {
            let start = find_cached_field_source(raw, &source, cursor, word_prefixes)?.ok_or_else(
                || OxmlError::MissingElement("nested simple-cache field source".to_owned()),
            )?;
            cursor = start + source.len();
            edits.push((start..cursor, replacement));
        }
        let mut updated = raw.to_vec();
        for (range, replacement) in edits.into_iter().rev() {
            updated.splice(range, replacement);
        }
        nested_source = updated;
        nested_source.as_slice()
    } else {
        raw
    };
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
                    let foreign_word_binding =
                        namespace_bindings(&prefixes)
                            .into_iter()
                            .any(|(prefix, namespace)| {
                                prefix == "w" && namespace != crate::namespace::W_NS
                            });
                    if foreign_word_binding {
                        // Keep the alias and original bindings on the cache owner.
                        // A canonical w declaration would change opaque descendants.
                        let mut replacement = element.clone().into_owned();
                        replacement.clear_attributes();
                        for attribute in element.attributes() {
                            let attribute = attribute?;
                            if !is_field_word_attribute(attribute.key.as_ref(), b"dirty", &prefixes)
                            {
                                replacement.push_attribute(attribute);
                            }
                        }
                        if let Some(dirty) = field.dirty {
                            let name = element.name();
                            let prefix = std::str::from_utf8(name.as_ref())?
                                .split_once(':')
                                .map(|(prefix, _)| prefix)
                                .or_else(|| {
                                    prefixes
                                        .iter()
                                        .find(|prefix| {
                                            !prefix.is_empty() && !prefix.starts_with('\0')
                                        })
                                        .map(String::as_str)
                                })
                                .ok_or_else(|| {
                                    OxmlError::InvalidValue(
                                        "field attribute requires a Word prefix".to_owned(),
                                    )
                                })?;
                            let attribute = format!("{prefix}:dirty");
                            replacement.push_attribute((
                                attribute.as_str(),
                                if dirty { "1" } else { "0" },
                            ));
                        }
                        writer.write_event(Event::Start(replacement))?;
                    } else {
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
                    }
                    contexts.push(FieldRewriteContext {
                        prefixes,
                        word_run: false,
                        canonical_end: (!foreign_word_binding).then_some("w:fldSimple"),
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
                            !prefixes.iter().any(|prefix| prefix == "w"),
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
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            !prefixes.iter().any(|prefix| prefix == "w"),
                        )?;
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
                            !prefixes.iter().any(|prefix| prefix == "w"),
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
                        write_field_result_content_with_binding(
                            &mut writer,
                            &field.cached_result,
                            !prefixes.iter().any(|prefix| prefix == "w"),
                        )?;
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
                if cached_changed && contexts.is_empty() && !wrote_result {
                    let foreign_word_namespace = namespace_bindings(&context.prefixes)
                        .into_iter()
                        .find(|(prefix, namespace)| {
                            prefix == "w" && namespace != crate::namespace::W_NS
                        })
                        .map(|(_, namespace)| namespace);
                    write_field_result_run(
                        &mut writer,
                        field,
                        foreign_word_namespace.as_deref(),
                        None,
                    )?;
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

/// Replace only the outer result, using its actual namespace and run boundaries.
/// A result sharing a control run is split into sibling runs with the original
/// control wrapper and properties retained on both preserved fragments.
enum CachedFieldSource {
    EmptySimple(BytesStart<'static>),
    Simple(std::ops::Range<usize>),
    Complex {
        result: std::ops::Range<usize>,
        first_open: std::ops::Range<usize>,
        first_close: std::ops::Range<usize>,
        last_open: std::ops::Range<usize>,
        last_close: std::ops::Range<usize>,
    },
}

fn cached_field_source(
    raw: &[u8],
    form: FieldForm,
    word_prefixes: &[String],
) -> Result<CachedFieldSource> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut contexts = Vec::<FieldRewriteContext>::new();
    let mut buffer = Vec::new();
    let mut physical_runs = Vec::<(usize, usize, usize, usize, usize)>::new();
    let mut current_run = None;
    let mut separate = None;
    let mut end = None;
    let mut depth = 0usize;
    let mut simple_open = None;
    let mut simple_close = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = raw.get(event_end.saturating_sub(2)..event_end) == Some(b"/>");
                let inherited = contexts
                    .last()
                    .map(|context| context.prefixes.as_slice())
                    .unwrap_or(word_prefixes);
                let prefixes = word_prefixes_at(&element, inherited)?;
                let word_run = contexts.is_empty()
                    && is_word_element(element.name().as_ref(), b"r", &prefixes);
                if word_run {
                    current_run = Some(physical_runs.len());
                    physical_runs.push((start, event_end, event_end, event_end, event_end));
                }
                if contexts.is_empty()
                    && is_word_element(element.name().as_ref(), b"fldSimple", &prefixes)
                {
                    if empty {
                        return Ok(CachedFieldSource::EmptySimple(element.into_owned()));
                    }
                    simple_open = Some(event_end);
                }
                if contexts.len() == 1
                    && contexts.last().is_some_and(|context| context.word_run)
                    && is_word_element(element.name().as_ref(), b"fldChar", &prefixes)
                {
                    let kind = optional_word_attribute(&element, b"fldCharType", &prefixes);
                    let marker_end = if empty {
                        event_end
                    } else {
                        reader.read_to_end_into(element.name(), &mut Vec::new())?;
                        reader.buffer_position() as usize
                    };
                    match kind.as_deref() {
                        Some("begin") => depth += 1,
                        Some("separate") if depth == 1 => {
                            separate = current_run.map(|run| (run, marker_end))
                        }
                        Some("end") => {
                            if depth == 1 {
                                end = current_run.map(|run| (run, start));
                            }
                            depth = depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    buffer.clear();
                    continue;
                }
                if empty {
                    if contexts.len() == 1
                        && contexts.last().is_some_and(|context| context.word_run)
                        && is_word_element(element.name().as_ref(), b"rPr", &prefixes)
                        && let Some(run) = current_run
                    {
                        physical_runs[run].2 = event_end;
                    }
                } else {
                    contexts.push(FieldRewriteContext {
                        prefixes,
                        word_run,
                        canonical_end: None,
                    });
                }
            }
            Event::End(element) => {
                let context = contexts.pop().ok_or_else(|| {
                    OxmlError::InvalidValue("unbalanced typed field source".into())
                })?;
                if contexts.len() == 1
                    && contexts.last().is_some_and(|context| context.word_run)
                    && is_word_element(element.name().as_ref(), b"rPr", &context.prefixes)
                    && let Some(run) = current_run
                {
                    physical_runs[run].2 = event_end;
                }
                if context.word_run {
                    let run = current_run.take().ok_or_else(|| {
                        OxmlError::MissingElement("typed cache physical run".into())
                    })?;
                    physical_runs[run].3 = start;
                    physical_runs[run].4 = event_end;
                }
                if contexts.is_empty()
                    && is_word_element(element.name().as_ref(), b"fldSimple", &context.prefixes)
                {
                    simple_close = Some(start);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    match form {
        FieldForm::Simple => {
            let (start, end) = simple_open
                .zip(simple_close)
                .ok_or_else(|| OxmlError::MissingElement("simple typed cache owner".into()))?;
            Ok(CachedFieldSource::Simple(start..end))
        }
        FieldForm::Complex => {
            let (first, start) = separate
                .ok_or_else(|| OxmlError::MissingElement("complex typed cache separator".into()))?;
            let (last, end) =
                end.ok_or_else(|| OxmlError::MissingElement("complex typed cache end".into()))?;
            if start > end || first > last {
                return Err(OxmlError::InvalidValue("inverted typed cache span".into()));
            }
            let first = physical_runs[first];
            let last = physical_runs[last];
            Ok(CachedFieldSource::Complex {
                result: start..end,
                first_open: first.0..first.2,
                first_close: first.3..first.4,
                last_open: last.0..last.2,
                last_close: last.3..last.4,
            })
        }
    }
}

fn validate_cached_comment_ranges(runs: &[CT_R], ranges: &[CommentRangeMarker]) -> Result<()> {
    let mut references = std::collections::BTreeMap::<i32, usize>::new();
    for (index, run) in runs.iter().enumerate() {
        for content in &run.content {
            if let RunContent::CommentReference { id, .. } = content
                && references.insert(*id, index).is_some()
            {
                return Err(OxmlError::InvalidValue(
                    "cached annotation reference IDs must be unique".into(),
                ));
            }
        }
    }
    let mut open = std::collections::BTreeMap::<i32, usize>::new();
    let mut closed = std::collections::BTreeSet::new();
    let mut last_position = 0;
    for marker in ranges {
        let (id, position, raw_before, has_child_content, start) = match marker {
            CommentRangeMarker::Start {
                id,
                run_index,
                raw_before,
                has_child_content,
            } => (*id, *run_index, *raw_before, *has_child_content, true),
            CommentRangeMarker::End {
                id,
                run_index,
                raw_before,
                has_child_content,
            } => (*id, *run_index, *raw_before, *has_child_content, false),
        };
        if position > runs.len()
            || position < last_position
            || raw_before != 0
            || has_child_content
            || !references.contains_key(&id)
        {
            return Err(OxmlError::InvalidValue(
                "cached annotation marker has no checked run boundary or reference".into(),
            ));
        }
        last_position = position;
        if start {
            if closed.contains(&id) || open.insert(id, position).is_some() {
                return Err(OxmlError::InvalidValue(
                    "cached annotation range has duplicate start".into(),
                ));
            }
        } else {
            if open.remove(&id).is_none() || !closed.insert(id) || position > references[&id] {
                return Err(OxmlError::InvalidValue(
                    "cached annotation range has no paired reference boundary".into(),
                ));
            }
        }
    }
    if !open.is_empty() || closed.len() != references.len() {
        return Err(OxmlError::InvalidValue(
            "cached annotation ranges and typed references disagree".into(),
        ));
    }
    Ok(())
}

fn write_typed_cached_runs<W: std::io::Write>(
    writer: &mut Writer<W>,
    runs: &[CT_R],
    ranges: &[CommentRangeMarker],
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    for index in 0..=runs.len() {
        for marker in ranges {
            let (id, position, name) = match marker {
                CommentRangeMarker::Start { id, run_index, .. } => {
                    (*id, *run_index, "w:commentRangeStart")
                }
                CommentRangeMarker::End { id, run_index, .. } => {
                    (*id, *run_index, "w:commentRangeEnd")
                }
            };
            if position != index {
                continue;
            }
            let mut element = BytesStart::new(name);
            element.push_attribute(("w:id", id.to_string().as_str()));
            if foreign_word_namespace.is_some() {
                element.push_attribute(("xmlns:w", crate::namespace::W_NS));
            }
            writer.write_event(Event::Empty(element))?;
        }
        if let Some(run) = runs.get(index) {
            write_mixed_field_run(writer, run, foreign_word_namespace)?;
        }
    }
    Ok(())
}

fn replace_typed_field_cache(
    raw: &[u8],
    form: FieldForm,
    word_prefixes: &[String],
    runs: &[CT_R],
    ranges: &[CommentRangeMarker],
) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    write_typed_cached_runs(&mut writer, runs, ranges, Some("typed-cache"))?;
    let payload = writer.into_inner();
    let mut output = Vec::new();
    match cached_field_source(raw, form, word_prefixes)? {
        CachedFieldSource::EmptySimple(element) => {
            let mut root = Writer::new(Vec::new());
            root.write_event(Event::Start(element.clone()))?;
            output.extend(root.into_inner());
            output.extend_from_slice(&payload);
            output.extend_from_slice(b"</");
            output.extend_from_slice(element.name().as_ref());
            output.extend_from_slice(b">");
        }
        CachedFieldSource::Simple(result) => {
            output.extend_from_slice(&raw[..result.start]);
            output.extend_from_slice(&payload);
            output.extend_from_slice(&raw[result.end..]);
        }
        CachedFieldSource::Complex {
            result,
            first_close,
            last_open,
            ..
        } => {
            output.extend_from_slice(&raw[..result.start]);
            output.extend_from_slice(&raw[first_close]);
            output.extend_from_slice(&payload);
            output.extend_from_slice(&raw[last_open]);
            output.extend_from_slice(&raw[result.end..]);
        }
    }
    Ok(output)
}

fn source_has_note_reference_name(raw: &[u8]) -> bool {
    [
        b"footnoteReference".as_slice(),
        b"endnoteReference".as_slice(),
        b"commentReference".as_slice(),
    ]
    .iter()
    .any(|name| raw.windows(name.len()).any(|window| window == *name))
}

fn parsed_complex_cache_runs(
    raw: &[u8],
    word_prefixes: &[String],
) -> Result<(Vec<CT_R>, Vec<CommentRangeMarker>)> {
    let CachedFieldSource::Complex {
        result,
        first_open,
        last_close,
        ..
    } = cached_field_source(raw, FieldForm::Complex, word_prefixes)?
    else {
        unreachable!("complex field cache");
    };
    let mut cache = b"<w:p>".to_vec();
    cache.extend_from_slice(&raw[first_open]);
    cache.extend_from_slice(&raw[result]);
    cache.extend_from_slice(&raw[last_close]);
    cache.extend_from_slice(b"</w:p>");
    let mut reader = Reader::from_reader(cache.as_slice());
    let mut buffer = Vec::new();
    let Event::Start(root) = reader.read_event_into(&mut buffer)? else {
        unreachable!("cache root");
    };
    let mut prefixes = word_prefixes.to_vec();
    if !prefixes.iter().any(|prefix| prefix == "w") {
        prefixes.push("w".into());
    }
    let paragraph = CT_P::from_xml_with_prefixes_and_root(&mut reader, &prefixes, Some(&root))?;
    Ok((
        paragraph.runs().into_iter().cloned().collect(),
        paragraph.comment_ranges,
    ))
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
    canonical_word_binding: bool,
) -> Result<()> {
    reader.read_to_end_into(source.name(), &mut Vec::new())?;
    if let Some(value) = value {
        write_field_result_content_inner(writer, value, Some(source), canonical_word_binding)?;
    } else {
        write_field_text_with_source(writer, "", Some(source), canonical_word_binding)?;
    }
    Ok(())
}

fn write_updated_empty_text_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    source: &BytesStart<'_>,
    value: Option<&str>,
    canonical_word_binding: bool,
) -> Result<()> {
    if let Some(value) = value {
        write_field_result_content_inner(writer, value, Some(source), canonical_word_binding)?;
    } else {
        write_field_text_with_source(writer, "", Some(source), canonical_word_binding)?;
    }
    Ok(())
}

fn write_field_result_content<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &str,
) -> Result<()> {
    write_field_result_content_with_binding(writer, value, false)
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
    // Keep the source Word alias when canonical w would shadow foreign
    // attribute bindings. Generated text without a source gets its own scope.
    let name = if canonical_word_binding {
        source
            .map(|source| {
                let name = source.name();
                let name = std::str::from_utf8(name.as_ref())?;
                Ok::<_, OxmlError>(
                    name.rsplit_once(':')
                        .map_or_else(|| "t".to_owned(), |(prefix, _)| format!("{prefix}:t")),
                )
            })
            .transpose()?
            .unwrap_or_else(|| "w:t".to_owned())
    } else {
        "w:t".to_owned()
    };
    let mut element = BytesStart::new(name.as_str());
    if let Some(source) = source {
        for attribute in source.attributes() {
            let attribute = attribute?;
            if attribute.key.as_ref() != b"xml:space" {
                element.push_attribute(attribute);
            }
        }
    }
    if canonical_word_binding && source.is_none() {
        element.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    if needs_space_preservation(value) {
        element.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(element))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new(name.as_str())))?;
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
            ..
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
    left.arguments.len() == right.arguments.len()
        && left.switches.len() == right.switches.len()
        && argument_identities_match
        && switch_identities_match
}

fn field_source_identity_eq(left: &Field, right: &Field) -> bool {
    match (&left.source, &right.source) {
        (
            FieldSource::Parsed {
                source_id: left_id, ..
            },
            FieldSource::Parsed {
                source_id: right_id,
                ..
            },
        ) => {
            left_id == right_id
                && instruction_source_identity_eq(&left.instruction, &right.instruction)
        }
        _ => false,
    }
}

fn canonical_instruction_text(instruction: &FieldInstruction) -> String {
    let mut text = instruction.name.clone();
    for argument in &instruction.arguments {
        if let FieldArgument::Text(value) = argument {
            push_canonical_field_token(&mut text, value, false);
        }
    }
    for switch in &instruction.switches {
        text.push(' ');
        text.push('\\');
        text.push_str(&switch.name);
        if let Some(FieldArgument::Text(value)) = &switch.argument {
            push_canonical_field_token(
                &mut text,
                value,
                !switch_takes_argument(&instruction.name, &switch.name),
            );
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
    result_properties: Option<&CT_RPr>,
) -> Result<()> {
    let mut element = BytesStart::new("w:fldSimple");
    if foreign_word_namespace.is_some() {
        element.push_attribute(("xmlns:w", crate::namespace::W_NS));
    }
    element.push_attribute(("w:instr", field.instruction.raw.as_str()));
    push_dirty_attribute(&mut element, field.dirty);
    push_lock_attribute(&mut element, field.locked);
    writer.write_event(Event::Start(element))?;
    write_field_result_run(writer, field, foreign_word_namespace, result_properties)?;
    writer.write_event(Event::End(BytesEnd::new("w:fldSimple")))?;
    Ok(())
}

fn write_complex_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
    result_properties: Option<&CT_RPr>,
) -> Result<()> {
    write_field_char_run(
        writer,
        "begin",
        field.dirty,
        field.locked,
        foreign_word_namespace,
    )?;
    let mut text = format!(" {}", field.instruction.name);
    for argument in &field.instruction.arguments {
        match argument {
            FieldArgument::Text(value) => push_canonical_field_token(&mut text, value, false),
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
                FieldArgument::Text(value) => push_canonical_field_token(
                    &mut text,
                    value,
                    !switch_takes_argument(&field.instruction.name, &switch.name),
                ),
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
    write_field_char_run(writer, "separate", None, None, foreign_word_namespace)?;
    write_field_result_run(writer, field, foreign_word_namespace, result_properties)?;
    write_field_char_run(writer, "end", None, None, foreign_word_namespace)?;
    Ok(())
}

fn write_nested_instruction_field<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    let mut field = field.clone();
    field.instruction = instruction_for_write(&field);
    write_complex_field(writer, &field, foreign_word_namespace, None)
}

fn push_canonical_field_token(output: &mut String, value: &str, force_quotes: bool) {
    output.push(' ');
    if force_quotes
        || value.is_empty()
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
    locked: Option<bool>,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    write_word_run_start(writer, foreign_word_namespace)?;
    let mut element = BytesStart::new("w:fldChar");
    element.push_attribute(("w:fldCharType", kind));
    push_dirty_attribute(&mut element, dirty);
    push_lock_attribute(&mut element, locked);
    writer.write_event(Event::Empty(element))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn write_field_result_run<W: std::io::Write>(
    writer: &mut Writer<W>,
    field: &Field,
    foreign_word_namespace: Option<&str>,
    properties: Option<&CT_RPr>,
) -> Result<()> {
    if let Some(runs) = field.typed_cache() {
        return write_typed_cached_runs(
            writer,
            runs,
            &field.typed_cached_comment_ranges,
            foreign_word_namespace,
        );
    }
    if let FieldSource::New {
        cached_runs: Some(runs),
        original_cached_result,
        ..
    } = &field.source
        && (field.cached_result == *original_cached_result
            || field.nested_cached_projection().as_ref() == Some(&field.cached_result))
    {
        for run in runs {
            if run
                .content
                .iter()
                .any(|content| matches!(content, RunContent::Field(_)))
            {
                write_mixed_field_run(writer, run, foreign_word_namespace)?;
            } else {
                run.to_xml_with_word_override(writer, foreign_word_namespace)?;
            }
        }
        return Ok(());
    }
    write_word_run_start(writer, foreign_word_namespace)?;
    let properties = properties.or_else(|| match &field.source {
        FieldSource::New {
            cached_runs: Some(runs),
            ..
        } => runs.first().and_then(|run| run.properties.as_ref()),
        _ => None,
    });
    if let Some(properties) = properties {
        properties.to_xml_with_word_override(writer, foreign_word_namespace)?;
    }
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

fn push_lock_attribute(element: &mut BytesStart<'_>, locked: Option<bool>) {
    if let Some(locked) = locked {
        element.push_attribute(("w:fldLock", if locked { "1" } else { "0" }));
    }
}

struct FieldLockOwner<'a> {
    start: usize,
    end: usize,
    element: BytesStart<'a>,
    prefixes: Vec<String>,
    empty: bool,
}

fn field_lock_owner<'a>(
    raw: &'a [u8],
    form: FieldForm,
    word_prefixes: &[String],
) -> Result<Option<FieldLockOwner<'a>>> {
    let mut reader = Reader::from_reader(raw);
    let mut contexts = Vec::<(Vec<String>, bool)>::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event()?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let prefixes = word_prefixes_at(
                    element,
                    contexts
                        .last()
                        .map(|context| context.0.as_slice())
                        .unwrap_or(word_prefixes),
                )?;
                let owner = match form {
                    FieldForm::Simple => {
                        contexts.is_empty()
                            && is_word_element(element.name().as_ref(), b"fldSimple", &prefixes)
                    }
                    FieldForm::Complex => {
                        contexts.len() == 1
                            && contexts[0].1
                            && is_word_element(element.name().as_ref(), b"fldChar", &prefixes)
                            && optional_word_attribute(element, b"fldCharType", &prefixes)
                                .as_deref()
                                == Some("begin")
                    }
                };
                if owner {
                    return Ok(Some(FieldLockOwner {
                        start,
                        end: reader.buffer_position() as usize,
                        element: element.clone(),
                        prefixes,
                        empty: matches!(event, Event::Empty(_)),
                    }));
                }
                if matches!(event, Event::Start(_)) {
                    let word_run = contexts.is_empty()
                        && is_word_element(element.name().as_ref(), b"r", &prefixes);
                    contexts.push((prefixes, word_run));
                }
            }
            Event::End(_) => {
                contexts.pop();
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}

fn field_source_lock(
    raw: &[u8],
    form: FieldForm,
    word_prefixes: &[String],
) -> Result<Option<bool>> {
    Ok(
        field_lock_owner(raw, form, word_prefixes)?.and_then(|owner| {
            optional_word_attribute(&owner.element, b"fldLock", &owner.prefixes)
                .and_then(|value| parse_field_bool(&value))
        }),
    )
}

fn rewrite_field_lock(
    raw: &[u8],
    form: FieldForm,
    word_prefixes: &[String],
    locked: Option<bool>,
) -> Result<Vec<u8>> {
    let owner = field_lock_owner(raw, form, word_prefixes)?
        .ok_or_else(|| OxmlError::MissingElement("field lock owner".into()))?;
    let owner_name = std::str::from_utf8(owner.element.name().as_ref())?.to_owned();
    let mut replacement = BytesStart::new(owner_name.as_str());
    for attribute in owner.element.attributes() {
        let attribute = attribute?;
        if !is_field_word_attribute(attribute.key.as_ref(), b"fldLock", &owner.prefixes) {
            replacement.push_attribute(attribute);
        }
    }
    if let Some(value) = locked {
        let prefix = owner_name
            .split_once(':')
            .map(|(prefix, _)| prefix)
            .unwrap_or("w");
        if !owner_name.contains(':') {
            replacement.push_attribute(("xmlns:w", crate::namespace::W_NS));
        }
        let name = format!("{prefix}:fldLock");
        replacement.push_attribute((name.as_str(), if value { "1" } else { "0" }));
    }
    let mut writer = Writer::new(raw[..owner.start].to_vec());
    writer.write_event(if owner.empty {
        Event::Empty(replacement)
    } else {
        Event::Start(replacement)
    })?;
    writer.get_mut().extend_from_slice(&raw[owner.end..]);
    Ok(writer.into_inner())
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
    has_child_content: bool,
) {
    let raw_before = raw_xml_count_at(extra_xml, run_index);
    let marker = if start {
        CommentRangeMarker::Start {
            id,
            run_index,
            raw_before,
            has_child_content,
        }
    } else {
        CommentRangeMarker::End {
            id,
            run_index,
            raw_before,
            has_child_content,
        }
    };
    markers.push(marker);
}

fn raw_element_has_child_content(raw: &[u8]) -> bool {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) => {
                if depth > 0 {
                    return true;
                }
                depth += 1;
            }
            Ok(Event::Empty(_)) if depth > 0 => return true,
            Ok(Event::Text(text))
                if depth > 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return true;
            }
            Ok(Event::CData(text))
                if depth > 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return true;
            }
            Ok(Event::GeneralRef(reference)) if depth > 0 => match reference.resolve_char_ref() {
                Ok(Some(character)) if !character.is_ascii_whitespace() => return true,
                Ok(Some(_)) => {}
                Ok(None) => return true,
                Err(_) => return false,
            },
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
        buffer.clear();
    }
}

fn raw_xml_count_at(extra_xml: &[(usize, Vec<u8>)], run_index: usize) -> usize {
    extra_xml
        .iter()
        .filter(|(position, raw)| *position == run_index && !is_xml_whitespace(raw))
        .count()
}

fn common_prefix_len(left: &[AcceptedRunPathSegment], right: &[AcceptedRunPathSegment]) -> usize {
    left.iter()
        .zip(right)
        .take_while(|(left, right)| left == right)
        .count()
}

/// Serialize one canonical comment or bookmark range marker.
fn range_marker_xml(anchor: RangeAnchor<'_>, start: bool) -> Option<Vec<u8>> {
    let tag = match anchor {
        RangeAnchor::Comment(_) if start => "w:commentRangeStart",
        RangeAnchor::Comment(_) => "w:commentRangeEnd",
        RangeAnchor::Bookmark { .. } if start => "w:bookmarkStart",
        RangeAnchor::Bookmark { .. } => "w:bookmarkEnd",
        RangeAnchor::Permission { .. } if start => "w:permStart",
        RangeAnchor::Permission { .. } => "w:permEnd",
        RangeAnchor::Proofing { .. } => "w:proofErr",
    };
    let mut value = itoa::Buffer::new();
    let mut element = BytesStart::new(tag);
    match anchor {
        RangeAnchor::Comment(id)
        | RangeAnchor::Bookmark { id, .. }
        | RangeAnchor::Permission { id, .. } => {
            element.push_attribute(("w:id", value.format(id)));
        }
        RangeAnchor::Proofing { kind } => {
            let value = match (kind, start) {
                ("spell", true) => "spellStart",
                ("spell", false) => "spellEnd",
                ("gram", true) => "gramStart",
                ("gram", false) => "gramEnd",
                _ => return None,
            };
            element.push_attribute(("w:type", value));
        }
    }
    if let RangeAnchor::Bookmark { name, .. } = anchor
        && start
    {
        element.push_attribute(("w:name", name));
    }
    if let RangeAnchor::Permission { editor, group, .. } = anchor
        && start
    {
        if let Some(editor) = editor {
            element.push_attribute(("w:ed", editor));
        }
        if let Some(group) = group {
            element.push_attribute(("w:edGrp", group));
        }
    }
    let mut raw = Vec::new();
    Writer::new(&mut raw)
        .write_event(Event::Empty(element))
        .ok()?;
    Some(raw)
}

/// Rewrite the id of one facade-authored bookmark marker through `remap`,
/// leaving any other raw child unchanged.
pub(crate) fn remap_authored_bookmark_marker(
    raw: &mut Vec<u8>,
    remap: &std::collections::HashMap<i32, i32>,
) {
    let Ok(text) = std::str::from_utf8(raw) else {
        return;
    };
    if !(text.starts_with("<w:bookmarkStart ") || text.starts_with("<w:bookmarkEnd ")) {
        return;
    }
    let Some(attribute) = text.find("w:id=\"") else {
        return;
    };
    let value_start = attribute + "w:id=\"".len();
    let Some(value_end) = text[value_start..].find('"').map(|end| value_start + end) else {
        return;
    };
    let Ok(old) = text[value_start..value_end].parse::<i32>() else {
        return;
    };
    let Some(updated) = remap.get(&old) else {
        return;
    };
    let mut replaced = Vec::with_capacity(raw.len());
    replaced.extend_from_slice(&raw[..value_start]);
    replaced.extend_from_slice(updated.to_string().as_bytes());
    replaced.extend_from_slice(&raw[value_end..]);
    *raw = replaced;
}

/// Whether a preserved run child is a field character or field code, which
/// a run keeps as raw XML when the rest of its complex field lies in other
/// runs.
fn raw_is_complex_field_part(raw: &[u8]) -> bool {
    raw_root_is_one_of(raw, &[b"fldChar", b"instrText", b"delInstrText"])
}

/// Whether the root element of a preserved child has one of `local_names`.
fn raw_root_is_one_of(raw: &[u8], local_names: &[&[u8]]) -> bool {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                return local_names
                    .iter()
                    .any(|local| matches_local_name(element.name().as_ref(), local));
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
        buffer.clear();
    }
}

/// Return the id of a preserved comment range marker, or `None` for any
/// other raw child.
pub(crate) fn raw_comment_marker_id(raw: &[u8], word_prefixes: &[String]) -> Option<i32> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(element) | Event::Empty(element) => {
                let prefixes = word_prefixes_at(&element, word_prefixes).ok()?;
                let name = element.name();
                if !is_word_element(name.as_ref(), b"commentRangeStart", &prefixes)
                    && !is_word_element(name.as_ref(), b"commentRangeEnd", &prefixes)
                {
                    return None;
                }
                return required_word_i32_attribute(&element, b"id", &prefixes).ok();
            }
            Event::Eof => return None,
            _ => {}
        }
        buffer.clear();
    }
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
    let projection =
        accepted_hyperlink_bookmark_projection(&runs, &revisions, &extra_xml, word_prefixes);
    Ok(ParsedHyperlinkChildren {
        runs,
        run_sources,
        revisions,
        extra_xml,
        bookmark_markers: projection.markers,
        projected_run_count: projection.projected_run_count,
        tracked_run_count: projection.tracked_run_count,
    })
}

fn bookmark_marker_projection_from_raw(
    raw: &[u8],
    word_prefixes: &[String],
    run_index: usize,
    raw_before: usize,
    projected_run_index: usize,
    tracked_run_index: usize,
) -> Option<BookmarkMarker> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).ok()?;
        let (element, has_child_content) = match event {
            Event::Start(element) => (element, raw_element_has_child_content(raw)),
            Event::Empty(element) => (element, false),
            Event::Text(text) if is_xml_whitespace(text.as_ref()) => {
                buffer.clear();
                continue;
            }
            Event::Eof => return None,
            _ => return None,
        };
        let prefixes = word_prefixes_at(&element, word_prefixes).ok()?;
        let start = is_word_element(element.name().as_ref(), b"bookmarkStart", &prefixes);
        if !start && !is_word_element(element.name().as_ref(), b"bookmarkEnd", &prefixes) {
            return None;
        }
        return Some(BookmarkMarker {
            start,
            id: optional_word_attribute(&element, b"id", &prefixes)
                .and_then(|value| value.parse().ok()),
            name: optional_word_attribute(&element, b"name", &prefixes),
            run_index,
            raw_before,
            projected_run_index,
            tracked_run_index,
            word_prefixes: prefixes,
            has_child_content,
        });
    }
}

fn append_projection(
    output: &mut AcceptedBookmarkProjection,
    projection: AcceptedBookmarkProjection,
    direct_run_offset: usize,
    raw_before: usize,
) {
    let projected_offset = output.projected_run_count;
    let tracked_offset = output.tracked_run_count;
    output
        .markers
        .extend(projection.markers.into_iter().map(|mut marker| {
            marker.run_index += direct_run_offset;
            marker.raw_before = raw_before;
            marker.projected_run_index += projected_offset;
            marker.tracked_run_index += tracked_offset;
            marker
        }));
    output.projected_run_count += projection.projected_run_count;
    output.tracked_run_count += projection.tracked_run_count;
}

fn append_nested_bookmark_projection(
    markers: &mut Vec<BookmarkMarker>,
    projected_run_count: &mut usize,
    tracked_run_count: &mut usize,
    projection: AcceptedBookmarkProjection,
    direct_run_offset: usize,
    raw_before: usize,
) {
    let projected_offset = *projected_run_count;
    let tracked_offset = *tracked_run_count;
    markers.extend(projection.markers.into_iter().map(|mut marker| {
        marker.run_index += direct_run_offset;
        marker.raw_before = raw_before;
        marker.projected_run_index += projected_offset;
        marker.tracked_run_index += tracked_offset;
        marker
    }));
    *projected_run_count += projection.projected_run_count;
    *tracked_run_count += projection.tracked_run_count;
}

fn accepted_paragraph_bookmark_projection(paragraph: &CT_P) -> AcceptedBookmarkProjection {
    let mut projection = AcceptedBookmarkProjection {
        markers: paragraph.bookmark_markers.clone(),
        projected_run_count: accepted_paragraph_runs(paragraph).len(),
        tracked_run_count: tracked_paragraph_runs(paragraph).len(),
    };
    projection
        .markers
        .sort_by_key(BookmarkMarker::projected_run_index);
    projection
}

fn accepted_revision_bookmark_projection(revision: &CT_Revision) -> AcceptedBookmarkProjection {
    if matches!(
        revision.kind(),
        RevisionKind::Insertion | RevisionKind::MoveTo
    ) && let Some(paragraph) = revision.content_paragraph()
    {
        return accepted_paragraph_bookmark_projection(paragraph);
    }
    let mut runs = Vec::new();
    append_tracked_revision_runs(revision, &mut runs);
    AcceptedBookmarkProjection {
        markers: Vec::new(),
        projected_run_count: 0,
        tracked_run_count: runs.len(),
    }
}

fn accepted_control_bookmark_projection(control: &CT_Sdt) -> AcceptedBookmarkProjection {
    let mut output = AcceptedBookmarkProjection::default();
    for boundary in 0..=control.content.len() {
        for (_, revision) in control.revisions().iter().filter(|(at, _)| *at == boundary) {
            let projection = accepted_revision_bookmark_projection(revision);
            append_projection(&mut output, projection, 0, 0);
        }
        let Some(content) = control.content.get(boundary) else {
            continue;
        };
        match content {
            SdtContent::Run(_) => {
                output.projected_run_count += 1;
                output.tracked_run_count += 1;
            }
            SdtContent::RawXml(raw) => {
                if let Some(marker) = bookmark_marker_projection_from_raw(
                    raw,
                    control.word_prefixes(),
                    0,
                    0,
                    output.projected_run_count,
                    output.tracked_run_count,
                ) {
                    output.markers.push(marker);
                }
            }
            SdtContent::ContentControl(control) => {
                let projection = accepted_control_bookmark_projection(control);
                append_projection(&mut output, projection, 0, 0);
            }
            SdtContent::Paragraph(paragraph) => {
                let projection = accepted_paragraph_bookmark_projection(paragraph);
                append_projection(&mut output, projection, 0, 0);
            }
            SdtContent::Table(table) => {
                append_control_table_bookmark_projection(&mut output, table)
            }
            SdtContent::Row(row) => append_control_row_bookmark_projection(&mut output, row),
            SdtContent::Cell(cell) => append_control_cell_bookmark_projection(&mut output, cell),
        }
    }
    output
}

fn append_control_table_bookmark_projection(
    output: &mut AcceptedBookmarkProjection,
    table: &CT_Tbl,
) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            let projection = accepted_control_bookmark_projection(control);
            append_projection(output, projection, 0, 0);
        }
        if let Some(row) = table.rows.get(boundary) {
            append_control_row_bookmark_projection(output, row);
        }
    }
}

fn append_control_row_bookmark_projection(output: &mut AcceptedBookmarkProjection, row: &CT_Row) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            let projection = accepted_control_bookmark_projection(control);
            append_projection(output, projection, 0, 0);
        }
        if let Some(cell) = row.cells.get(boundary) {
            append_control_cell_bookmark_projection(output, cell);
        }
    }
}

fn append_control_cell_bookmark_projection(output: &mut AcceptedBookmarkProjection, cell: &CT_Tc) {
    for content in &cell.content {
        let projection = match content {
            CellContent::Paragraph(paragraph) => accepted_paragraph_bookmark_projection(paragraph),
            CellContent::Table(table) => {
                append_control_table_bookmark_projection(output, table);
                continue;
            }
            CellContent::ContentControl(control) => accepted_control_bookmark_projection(control),
        };
        append_projection(output, projection, 0, 0);
    }
}

fn accepted_hyperlink_bookmark_projection(
    runs: &[CT_R],
    revisions: &[(usize, CT_Revision)],
    extra_xml: &[(usize, usize, Vec<u8>)],
    word_prefixes: &[String],
) -> AcceptedBookmarkProjection {
    let mut output = AcceptedBookmarkProjection::default();
    for boundary in 0..=runs.len() {
        let boundary_revisions = revisions
            .iter()
            .filter(|(at, _)| *at == boundary)
            .map(|(_, revision)| revision)
            .collect::<Vec<_>>();
        for revision_before in 0..=boundary_revisions.len() {
            for (_, _, raw) in extra_xml
                .iter()
                .filter(|(at, before, _)| *at == boundary && *before == revision_before)
            {
                if let Some(marker) = bookmark_marker_projection_from_raw(
                    raw,
                    word_prefixes,
                    boundary,
                    0,
                    output.projected_run_count,
                    output.tracked_run_count,
                ) {
                    output.markers.push(marker);
                }
            }
            if let Some(revision) = boundary_revisions.get(revision_before) {
                let projection = accepted_revision_bookmark_projection(revision);
                append_projection(&mut output, projection, boundary, 0);
            }
        }
        if boundary < runs.len() {
            output.projected_run_count += 1;
            output.tracked_run_count += 1;
        }
    }
    output
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
            } else if let Some((_, _, revision)) =
                boundary.revisions.iter().find(|(at, slot, _)| {
                    *at == run_index
                        && hyperlink_revision_index(*slot).is_none()
                        && *slot == raw_index
                })
            {
                revision.write_xml(writer)?;
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
            return Ok(attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                .parse()?);
        }
    }
    Err(OxmlError::MissingElement(format!(
        "{} attribute",
        String::from_utf8_lossy(local)
    )))
}

pub(crate) fn optional_word_attribute(
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

/// Whether the shared field lexer and parser admit a nonempty first token.
#[doc(hidden)]
pub fn story_field_instruction_has_name(instruction: &str) -> bool {
    !parse_field_instruction(instruction).name.is_empty()
}

#[derive(Debug)]
enum InstructionPart {
    Text(String),
    Nested(Field),
}

struct InstructionToken {
    argument: FieldArgument,
    quoted: bool,
    source_span: Option<std::ops::Range<usize>>,
    well_formed: bool,
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
                    source_span: None,
                    well_formed: true,
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
            ..
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
        let takes_argument = switch_takes_argument(&name, switch_name)
            || !switch_is_known_flag(&name, switch_name)
                && remaining.get(index + 1).is_some_and(|token| {
                    token.quoted || matches!(token.argument, FieldArgument::Nested(_))
                });
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

fn switch_is_known_flag(field_name: &str, switch_name: &str) -> bool {
    let flags: &[&str] = match field_name {
        "REF" => &["h", "n", "t", "w", "p"],
        "PAGEREF" => &["h", "p"],
        "SEQ" => &["n", "c", "h"],
        "STYLEREF" => &["l", "n", "t", "w", "p"],
        "TOC" => &["h", "z", "u", "w", "x"],
        "TC" => &["n"],
        "TOA" => &["b", "h", "p"],
        "INDEX" => &["r"],
        "XE" | "TA" => &["b", "i"],
        "CITATION" => &["n", "y", "t"],
        "INCLUDETEXT" => &["!"],
        "FILENAME" => &["p"],
        "MERGEFIELD" => &["m", "v"],
        _ => &[],
    };
    flags.contains(&switch_name.to_ascii_lowercase().as_str())
}

fn switch_takes_argument(field_name: &str, switch_name: &str) -> bool {
    let switch_name = switch_name.to_ascii_lowercase();
    if field_name == "INDEX" && switch_name == "r" {
        return false;
    }
    matches!(
        switch_name.as_str(),
        "*" | "#" | "@" | "r" | "s" | "d" | "b" | "f"
    ) || field_name == "INCLUDETEXT" && switch_name == "c"
        || field_name == "TOC" && matches!(switch_name.as_str(), "o" | "n" | "t" | "p")
        || field_name == "TC" && switch_name == "l"
        || field_name == "CITATION"
            && matches!(switch_name.as_str(), "l" | "m" | "p" | "v" | "f" | "s")
        || field_name == "BIBLIOGRAPHY" && matches!(switch_name.as_str(), "l" | "m" | "f")
        || field_name == "INDEX"
            && matches!(
                switch_name.as_str(),
                "z" | "f" | "h" | "e" | "l" | "g" | "k" | "s"
            )
        || field_name == "XE" && matches!(switch_name.as_str(), "f" | "r" | "t")
        || field_name == "TA" && matches!(switch_name.as_str(), "l" | "s" | "c" | "r")
        || field_name == "TOA" && matches!(switch_name.as_str(), "c" | "e" | "l" | "g")
        || field_name == "TOC" && matches!(switch_name.as_str(), "c" | "a")
        || matches!(field_name, "DISPLAYBARCODE" | "MERGEBARCODE")
            && matches!(switch_name.as_str(), "h" | "q" | "p" | "c")
}

fn lex_field_text(input: &str) -> Vec<InstructionToken> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut characters = input.char_indices().peekable();
    let mut quoted = false;
    let mut token_was_quoted = false;
    let mut token_start = None;
    let mut quote_closed = false;
    let mut well_formed = true;
    while let Some((offset, character)) = characters.next() {
        if !quoted && character.is_whitespace() {
            if !current.is_empty() || token_was_quoted {
                tokens.push(InstructionToken {
                    argument: FieldArgument::Text(std::mem::take(&mut current)),
                    quoted: token_was_quoted,
                    source_span: token_start.map(|start| start..offset),
                    well_formed,
                });
            }
            token_start = None;
            token_was_quoted = false;
            quote_closed = false;
            well_formed = true;
            continue;
        }
        token_start.get_or_insert(offset);
        if quoted {
            match character {
                '"' => {
                    quoted = false;
                    quote_closed = true;
                }
                '\\' if characters
                    .peek()
                    .is_some_and(|(_, next)| matches!(next, '"' | '\\')) =>
                {
                    current.push(characters.next().expect("peeked character exists").1);
                }
                _ => current.push(character),
            }
        } else {
            if quote_closed {
                well_formed = false;
            }
            match character {
                '"' => {
                    if !current.is_empty() {
                        well_formed = false;
                    }
                    quoted = true;
                    token_was_quoted = true;
                }
                _ => current.push(character),
            }
        }
    }
    if !current.is_empty() || token_was_quoted {
        tokens.push(InstructionToken {
            argument: FieldArgument::Text(current),
            quoted: token_was_quoted,
            source_span: token_start.map(|start| start..input.len()),
            well_formed: well_formed && !quoted,
        });
    }
    tokens
}

/// Replace only bibliography l, f and m token spans, retaining other raw spelling.
///
/// This checked edit uses the shared field lexer. Ambiguous source boundaries
/// return an error before a caller can publish any instruction changes.
#[doc(hidden)]
pub fn bibliography_option_switch_edits(
    raw: &str,
    replacement: &[FieldSwitch],
) -> Result<Vec<(std::ops::Range<usize>, String)>> {
    let tokens = lex_field_text(raw);
    if tokens.iter().any(|token| !token.well_formed) {
        return Err(OxmlError::InvalidValue(
            "ambiguous bibliography token boundary".into(),
        ));
    }
    let instruction = parse_field_instruction(raw);
    if instruction.name != "BIBLIOGRAPHY" || !instruction.arguments.is_empty() {
        return Err(OxmlError::InvalidValue(
            "invalid bibliography instruction shape".into(),
        ));
    }
    if replacement.iter().any(|switch| {
        !matches!(switch.name.as_str(), "l" | "f" | "m")
            || !matches!(switch.argument, Some(FieldArgument::Text(_)))
    }) {
        return Err(OxmlError::InvalidValue(
            "invalid owned bibliography replacement switch".into(),
        ));
    }
    let authored = FieldInstruction::new("BIBLIOGRAPHY", Vec::new(), replacement.to_vec())?;
    let mut ranges = Vec::new();
    let mut index = 1;
    while index < tokens.len() {
        let FieldArgument::Text(value) = &tokens[index].argument else {
            return Err(OxmlError::InvalidValue("nested bibliography token".into()));
        };
        let Some(name) = (!tokens[index].quoted)
            .then(|| value.strip_prefix('\\'))
            .flatten()
        else {
            return Err(OxmlError::InvalidValue(
                "unowned bibliography argument".into(),
            ));
        };
        let name = name.to_ascii_lowercase();
        let takes_argument = switch_takes_argument("BIBLIOGRAPHY", &name)
            || !switch_is_known_flag("BIBLIOGRAPHY", &name)
                && tokens.get(index + 1).is_some_and(|token| token.quoted);
        let has_argument = takes_argument && tokens.get(index + 1).is_some_and(|token| {
            token.quoted
                || !matches!(&token.argument, FieldArgument::Text(value) if value.starts_with('\\'))
        });
        let start = tokens[index]
            .source_span
            .as_ref()
            .ok_or_else(|| OxmlError::InvalidValue("missing bibliography token span".into()))?
            .start;
        let end_index = index + usize::from(has_argument);
        if matches!(name.as_str(), "l" | "m") && !has_argument {
            return Err(OxmlError::InvalidValue(
                "missing bibliography option operand".into(),
            ));
        }
        if matches!(name.as_str(), "l" | "f" | "m") {
            let end = tokens[end_index]
                .source_span
                .as_ref()
                .ok_or_else(|| OxmlError::InvalidValue("missing bibliography operand span".into()))?
                .end;
            ranges.push(start..end);
        }
        index = end_index + 1;
    }
    let owned = instruction
        .switches
        .iter()
        .filter(|switch| matches!(switch.name.as_str(), "l" | "f" | "m"))
        .cloned()
        .collect::<Vec<_>>();
    if owned == replacement {
        return Ok(Vec::new());
    }
    let mut edits = ranges
        .into_iter()
        .map(|range| (range, String::new()))
        .collect::<Vec<_>>();
    let appended = &authored.raw["BIBLIOGRAPHY".len()..];
    if !appended.is_empty() {
        edits.push((raw.len()..raw.len(), appended.to_owned()));
    }
    Ok(edits)
}

impl Default for CT_P {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {

    #[test]
    fn comment_literal_source_validates_namespaces_outside_selected_range() {
        let range = r#"<w:commentRangeStart w:id="4"/><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r><w:commentRangeEnd w:id="4"/>"#;
        for invalid in [
            "<q:bad/>",
            "<w:r q:bad=\"value\"/>",
            "<w:r q:bad=\"&invalid;\"/>",
        ] {
            for (before, inside, after) in [(invalid, "", ""), ("", invalid, ""), ("", "", invalid)]
            {
                let selected = range.replace("<w:tab/>", &format!("<w:tab/>{inside}"));
                let xml = format!(
                    r#"<w:p xmlns:w="{}">{before}{selected}{after}</w:p>"#,
                    crate::namespace::W_NS
                );
                assert!(
                    super::CT_P::comment_range_text_from_source(xml.as_bytes(), 4).is_err(),
                    "{xml}"
                );
            }
        }
        let xml = format!(
            r#"<w:p xmlns:w="{}" xmlns:q="urn:foreign"><q:box q:keep="value"/>{range}<q:box xmlns:q="urn:other"><q:t>IGNORED</q:t></q:box></w:p>"#,
            crate::namespace::W_NS
        );
        assert_eq!(
            super::CT_P::comment_range_text_from_source(xml.as_bytes(), 4)
                .unwrap()
                .as_deref(),
            Some("AB")
        );
    }
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
    fn bibliography_owned_token_edits_preserve_unowned_source_spelling() {
        let raw = "  bIbLiOgRaPhY\t\\L 1033  \\M \"a \\\"b\\\" \\\\ c\"\n\\Q \"keep \\\"exact\\\"\"  \\* MERGEFORMAT \\f 1036 \\m bare  ";
        let mut output = raw.to_owned();
        let edits = bibliography_option_switch_edits(
            raw,
            &[
                FieldSwitch {
                    name: "l".into(),
                    argument: Some(FieldArgument::Text("1031".into())),
                },
                FieldSwitch {
                    name: "m".into(),
                    argument: Some(FieldArgument::Text("new source".into())),
                },
            ],
        )
        .unwrap();
        for (range, replacement) in edits.into_iter().rev() {
            output.replace_range(range, &replacement);
        }
        assert_eq!(
            output,
            "  bIbLiOgRaPhY\t  \n\\Q \"keep \\\"exact\\\"\"  \\* MERGEFORMAT     \\l 1031 \\m \"new source\""
        );
        let parsed = parse_field_instruction(&output);
        assert!(parsed.arguments.is_empty());
        assert_eq!(
            parsed.switches.last().unwrap().argument,
            Some(FieldArgument::Text("new source".into()))
        );
        let raw = "BIBLIOGRAPHY \\f \\* MERGEFORMAT";
        let mut output = raw.to_owned();
        for (range, replacement) in bibliography_option_switch_edits(raw, &[])
            .unwrap()
            .into_iter()
            .rev()
        {
            output.replace_range(range, &replacement);
        }
        assert_eq!(output, "BIBLIOGRAPHY  \\* MERGEFORMAT");
    }

    #[test]
    fn bibliography_owned_token_edits_reject_ambiguous_boundaries() {
        for raw in [
            r#"BIBLIOGRAPHY \l "1033"adjacent"#,
            r#"BIBLIOGRAPHY \m "unterminated"#,
            r#"BIBLIOGRAPHY \m \l 1033"#,
            r#"BIBLIOGRAPHY \Q bare \l 1033"#,
            r#"BIBLIOGRAPHY unexpected"#,
        ] {
            assert!(bibliography_option_switch_edits(raw, &[]).is_err(), "{raw}");
        }
    }

    #[test]
    fn citation_and_bibliography_bare_operands_keep_opcode_specific_association() {
        let instruction = parse_field_instruction(
            r#"CITATION first \l 1033 \p 17 \v 3 \n \m second \f "before " \y \s " after" \t \l 1036"#,
        );
        assert_eq!(instruction.arguments, [FieldArgument::Text("first".into())]);
        assert_eq!(
            instruction
                .switches
                .iter()
                .map(|switch| switch.name.as_str())
                .collect::<Vec<_>>(),
            ["l", "p", "v", "n", "m", "f", "y", "s", "t", "l"]
        );
        assert!(
            instruction
                .switches
                .iter()
                .filter(|switch| matches!(switch.name.as_str(), "n" | "y" | "t"))
                .all(|switch| switch.argument.is_none())
        );
        let bibliography =
            parse_field_instruction(r#"BIBLIOGRAPHY \l 1033 \m first \m "second tag" \f"#);
        assert!(bibliography.arguments.is_empty());
        assert_eq!(bibliography.switches.len(), 4);
        assert!(bibliography.switches[3].argument.is_none());
        let unrelated = parse_field_instruction(r#"UNKNOWN first \l 1033"#);
        assert_eq!(unrelated.arguments.len(), 2);
        assert!(unrelated.switches[0].argument.is_none());
    }

    #[test]
    fn parse_simple_paragraph() {
        let p = parse_paragraph(r#"<w:r><w:t>Hello World</w:t></w:r>"#);
        assert_eq!(p.text(), "Hello World");
        assert_eq!(p.runs.len(), 1);
    }

    #[test]
    fn paragraph_root_attributes_reject_alias_duplicates_and_authored_id_wins() {
        let duplicate = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:b="http://schemas.microsoft.com/office/word/2010/wordml" a:paraId="11111111" b:paraId="22222222"/>"#;
        let error = CT_P::from_xml_fragment(duplicate).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("duplicate expanded root attribute")
        );

        let source = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:i="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:producer" i:paraId="11111111" i:textId="22222222" x:keep="yes"/>"#;
        let paragraph = CT_P::from_xml_fragment(source).unwrap();
        let mut output = Vec::new();
        paragraph
            .to_xml_with_para_id(&mut Writer::new(&mut output), Some("AAAAAAAA"))
            .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert_eq!(output.matches(":paraId=").count(), 1, "{output}");
        assert!(output.contains(r#"w14:paraId="AAAAAAAA""#), "{output}");
        assert!(output.contains(r#"i:textId="22222222""#), "{output}");
        assert!(output.contains(r#"x:keep="yes""#), "{output}");
    }

    #[test]
    fn a_retained_root_attribute_keeps_the_declaration_its_own_prefix_needs() {
        let source = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:i="http://schemas.microsoft.com/office/word/2010/wordml" i:paraId="11111111"/>"#;
        let paragraph = CT_P::from_xml_fragment(source).unwrap();
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"i:paraId="11111111""#), "{output}");
        assert!(
            output.contains(&format!(r#"xmlns:i="{W14_NS}""#)),
            "{output}"
        );
    }

    #[test]
    fn a_reopened_paragraph_identity_does_not_rebind_the_prefix_its_part_root_owns() {
        // A paragraph written with a bare `w14:paraId` and reopened used to
        // come back carrying its own `xmlns:w14`, so saving the reopened
        // document produced different bytes from the save it was read from.
        let w_ns = crate::namespace::W_NS;
        let source =
            format!(r#"<w:p xmlns:w="{w_ns}" xmlns:w14="{W14_NS}" w14:paraId="00000001"/>"#);
        let paragraph = CT_P::from_xml_fragment(source.as_bytes()).unwrap();
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"w14:paraId="00000001""#), "{output}");
        assert!(
            !output.contains("xmlns:w14"),
            "the paragraph rebound a prefix its part root already owns: {output}"
        );
    }

    #[test]
    fn dropping_w14_identities_keeps_every_other_retained_attribute() {
        // An alias prefix goes with the identities it bound, and a record
        // left with nothing else goes away.
        let w_ns = crate::namespace::W_NS;
        let source = format!(
            r#"<w:p xmlns:w="{w_ns}" xmlns:i="{W14_NS}" xmlns:x="urn:producer" i:paraId="11111111" w:rsidR="00A1B2C3" i:textId="22222222" x:keep="yes"/>"#
        );
        let mut paragraph = CT_P::from_xml_fragment(source.as_bytes()).unwrap();
        drop_w14_paragraph_identities(&mut paragraph.extra_xml).unwrap();
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(
            output.starts_with(r#"<w:p w:rsidR="00A1B2C3" x:keep="yes""#),
            "{output}"
        );
        assert!(output.contains(r#"xmlns:x="urn:producer""#), "{output}");
        assert!(!output.contains("Id="), "{output}");
        assert!(!output.contains("xmlns:i="), "{output}");

        let source = format!(
            r#"<w:p xmlns:w="{w_ns}" xmlns:w14="{W14_NS}" w14:paraId="11111111" w14:textId="22222222"/>"#
        );
        let mut paragraph = CT_P::from_xml_fragment(source.as_bytes()).unwrap();
        drop_w14_paragraph_identities(&mut paragraph.extra_xml).unwrap();
        assert!(paragraph.extra_xml.is_empty(), "{:?}", paragraph.extra_xml);
    }

    #[test]
    fn a_part_root_declares_w14_only_when_its_content_needs_it() {
        let w_ns = crate::namespace::W_NS;
        let part = |root: &str, body: &str| {
            format!(
                r#"<?xml version="1.0"?><w:document xmlns:w="{w_ns}"{root}><w:body>{body}</w:body></w:document>"#
            )
            .into_bytes()
        };
        let identity = r#"<w:p w14:paraId="00000001"/>"#;
        let mut unbound = part("", identity);
        declare_w14_on_part_root(&mut unbound).unwrap();
        assert_eq!(
            unbound,
            part(&format!(r#" xmlns:w14="{W14_NS}""#), identity)
        );
        crate::document::CT_Document::from_xml(&unbound).unwrap();

        for unchanged in [
            part(&format!(r#" xmlns:w14="{W14_NS}""#), identity),
            part(r#" xmlns:w14="urn:other""#, identity),
            part("", r#"<w:p><w:r><w:t>w14</w:t></w:r></w:p>"#),
        ] {
            let mut xml = unchanged.clone();
            declare_w14_on_part_root(&mut xml).unwrap();
            assert_eq!(xml, unchanged);
        }
    }

    #[test]
    fn a_root_with_only_namespace_declarations_records_nothing() {
        let start = BytesStart::from_content(
            r#"w:sectPr xmlns:sa="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sx="urn:section""#,
            "w:sectPr".len(),
        );
        let record =
            capture_root_attribute_record(&start, &["w".to_owned()]).expect("capture succeeds");
        assert!(record.is_none(), "{record:?}");
    }

    #[test]
    fn a_run_identity_resolves_the_word_prefix_the_default_scope_names() {
        // `CT_P::from_xml` names `w` as the Word prefix without binding it,
        // which is the scope a paragraph cut out of its part can be parsed in.
        // A run identity under it used to fail the paragraph as unbound.
        let paragraph = parse_paragraph(
            r#"<w:r w:rsidR="00A1B2C3" w:rsidRPr="00D4E5F6"><w:t>entry</w:t></w:r>"#,
        );
        assert_eq!(paragraph.text(), "entry");
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(
            output.contains(r#"<w:r w:rsidR="00A1B2C3" w:rsidRPr="00D4E5F6""#),
            "{output}"
        );
    }

    #[test]
    fn the_default_word_prefix_records_what_its_explicit_binding_records() {
        // The fallback must produce the record a part that binds `w` produces,
        // must never override an explicit binding, and must not reach a
        // prefix the scope does not name as Word.
        let start = BytesStart::from_content(r#"w:r w:rsidR="00A1B2C3""#, "w:r".len());
        let default_scope = capture_root_attribute_record(&start, &["w".to_owned()]).unwrap();
        let bound = capture_root_attribute_record(
            &start,
            &[format!("\0w\0{}", crate::namespace::W_NS), "w".to_owned()],
        )
        .unwrap();
        assert!(default_scope.is_some());
        assert_eq!(default_scope, bound);

        let rebound =
            capture_root_attribute_record(&start, &["w".to_owned(), "\0w\0urn:other".to_owned()])
                .unwrap()
                .unwrap();
        let rebound = String::from_utf8(rebound).unwrap();
        assert!(rebound.contains(r#"xmlns:w="urn:other""#), "{rebound}");

        let foreign = BytesStart::from_content(r#"w:r x:id="1""#, "w:r".len());
        let error = capture_root_attribute_record(&foreign, &["w".to_owned()]).unwrap_err();
        assert!(
            error.to_string().contains("prefix `x` is unbound"),
            "{error}"
        );
    }

    #[test]
    fn a_retained_record_declares_only_a_binding_the_written_element_lacks() {
        // The written element's own `w:` name needs the canonical `w` binding
        // in scope, so repeating it on every element with a `w:rsid*`
        // attribute only grew the part. An alias prefix and a foreign prefix
        // still need their declaration.
        let w_ns = crate::namespace::W_NS;
        let written = |attributes: &str, prefixes: &[String]| {
            let source = BytesStart::from_content(format!("w:p {attributes}"), "w:p".len());
            let record = capture_root_attribute_record(&source, prefixes)
                .unwrap()
                .unwrap();
            let mut start = BytesStart::new("w:p");
            push_root_attribute_record(&mut start, &record, None).unwrap();
            let mut output = Vec::new();
            Writer::new(&mut output)
                .write_event(Event::Empty(start))
                .unwrap();
            String::from_utf8(output).unwrap()
        };
        let word = format!("\0w\0{w_ns}");

        assert_eq!(
            written(r#"w:rsidR="00A1B2C3""#, std::slice::from_ref(&word)),
            r#"<w:p w:rsidR="00A1B2C3"/>"#
        );
        assert_eq!(
            written(
                r#"q:rsidR="00A1B2C3""#,
                &[word.clone(), format!("\0q\0{w_ns}")]
            ),
            format!(r#"<w:p q:rsidR="00A1B2C3" xmlns:q="{w_ns}"/>"#)
        );
        assert_eq!(
            written(
                r#"w:rsidR="00A1B2C3" x:keep="yes""#,
                &[word.clone(), "\0x\0urn:producer".to_owned()]
            ),
            r#"<w:p w:rsidR="00A1B2C3" x:keep="yes" xmlns:x="urn:producer"/>"#
        );
    }

    #[test]
    fn a_retained_element_does_not_rebind_a_prefix_its_part_root_declares() {
        let r_ns = crate::namespace::R_NS;
        let mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        let wp_ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        let xml = format!(
            r#"<w:document xmlns:w="{}" xmlns:r="{r_ns}" xmlns:mc="{mc_ns}" xmlns:wp="{wp_ns}"><w:body><w:p w:rsidR="00A1B2C3" r:stamp="relationship" mc:stamp="compatibility" wp:stamp="drawing" xmlns:x="urn:producer" x:stamp="paragraph"><w:r w:rsidRPr="00D4E5F6" xmlns:y="urn:new-binding" y:stamp="run"><w:t>Kept</w:t></w:r><x:opaque xmlns:x="urn:producer" x:flag="exact"><x:inside/></x:opaque></w:p><w:sectPr/></w:body></w:document>"#,
            crate::namespace::W_NS
        );
        let document = crate::document::CT_Document::from_xml(xml.as_bytes()).unwrap();
        let saved = String::from_utf8(document.to_xml().unwrap()).unwrap();
        assert_eq!(saved.matches("xmlns:w=").count(), 1, "{saved}");
        for prefix in ["r", "mc", "wp"] {
            assert_eq!(
                saved.matches(&format!("xmlns:{prefix}=")).count(),
                1,
                "{saved}"
            );
        }
        for attribute in [
            r#"w:rsidR="00A1B2C3""#,
            r#"w:rsidRPr="00D4E5F6""#,
            r#"r:stamp="relationship""#,
            r#"mc:stamp="compatibility""#,
            r#"wp:stamp="drawing""#,
            r#"x:stamp="paragraph""#,
            r#"y:stamp="run""#,
            r#"xmlns:x="urn:producer""#,
            r#"xmlns:y="urn:new-binding""#,
        ] {
            assert!(saved.contains(attribute), "missing {attribute}: {saved}");
        }
        let raw = r#"<x:opaque xmlns:x="urn:producer" x:flag="exact"><x:inside/></x:opaque>"#;
        assert!(saved.contains(raw), "raw child changed: {saved}");
        let reopened = crate::document::CT_Document::from_xml(saved.as_bytes()).unwrap();
        assert_eq!(reopened.to_xml().unwrap(), saved.as_bytes());
    }

    #[test]
    fn a_standalone_or_comment_paragraph_keeps_its_local_relationship_binding() {
        let source = format!(
            r#"<w:p xmlns:w="{}" xmlns:r="{R_NS}" r:stamp="kept"/>"#,
            crate::namespace::W_NS
        );
        let paragraph = CT_P::from_xml_fragment(source.as_bytes()).unwrap();
        let mut output = Vec::new();
        paragraph.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(&format!(r#"xmlns:r="{R_NS}""#)), "{output}");

        let comments = format!(
            r#"<w:comments xmlns:w="{}"><w:comment w:id="1" w:author="Ada"><w:p xmlns:r="{R_NS}" r:stamp="kept"/></w:comment></w:comments>"#,
            crate::namespace::W_NS
        );
        let comments = crate::comments::CT_Comments::from_xml(comments.as_bytes()).unwrap();
        let output = String::from_utf8(comments.to_xml().unwrap()).unwrap();
        assert!(output.contains(&format!(r#"xmlns:r="{R_NS}""#)), "{output}");
        assert!(output.contains(r#"r:stamp="kept""#), "{output}");
    }

    #[test]
    fn header_footer_and_note_roots_own_their_canonical_bindings() {
        let wp_ns = WP_NS;
        for root in ["hdr", "ftr"] {
            let source = format!(
                r#"<w:{root} xmlns:w="{}" xmlns:r="{R_NS}" xmlns:wp="{wp_ns}"><w:p r:stamp="link" wp:stamp="drawing"/></w:{root}>"#,
                crate::namespace::W_NS
            );
            let part = crate::header_footer::CT_HdrFtr::from_xml(source.as_bytes()).unwrap();
            let output = if root == "hdr" {
                part.to_xml_header().unwrap()
            } else {
                part.to_xml_footer().unwrap()
            };
            let output = String::from_utf8(output).unwrap();
            assert_eq!(output.matches("xmlns:r=").count(), 1, "{output}");
            assert_eq!(output.matches("xmlns:wp=").count(), 1, "{output}");
            assert!(output.contains(r#"r:stamp="link""#), "{output}");
            assert!(output.contains(r#"wp:stamp="drawing""#), "{output}");
        }

        for (root, note) in [("footnotes", "footnote"), ("endnotes", "endnote")] {
            let source = format!(
                r#"<w:{root} xmlns:w="{}" xmlns:r="{R_NS}"><w:{note} w:id="2"><w:p r:stamp="link"/></w:{note}></w:{root}>"#,
                crate::namespace::W_NS
            );
            let part = crate::footnotes::CT_Footnotes::from_xml(source.as_bytes()).unwrap();
            let output = if root == "footnotes" {
                part.to_xml_footnotes().unwrap()
            } else {
                part.to_xml_endnotes().unwrap()
            };
            let output = String::from_utf8(output).unwrap();
            assert_eq!(output.matches("xmlns:r=").count(), 1, "{output}");
            assert!(output.contains(r#"r:stamp="link""#), "{output}");
        }
    }

    #[test]
    fn a_root_prefix_shadow_with_another_uri_remains_local() {
        let source = format!(
            r#"<w:document xmlns:w="{}" xmlns:r="{R_NS}"><w:body><w:p xmlns:r="urn:producer" r:stamp="foreign"><w:r><w:t>Kept</w:t></w:r></w:p></w:body></w:document>"#,
            crate::namespace::W_NS
        );
        let document = crate::document::CT_Document::from_xml(source.as_bytes()).unwrap();
        let output = String::from_utf8(document.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"xmlns:r="urn:producer""#), "{output}");
        assert!(output.contains(r#"r:stamp="foreign""#), "{output}");
    }

    #[test]
    fn a_noncanonical_wp_root_does_not_claim_the_canonical_binding() {
        let source = format!(
            r#"<w:document xmlns:w="{}" xmlns:wp="urn:producer"><w:body><w:p xmlns:wp="{WP_NS}" wp:stamp="drawing"/></w:body></w:document>"#,
            crate::namespace::W_NS
        );
        let document = crate::document::CT_Document::from_xml(source.as_bytes()).unwrap();
        let output = String::from_utf8(document.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"xmlns:wp="urn:producer""#), "{output}");
        assert!(
            output.contains(&format!(r#"xmlns:wp="{WP_NS}""#)),
            "{output}"
        );
        assert!(output.contains(r#"wp:stamp="drawing""#), "{output}");
    }

    #[test]
    fn nested_root_binding_scopes_restore_on_unwind() {
        let before = ROOT_BINDINGS.with(Cell::get);
        let result = std::panic::catch_unwind(|| {
            let _outer = root_binding_scope(ROOT_R_BINDING);
            assert_eq!(ROOT_BINDINGS.with(Cell::get), ROOT_R_BINDING);
            let _inner = root_binding_scope(ROOT_MC_BINDING);
            assert_eq!(ROOT_BINDINGS.with(Cell::get), ROOT_MC_BINDING);
            panic!("scope restoration probe");
        });
        assert!(result.is_err());
        assert_eq!(ROOT_BINDINGS.with(Cell::get), before);
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
    fn empty_paragraph_properties_model_only_attribute_free_form() {
        let plain = parse_paragraph(r#"<w:pPr/><w:r><w:t>Text</w:t></w:r>"#);
        assert!(plain.properties.is_some());
        assert!(plain.extra_xml.is_empty());

        let attributed =
            parse_paragraph(r#"<w:pPr w:rsidR="00112233"/><w:r><w:t>Text</w:t></w:r>"#);
        assert!(attributed.properties.is_none());
        assert_eq!(attributed.extra_xml.len(), 1);
        assert!(String::from_utf8_lossy(&attributed.extra_xml[0].1).contains("00112233"));
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
            r#"<w:r xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><mc:AlternateContent><mc:Choice Requires="wps">"#,
            r#"<w:drawing><wp:anchor behindDoc="0">"#,
            r#"<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>"#,
            r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
            r#"<wp:extent cx="914400" cy="457200"/>"#,
            r#"<wp:docPr id="1" name="Shape"/>"#,
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

        // The fragment API supplies canonical w, but explicit foreign rebinding
        // must override that fallback at either the owner or paragraph boundary.
        for foreign in [
            src.replace(
                "<w:txbxContent>",
                "<w:txbxContent xmlns:w=\"urn:f279:foreign\">",
            ),
            src.replace(
                "<w:p><w:r><w:t>boxed",
                "<w:p xmlns:w=\"urn:f279:foreign\"><w:r><w:t>boxed",
            ),
        ] {
            let foreign_paragraph = parse_paragraph(&foreign);
            let shape = foreign_paragraph.runs[0].alt_drawings[0]
                .anchor
                .as_ref()
                .unwrap()
                .shape
                .as_ref()
                .unwrap();
            assert!(
                shape.text.is_empty(),
                "explicit foreign Word prefix is not admitted"
            );
            assert!(shape.text_body.as_ref().is_none_or(|body| {
                body.content
                    .iter()
                    .all(|content| !matches!(content, crate::document::BodyContent::Paragraph(_)))
            }));
            let mut preserved = Vec::new();
            foreign_paragraph
                .to_xml(&mut Writer::new(&mut preserved))
                .unwrap();
            assert!(
                String::from_utf8(preserved)
                    .unwrap()
                    .contains("urn:f279:foreign")
            );
        }

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
    fn parse_empty_fld_simple_preserves_the_field() {
        let paragraph = parse_paragraph(r#"<w:fldSimple w:instr=" PAGE "/>"#);
        let field = parsed_field(&paragraph, 0);

        assert_eq!(field.instruction.name, "PAGE");
        assert!(field.cached_result.is_empty());
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
    fn same_run_sibling_complex_fields_keep_both_typed_sources() {
        let xml = concat!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r>"#,
            r#"<w:fldChar w:fldCharType="begin"/><w:instrText>DATE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>first</w:t><w:fldChar w:fldCharType="end"/>"#,
            r#"<w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>second</w:t><w:fldChar w:fldCharType="end"/>"#,
            r#"</w:r><w:bookmarkStart w:id="1" w:name="after"/></w:p>"#,
        );
        let paragraph = CT_P::from_xml_fragment(xml.as_bytes()).unwrap();
        let fields = paragraph
            .runs()
            .into_iter()
            .flat_map(|run| run.content.iter())
            .filter_map(|content| match content {
                RunContent::Field(field) => Some(field),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].instruction.name, "DATE");
        assert_eq!(fields[0].cached_result, "first");
        assert_eq!(fields[1].instruction.name, "PAGE");
        assert_eq!(fields[1].cached_result, "second");
        assert_eq!(fields[0].source_owner_id(), fields[1].source_owner_id());
        let serialized = serialized_paragraph(&paragraph);
        assert_eq!(serialized.matches("<w:r>").count(), 1, "{serialized}");
        assert!(
            serialized.find("PAGE").unwrap() < serialized.find("bookmarkStart").unwrap(),
            "{serialized}"
        );
        assert_eq!(
            CT_P::story_complex_field_sources(xml.as_bytes(), &[])
                .unwrap()
                .len(),
            2
        );

        let mut changed = paragraph;
        let RunContent::Field(field) = &mut changed.runs[1].content[0] else {
            panic!("expected second sibling field")
        };
        field.cached_result = "changed".to_owned();
        let output = serialized_paragraph(&changed);
        assert_eq!(output.matches("<w:r>").count(), 2, "{output}");
        assert_eq!(output.matches("first").count(), 1, "{output}");
        assert_eq!(output.matches("changed").count(), 1, "{output}");
        assert!(!output.contains("second"), "{output}");
        assert!(
            output.find("PAGE").unwrap() < output.find("bookmarkStart").unwrap(),
            "{output}"
        );
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
    fn preserved_instruction_quote_balance_honors_escaped_quotes_and_backslashes() {
        for instruction in [
            r#"SEQ Figure \* "ARABIC""#,
            r#"INCLUDETEXT "\\server\share\file.docx" """#,
            r#"DOCPROPERTY "a\"b""#,
        ] {
            assert!(
                Field::new(instruction, "stored")
                    .effective_instruction()
                    .quotes_are_balanced(),
                "{instruction}"
            );
        }
        for instruction in [r#"SEQ Figure \* "ARABIC"#, r#"DOCPROPERTY "a\"b"#] {
            assert!(
                !Field::new(instruction, "stored")
                    .effective_instruction()
                    .quotes_are_balanced(),
                "{instruction}"
            );
        }
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
    fn extended_field_switches_use_the_shared_recursive_grammar() {
        let toc = parse_field_instruction_parts(vec![
            InstructionPart::Text("TOC \\b ".to_owned()),
            InstructionPart::Nested(Field::new("MERGEFIELD Scope", "nested")),
        ]);
        assert!(toc.arguments.is_empty());
        assert!(matches!(
            toc.switches.as_slice(),
            [FieldSwitch {
                name,
                argument: Some(FieldArgument::Nested(_)),
            }] if name == "b"
        ));

        let tc = parse_field_instruction(r#"TC "A\"B" \l 2"#);
        assert_eq!(tc.arguments, vec![FieldArgument::Text("A\"B".to_owned())]);
        assert_eq!(
            tc.switches,
            vec![FieldSwitch {
                name: "l".to_owned(),
                argument: Some(FieldArgument::Text("2".to_owned())),
            }]
        );

        let barcode = parse_field_instruction(r"DISPLAYBARCODE value QR \q 3 \p STD");
        assert_eq!(barcode.arguments.len(), 2);
        assert!(
            barcode
                .switches
                .iter()
                .all(|field_switch| field_switch.argument.is_some())
        );

        let source = concat!(
            r#"<q:fldSimple xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" q:instr="TOC \p &quot;:&quot;" data-token="kept">"#,
            r#"<q:r><q:t>cached</q:t><x:opaque xmlns:x="urn:producer"/></q:r>"#,
            r#"</q:fldSimple>"#,
        );
        let paragraph = parse_paragraph(source);
        let field = parsed_field(&paragraph, 0);
        assert!(matches!(
            field.instruction.switches.as_slice(),
            [FieldSwitch {
                name,
                argument: Some(FieldArgument::Text(separator)),
            }] if name == "p" && separator == ":"
        ));
        assert_eq!(
            serialized_paragraph(&paragraph),
            format!("<w:p>{source}</w:p>")
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
    fn accepted_inline_owners_project_bookmarks_once_in_exact_order() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:bookmarkStart w:id="1" w:name="direct"/>"#,
            r#"<w:hyperlink w:anchor="target"><w:r><w:t>h</w:t></w:r><w:bookmarkStart w:id="2" w:name="hyper"/></w:hyperlink>"#,
            r#"<w:ins w:id="3" w:author="Ada"><w:r><w:t>i</w:t></w:r><w:bookmarkStart w:id="4" w:name="revision"/></w:ins>"#,
            r#"<w:sdt><w:sdtContent><w:r><w:t>c</w:t></w:r><w:bookmarkStart w:id="5" w:name="control"/></w:sdtContent></w:sdt>"#,
            r#"<x:opaque xmlns:x="urn:opaque"><w:bookmarkStart w:id="6" w:name="opaque"/></x:opaque>"#,
            r#"<w:sdtContent><w:bookmarkStart w:id="7" w:name="malformed"/></w:sdtContent>"#,
            r#"<w:bookmarkEnd w:id="1"/>"#,
        ));

        assert_eq!(
            paragraph
                .bookmark_markers
                .iter()
                .filter_map(BookmarkMarker::name)
                .collect::<Vec<_>>(),
            ["direct", "hyper", "revision", "control"]
        );
        assert_eq!(
            paragraph
                .bookmark_markers
                .iter()
                .map(BookmarkMarker::projected_run_index)
                .collect::<Vec<_>>(),
            [0, 1, 2, 3, 3]
        );
        assert_eq!(paragraph.accepted_bookmark_runs().len(), 3);
        let output = serialized_paragraph(&paragraph);
        for name in [
            "direct",
            "hyper",
            "revision",
            "control",
            "opaque",
            "malformed",
        ] {
            assert_eq!(
                output.matches(&format!(r#"w:name="{name}""#)).count(),
                1,
                "{output}"
            );
        }
    }

    #[test]
    fn nested_control_local_word_alias_projects_bookmarks_after_reopen() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:sdt><w:sdtContent>"#,
            r#"<q:sdt xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:sdtContent>"#,
            r#"<q:bookmarkStart q:id="8" q:name="nested"/>"#,
            r#"<q:r><q:t>inside</q:t></q:r>"#,
            r#"<q:bookmarkEnd q:id="8"/>"#,
            r#"</q:sdtContent></q:sdt>"#,
            r#"</w:sdtContent></w:sdt>"#,
        ));

        assert_eq!(
            paragraph
                .bookmark_markers
                .iter()
                .filter_map(BookmarkMarker::name)
                .collect::<Vec<_>>(),
            ["nested"]
        );
        assert_eq!(
            paragraph
                .bookmark_markers
                .iter()
                .map(BookmarkMarker::projected_run_index)
                .collect::<Vec<_>>(),
            [0, 1]
        );
        let serialized = serialized_paragraph(&paragraph);
        assert!(
            serialized.contains(
                r#"xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#
            )
        );
        let reopened = parse_paragraph(
            serialized
                .strip_prefix("<w:p>")
                .and_then(|xml| xml.strip_suffix("</w:p>"))
                .expect("serialized paragraph wrapper"),
        );
        assert_eq!(reopened.bookmark_markers.len(), 2);
        assert_eq!(reopened.accepted_bookmark_runs()[0].text(), "inside");
    }

    #[test]
    fn complex_field_collapse_remaps_accepted_bookmark_boundaries() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText>AUTHOR</w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>cached</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:bookmarkStart w:id="9" w:name="afterField"/>"#,
            r#"<w:r><w:t>target</w:t></w:r>"#,
            r#"<w:bookmarkEnd w:id="9"/>"#,
        ));

        assert_eq!(paragraph.runs.len(), 2);
        assert_eq!(
            paragraph
                .bookmark_markers
                .iter()
                .map(BookmarkMarker::projected_run_index)
                .collect::<Vec<_>>(),
            [1, 2]
        );
        assert_eq!(paragraph.accepted_bookmark_runs().len(), 2);
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
            RunContent::FootnoteRef { id: 1, .. }
        ));
    }

    #[test]
    fn parse_endnote_reference() {
        let p = parse_paragraph(r#"<w:r><w:endnoteReference w:id="3"/></w:r>"#);
        assert_eq!(p.runs.len(), 1);
        assert!(matches!(
            p.runs[0].content[0],
            RunContent::EndnoteRef { id: 3, .. }
        ));
    }

    #[test]
    fn round_trip_footnote_reference() {
        let mut p = CT_P::new();
        p.add_run("Text before");
        p.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::FootnoteRef {
                id: 2,
                custom_mark: None,
            }],
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
            RunContent::FootnoteRef { id: 2, .. }
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
    fn encoded_comment_ids_decode_for_ranges_and_references() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:commentRangeStart w:id="&#55;"/>"#,
            r#"<w:r><w:t>commented</w:t></w:r>"#,
            r#"<w:commentRangeEnd w:id="&#x37;"/>"#,
            r#"<w:r><w:commentReference w:id="&#x37;"/></w:r>"#,
        ));

        assert!(matches!(
            paragraph.comment_ranges.as_slice(),
            [
                CommentRangeMarker::Start { id: 7, .. },
                CommentRangeMarker::End { id: 7, .. }
            ]
        ));
        assert!(matches!(
            paragraph.runs[1].content.as_slice(),
            [RunContent::CommentReference { id: 7, .. }]
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
    fn effective_instruction_text_matches_raw_structured_and_nested_edits() {
        let paragraph = parse_paragraph(
            r#"<w:fldSimple w:instr="  MERGEFIELD Name \* MERGEFORMAT  "><w:r><w:t>stored</w:t></w:r></w:fldSimple>"#,
        );
        let original_bytes = serialized_paragraph(&paragraph);
        let original = parsed_field(&paragraph, 0);
        assert_eq!(
            original.effective_instruction_text(),
            original.effective_instruction().raw
        );
        assert_eq!(serialized_paragraph(&paragraph), original_bytes);
        let mut raw = original.clone();
        raw.instruction.raw = "  AUTHOR  ".into();
        assert_eq!(
            raw.effective_instruction_text(),
            raw.effective_instruction().raw
        );
        assert_eq!(raw.effective_instruction_text(), "AUTHOR");
        let mut structured = original.clone();
        structured.instruction.name = "REF".into();
        structured.instruction.arguments = vec![FieldArgument::Text("Other Target".into())];
        assert_eq!(
            structured.effective_instruction_text(),
            structured.effective_instruction().raw
        );
        let mut nested = Field::new("IF", "stored");
        nested.instruction.arguments = vec![FieldArgument::Nested(Box::new(Field::new(
            "SEQ Figure",
            "old",
        )))];
        assert_eq!(
            nested.effective_instruction_text(),
            nested.effective_instruction().raw
        );
        let FieldArgument::Nested(child) = &mut nested.instruction.arguments[0] else {
            unreachable!()
        };
        child.instruction.name = "REF".into();
        child.instruction.arguments = vec![FieldArgument::Text("Target".into())];
        assert_eq!(
            nested.effective_instruction_text(),
            nested.effective_instruction().raw
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

    #[test]
    fn self_closing_simple_fields_count_in_bookmark_run_projections() {
        for field in [
            r#"<w:fldSimple w:instr=" PAGE "/>"#,
            r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"{field}<w:bookmarkStart w:id="1" w:name="after"/><w:r><w:t>text</w:t></w:r><w:bookmarkEnd w:id="1"/>"#
            ));

            assert_eq!(paragraph.runs.len(), 2, "{field}");
            let marker = &paragraph.bookmark_markers[0];
            assert_eq!(marker.run_index(), 1, "{field}");
            assert_eq!(marker.projected_run_index(), 1, "{field}");
            assert_eq!(marker.tracked_run_index(), 1, "{field}");
            assert_eq!(
                paragraph.bookmark_markers[1].projected_run_index(),
                2,
                "{field}"
            );
        }
    }

    #[test]
    fn mutation_refresh_retains_ancestor_only_bookmark_aliases() {
        let source = format!(
            r#"<q:p xmlns:q="{word}"><q:bookmarkStart q:id="7" q:name="target"/><q:r><q:t>target</q:t></q:r><q:bookmarkEnd q:id="7"/></q:p>"#,
            word = crate::namespace::W_NS,
        );
        let mut reader = Reader::from_str(&source);
        let mut buffer = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(start) => break start.into_owned(),
                Event::Eof => panic!("missing paragraph"),
                _ => buffer.clear(),
            }
        };
        let prefixes = word_prefixes_at(&start, &["w".to_owned()]).unwrap();
        let mut paragraph = CT_P::from_xml_with_prefixes(&mut reader, &prefixes).unwrap();

        assert_eq!(paragraph.bookmark_markers.len(), 2);
        assert!(paragraph.insert_unwrapped_run(1, CT_R::new("inserted")));
        assert_eq!(paragraph.bookmark_markers.len(), 2);
        assert_eq!(paragraph.bookmark_markers[0].projected_run_index(), 0);
        assert_eq!(paragraph.bookmark_markers[1].projected_run_index(), 2);
        let output = serialized_paragraph(&paragraph);
        assert!(output.contains(r#"<q:bookmarkStart q:id="7" q:name="target"/>"#));
        assert!(output.contains(r#"<q:bookmarkEnd q:id="7"/>"#));
    }

    #[test]
    fn inline_control_keeps_block_children_opaque_to_typed_projections() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:sdt><w:sdtContent>"#,
            r#"<w:p><w:bookmarkStart w:id="1" w:name="paragraph"/><w:r><w:t>paragraph</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
            r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            r#"<w:tr><w:tc><w:p><w:r><w:t>row</w:t></w:r></w:p></w:tc></w:tr>"#,
            r#"<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>"#,
            r#"<w:bookmarkStart w:id="2" w:name="inline"/><w:r><w:t>inline</w:t></w:r><w:bookmarkEnd w:id="2"/>"#,
            r#"</w:sdtContent></w:sdt>"#,
        ));

        assert_eq!(accepted_paragraph_runs(&paragraph).len(), 1);
        assert_eq!(tracked_paragraph_runs(&paragraph).len(), 1);
        assert_eq!(accepted_paragraph_runs(&paragraph)[0].text(), "inline");
        assert_eq!(
            paragraph
                .bookmark_markers
                .iter()
                .filter_map(BookmarkMarker::name)
                .collect::<Vec<_>>(),
            ["inline"]
        );
        let output = serialized_paragraph(&paragraph);
        for retained in ["paragraph", "table", "row", "cell", "inline"] {
            assert!(output.contains(retained), "{output}");
        }
    }

    #[test]
    fn legacy_form_data_accepts_aliases_and_rewrites_only_the_typed_value() {
        let mut paragraph = parse_paragraph(&format!(
            r#"<q:r xmlns:q="{word}" xmlns:x="urn:test"><q:fldChar q:fldCharType="begin"><q:ffData><q:name q:val="Choice"/><x:raw keep="yes"/><q:ddList><q:result q:val="0"/><q:default q:val="0"/><q:listEntry q:val="one"/><q:listEntry q:val="two"/></q:ddList></q:ffData></q:fldChar></q:r><q:r xmlns:q="{word}"><q:instrText> FORMDROPDOWN </q:instrText></q:r><q:r xmlns:q="{word}"><q:fldChar q:fldCharType="separate"/></q:r><q:r xmlns:q="{word}"><q:t>one</q:t></q:r><q:r xmlns:q="{word}"><q:fldChar q:fldCharType="end"/></q:r>"#,
            word = crate::namespace::W_NS,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        let form = field.legacy_form.as_ref().expect("typed form data");
        assert_eq!(form.kind, LegacyFormFieldKind::DropDownList);
        assert_eq!(form.value, LegacyFormFieldValue::SelectedIndex(0));
        assert_eq!(form.choices, ["one", "two"]);
        field
            .set_legacy_form_value(LegacyFormFieldValue::SelectedIndex(1))
            .unwrap();
        assert_eq!(field.cached_result, "two");
        let output = serialized_paragraph(&paragraph);
        assert!(output.contains(r#"<x:raw keep="yes"/>"#));
        assert!(output.contains(r#"q:result q:val="1""#));
        assert_eq!(output.matches("q:result").count(), 1);
    }

    #[test]
    fn sibling_namespace_shadow_does_not_hide_a_later_legacy_form_kind() {
        let paragraph = parse_paragraph(&format!(
            r#"<q:r xmlns:q="{word}" xmlns:x="urn:test"><q:fldChar q:fldCharType="begin"><q:ffData><x:raw xmlns:q="urn:producer"/><q:ddList><q:result q:val="0"/><q:listEntry q:val="one"/></q:ddList></q:ffData></q:fldChar></q:r><q:r xmlns:q="{word}"><q:instrText> FORMDROPDOWN </q:instrText></q:r><q:r xmlns:q="{word}"><q:fldChar q:fldCharType="end"/></q:r>"#,
            word = crate::namespace::W_NS,
        ));
        let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        assert_eq!(
            field.legacy_form.as_ref().map(|form| form.kind),
            Some(LegacyFormFieldKind::DropDownList)
        );
    }

    #[test]
    fn legacy_form_value_patch_targets_the_expanded_value_attribute_only() {
        let mut paragraph = parse_paragraph(&format!(
            r#"<q:r xmlns:q="{word}" xmlns:x="urn:test"><q:fldChar q:fldCharType="begin"><q:ffData><q:ddList><q:result x:a="q:val" x:b="keep" q:val="0"/><q:listEntry q:val="one"/><q:listEntry q:val="two"/></q:ddList></q:ffData></q:fldChar></q:r><q:r xmlns:q="{word}"><q:instrText> FORMDROPDOWN </q:instrText></q:r><q:r xmlns:q="{word}"><q:fldChar q:fldCharType="end"/></q:r>"#,
            word = crate::namespace::W_NS,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        field
            .set_legacy_form_value(LegacyFormFieldValue::SelectedIndex(1))
            .unwrap();
        let output = serialized_paragraph(&paragraph);
        assert!(output.contains(r#"x:a="q:val" x:b="keep" q:val="1""#));
    }

    #[test]
    fn default_namespace_missing_form_value_insertion_declares_its_prefix() {
        let mut paragraph = parse_paragraph(&format!(
            r#"<r xmlns="{word}" xmlns:q="{word}"><fldChar q:fldCharType="begin"><ffData><checkBox><sizeAuto/><default q:val="0"/></checkBox></ffData></fldChar></r><r xmlns="{word}"><instrText> FORMCHECKBOX </instrText></r><r xmlns="{word}" xmlns:q="{word}"><fldChar q:fldCharType="end"/></r>"#,
            word = crate::namespace::W_NS,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        field
            .set_legacy_form_value(LegacyFormFieldValue::Checked(true))
            .unwrap();
        let output = serialized_paragraph(&paragraph);
        assert!(output.contains(&format!(
            r#"<w:checked xmlns:w="{}" w:val="1"/>"#,
            crate::namespace::W_NS
        )));
        let scoped = output.replacen(
            "<w:p>",
            &format!(r#"<w:p xmlns="{0}" xmlns:q="{0}">"#, crate::namespace::W_NS),
            1,
        );
        let mut reader = Reader::from_str(&scoped);
        let mut buffer = Vec::new();
        let Event::Start(start) = reader.read_event_into(&mut buffer).unwrap() else {
            panic!("reparsed paragraph start")
        };
        let prefixes = word_prefixes_at(&start, &["w".to_owned()]).unwrap();
        let reparsed = CT_P::from_xml_with_prefixes(&mut reader, &prefixes).unwrap();
        let RunContent::Field(field) = &reparsed.runs[0].content[0] else {
            panic!("reparsed legacy field")
        };
        assert_eq!(
            field.legacy_form.as_ref().map(|form| form.value.clone()),
            Some(LegacyFormFieldValue::Checked(true))
        );
    }

    #[test]
    fn default_namespace_valueless_form_element_gets_a_declared_value_prefix() {
        let xml = format!(
            r#"<p xmlns="{word}" xmlns:q="{word}"><r><fldChar q:fldCharType="begin"><ffData><checkBox><sizeAuto/><checked/></checkBox></ffData></fldChar></r><r><instrText> FORMCHECKBOX </instrText></r><r><fldChar q:fldCharType="end"/></r></p>"#,
            word = crate::namespace::W_NS,
        );
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        let Event::Start(start) = reader.read_event_into(&mut buffer).unwrap() else {
            panic!("paragraph start")
        };
        let prefixes = word_prefixes_at(&start, &[]).unwrap();
        let mut paragraph = CT_P::from_xml_with_prefixes(&mut reader, &prefixes).unwrap();
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        field
            .set_legacy_form_value(LegacyFormFieldValue::Checked(false))
            .unwrap();
        let output = serialized_paragraph(&paragraph);
        assert!(output.contains(r#"<checked q:val="0"/>"#), "{output}");
        let scoped = output.replacen(
            "<w:p>",
            &format!(r#"<w:p xmlns="{0}" xmlns:q="{0}">"#, crate::namespace::W_NS),
            1,
        );
        let mut reader = Reader::from_str(&scoped);
        let mut buffer = Vec::new();
        let Event::Start(start) = reader.read_event_into(&mut buffer).unwrap() else {
            panic!("reparsed paragraph start")
        };
        let prefixes = word_prefixes_at(&start, &["w".to_owned()]).unwrap();
        let reparsed = CT_P::from_xml_with_prefixes(&mut reader, &prefixes).unwrap();
        let RunContent::Field(field) = &reparsed.runs[0].content[0] else {
            panic!("reparsed legacy field")
        };
        assert_eq!(
            field.legacy_form.as_ref().map(|form| form.value.clone()),
            Some(LegacyFormFieldValue::Checked(false))
        );
    }

    #[test]
    fn malformed_or_duplicate_legacy_form_singletons_fail_closed() {
        let cases = [
            (
                "duplicate-name",
                r#"<w:name w:val="one"/><w:name w:val="two"/><w:textInput><w:default w:val="old"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "duplicate-enabled",
                r#"<w:enabled/><w:enabled w:val="0"/><w:textInput><w:default w:val="old"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "duplicate-calculate",
                r#"<w:calcOnExit/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="old"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "duplicate-text-default",
                r#"<w:textInput><w:default w:val="one"/><w:default w:val="two"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "duplicate-max-length",
                r#"<w:textInput><w:maxLength w:val="1"/><w:maxLength w:val="2"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "duplicate-checked",
                r#"<w:checkBox><w:checked/><w:checked w:val="0"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                "duplicate-check-default",
                r#"<w:checkBox><w:default/><w:default w:val="0"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                "duplicate-result",
                r#"<w:ddList><w:result w:val="0"/><w:result w:val="1"/><w:listEntry w:val="one"/><w:listEntry w:val="two"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
            (
                "duplicate-list-default",
                r#"<w:ddList><w:default w:val="0"/><w:default w:val="1"/><w:listEntry w:val="one"/><w:listEntry w:val="two"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
            (
                "invalid-enabled-token",
                r#"<w:enabled w:val="maybe"/><w:textInput><w:default w:val="old"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "invalid-calculate-token",
                r#"<w:calcOnExit w:val="2"/><w:textInput><w:default w:val="old"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "invalid-checked-token",
                r#"<w:checkBox><w:checked w:val="maybe"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                "invalid-check-default-token",
                r#"<w:checkBox><w:default w:val="maybe"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                "invalid-max-length",
                r#"<w:textInput><w:maxLength w:val="many"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "invalid-result",
                r#"<w:ddList><w:result w:val="first"/><w:listEntry w:val="one"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
            (
                "invalid-list-default",
                r#"<w:ddList><w:default w:val="first"/><w:listEntry w:val="one"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
            (
                "missing-name-value",
                r#"<w:name/><w:textInput><w:default w:val="old"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "missing-text-default-value",
                r#"<w:textInput><w:default/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "missing-max-length-value",
                r#"<w:textInput><w:maxLength/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "missing-result-value",
                r#"<w:ddList><w:result/><w:listEntry w:val="one"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
            (
                "missing-list-entry-value",
                r#"<w:ddList><w:listEntry/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
        ];
        for (case, ff_data, instruction) in cases {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("{case}: legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{case}");
        }
    }

    #[test]
    fn legacy_form_ffdata_and_kind_children_require_schema_order() {
        let cases = [
            (
                "common-after-kind",
                r#"<w:checkBox><w:checked/></w:checkBox><w:name w:val="late"/>"#,
                "FORMCHECKBOX",
            ),
            (
                "checkbox-default-after-checked",
                r#"<w:checkBox><w:checked/><w:default/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                "text-default-after-max-length",
                r#"<w:textInput><w:maxLength w:val="12"/><w:default w:val="late"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                "dropdown-result-after-entry",
                r#"<w:ddList><w:listEntry w:val="one"/><w:result w:val="0"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
        ];
        for (case, ff_data, instruction) in cases {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("{case}: legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{case}");
        }
    }

    #[test]
    fn missing_contradictory_and_out_of_range_legacy_form_kinds_fail_closed() {
        let cases = [
            ("missing-kind", r#"<w:name w:val="missing"/>"#, "FORMTEXT"),
            (
                "contradictory-kind",
                r#"<w:checkBox><w:checked/></w:checkBox>"#,
                "FORMTEXT",
            ),
            (
                "out-of-range-selection",
                r#"<w:ddList><w:result w:val="2"/><w:listEntry w:val="one"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
        ];
        for (case, ff_data, instruction) in cases {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("{case}: legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{case}");
        }
    }

    #[test]
    fn known_form_descendants_under_unmodeled_wrappers_are_rejected() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>"#,
            r#"<w:checkBox><w:sizeAuto/></w:checkBox><w:unsupported><w:checked/></w:unsupported>"#,
            r#"</w:ffData></w:fldChar></w:r>"#,
            r#"<w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        assert!(field.validate_legacy_form_owner().is_err());
    }

    #[test]
    fn known_legacy_form_singletons_and_checkbox_size_choice_are_exclusive() {
        let cases = [
            (
                r#"<w:entryMacro/><w:entryMacro/><w:textInput/>"#,
                "FORMTEXT",
            ),
            (r#"<w:exitMacro/><w:exitMacro/><w:textInput/>"#, "FORMTEXT"),
            (r#"<w:helpText/><w:helpText/><w:textInput/>"#, "FORMTEXT"),
            (
                r#"<w:statusText/><w:statusText/><w:textInput/>"#,
                "FORMTEXT",
            ),
            (
                r#"<w:textInput><w:type/><w:type/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                r#"<w:textInput><w:format/><w:format/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                r#"<w:checkBox><w:size/><w:size/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                r#"<w:checkBox><w:sizeAuto/><w:sizeAuto/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                r#"<w:checkBox><w:size/><w:sizeAuto/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
        ];
        for (ff_data, instruction) in cases {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }

        for ff_data in [
            r#"&#x20;<w:textInput/>"#,
            r#"<w:textInput><![CDATA[ ]]></w:textInput>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_ok(), "{ff_data}");
        }
    }

    #[test]
    fn typed_legacy_form_values_reject_duplicate_or_unescape_failing_attributes() {
        for ff_data in [
            r#"<w:name xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="one" q:val="two"/><w:textInput/>"#,
            r#"<w:textInput><w:default w:val="&undefined;"/></w:textInput>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
        let malformed = BytesStart::from_content("w:name w:val", 6);
        assert!(legacy_form_value_attribute(&malformed, &["w".to_owned()]).is_err());
    }

    #[test]
    fn known_legacy_form_leaf_properties_require_values_tokens_and_empty_content() {
        let cases = [
            (r#"<w:entryMacro/><w:textInput/>"#, "FORMTEXT"),
            (r#"<w:exitMacro/><w:textInput/>"#, "FORMTEXT"),
            (r#"<w:textInput><w:type/></w:textInput>"#, "FORMTEXT"),
            (
                r#"<w:textInput><w:type w:val="wrong"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (r#"<w:textInput><w:format/></w:textInput>"#, "FORMTEXT"),
            (r#"<w:checkBox><w:size/></w:checkBox>"#, "FORMCHECKBOX"),
            (
                r#"<w:checkBox><w:size w:val="large"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                r#"<w:checkBox><w:sizeAuto w:val="maybe"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                r#"<w:checkBox><w:size w:val="20"><w:checked/></w:size></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
        ];
        for (ff_data, instruction) in cases {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
        for size in [
            r#"<w:size w:val="20"/>"#,
            r#"<w:size w:val="10.5pt"/>"#,
            r#"<w:sizeAuto/>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox>{size}</w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_ok(), "{size}");
        }
    }

    #[test]
    fn legacy_help_and_status_text_types_accept_only_schema_tokens() {
        for leaf in [
            r#"<w:helpText w:type="bogus"/>"#,
            r#"<w:statusText w:type="bogus"/>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{leaf}<w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{leaf}");
        }
        for leaf in [
            r#"<w:helpText w:type="text"/>"#,
            r#"<w:statusText w:type="autoText"/>"#,
            r#"<w:helpText/>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{leaf}<w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_ok(), "{leaf}");
        }
    }

    #[test]
    fn legacy_form_schema_facets_bound_strings_ranges_and_dropdown_cardinality() {
        let too_long_name = "n".repeat(21);
        let too_long_macro = "m".repeat(34);
        let too_long_string = "s".repeat(256);
        let too_many_entries = (0..26)
            .map(|index| format!(r#"<w:listEntry w:val="{index}"/>"#))
            .collect::<String>();
        let cases = [
            (
                format!(r#"<w:name w:val="{too_long_name}"/><w:textInput/>"#),
                "FORMTEXT",
            ),
            (
                format!(r#"<w:entryMacro w:val="{too_long_macro}"/><w:textInput/>"#),
                "FORMTEXT",
            ),
            (
                format!(r#"<w:textInput><w:default w:val="{too_long_string}"/></w:textInput>"#),
                "FORMTEXT",
            ),
            (
                format!(r#"<w:textInput><w:format w:val="{too_long_string}"/></w:textInput>"#),
                "FORMTEXT",
            ),
            (
                r#"<w:textInput><w:maxLength w:val="0"/></w:textInput>"#.to_owned(),
                "FORMTEXT",
            ),
            (
                r#"<w:textInput><w:maxLength w:val="32768"/></w:textInput>"#.to_owned(),
                "FORMTEXT",
            ),
            (
                format!(r#"<w:ddList><w:listEntry w:val="{too_long_string}"/></w:ddList>"#),
                "FORMDROPDOWN",
            ),
            (
                format!(r#"<w:ddList>{too_many_entries}</w:ddList>"#),
                "FORMDROPDOWN",
            ),
        ];
        for (ff_data, instruction) in cases {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
        let boundary_entries = (0..25)
            .map(|index| format!(r#"<w:listEntry w:val="{index}"/>"#))
            .collect::<String>();
        for (ff_data, instruction) in [
            (
                format!(
                    r#"<w:name w:val="{}"/><w:entryMacro w:val="{}"/><w:textInput><w:default w:val="{}"/><w:maxLength w:val="32767"/><w:format w:val="{}"/></w:textInput>"#,
                    "n".repeat(20),
                    "m".repeat(33),
                    "s".repeat(255),
                    "f".repeat(64),
                ),
                "FORMTEXT",
            ),
            (
                format!(r#"<w:ddList>{boundary_entries}</w:ddList>"#),
                "FORMDROPDOWN",
            ),
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_ok(), "{ff_data}");
        }
    }

    #[test]
    fn legacy_form_name_and_format_use_schema_maxima() {
        for (ff_data, instruction) in [
            (
                format!(r#"<w:name w:val="{}"/><w:textInput/>"#, "n".repeat(21)),
                "FORMTEXT",
            ),
            (
                format!(
                    r#"<w:textInput><w:format w:val="{}"/></w:textInput>"#,
                    "f".repeat(65)
                ),
                "FORMTEXT",
            ),
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }

        let boundary = format!(
            r#"<w:name w:val="{}"/><w:textInput><w:format w:val="{}"/></w:textInput>"#,
            "n".repeat(20),
            "f".repeat(64)
        );
        let paragraph = parse_paragraph(&format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{boundary}</w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        ));
        let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        assert!(field.validate_legacy_form_owner().is_ok());
    }

    #[test]
    fn legacy_help_and_status_values_obey_schema_maxima() {
        for (leaf, max) in [("helpText", 255), ("statusText", 140)] {
            for (length, valid) in [(max, true), (max + 1, false)] {
                let value = "v".repeat(length);
                let paragraph = parse_paragraph(&format!(
                    r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:{leaf} w:type="text" w:val="{value}"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
                ));
                let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                    panic!("legacy field")
                };
                assert_eq!(
                    field.validate_legacy_form_owner().is_ok(),
                    valid,
                    "{leaf} {length}"
                );
            }
        }
    }

    #[test]
    fn dropdown_default_index_is_validated_even_when_result_is_present() {
        let entries = (0..25)
            .map(|index| format!(r#"<w:listEntry w:val="{index}"/>"#))
            .collect::<String>();
        let paragraph = parse_paragraph(&format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:result w:val="0"/><w:default w:val="25"/>{entries}</w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        ));
        let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        assert!(field.validate_legacy_form_owner().is_err());
    }

    #[test]
    fn checkbox_requires_exactly_one_size_choice() {
        for ff_data in [
            r#"<w:checkBox/>"#,
            r#"<w:checkBox><w:default/></w:checkBox>"#,
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
    }

    #[test]
    fn cross_kind_known_legacy_form_children_are_rejected() {
        for (ff_data, instruction) in [
            (r#"<w:textInput><w:checked/></w:textInput>"#, "FORMTEXT"),
            (
                r#"<w:checkBox><w:sizeAuto/><w:listEntry w:val="foreign"/></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                r#"<w:ddList><w:format w:val="foreign"/><w:listEntry w:val="one"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
    }

    #[test]
    fn legacy_form_element_only_containers_reject_character_data() {
        for (ff_data, instruction) in [
            (r#"text<w:textInput/>"#, "FORMTEXT"),
            (r#"<w:textInput>text</w:textInput>"#, "FORMTEXT"),
            (
                r#"<w:checkBox><w:sizeAuto/><![CDATA[text]]></w:checkBox>"#,
                "FORMCHECKBOX",
            ),
            (
                r#"<w:ddList>&amp;<w:listEntry w:val="one"/></w:ddList>"#,
                "FORMDROPDOWN",
            ),
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
    }

    #[test]
    fn legacy_form_rewriter_targets_only_direct_ffdata_owner() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r xmlns:x="urn:producer"><w:fldChar w:fldCharType="begin">"#,
            r#"<x:producer><w:ffData><w:textInput><w:default w:val="decoy"/></w:textInput></w:ffData></x:producer>"#,
            r#"<w:ffData><w:textInput><w:default w:val="old"/></w:textInput></w:ffData>"#,
            r#"</w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        field
            .set_legacy_form_value(LegacyFormFieldValue::Text("new".to_owned()))
            .unwrap();
        let output = serialized_paragraph(&paragraph);
        assert!(
            output.contains(
                r#"<x:producer><w:ffData><w:textInput><w:default w:val="decoy"/></w:textInput></w:ffData></x:producer>"#
            ),
            "{output}"
        );
        assert!(
            output.contains(
                r#"<w:ffData><w:textInput><w:default w:val="new"/></w:textInput></w:ffData>"#
            ),
            "{output}"
        );
    }

    #[test]
    fn known_legacy_form_vocabulary_at_wrong_container_levels_is_rejected() {
        for (ff_data, instruction) in [
            (r#"<w:checked/><w:textInput/>"#, "FORMTEXT"),
            (
                r#"<w:textInput><w:name w:val="nested"/></w:textInput>"#,
                "FORMTEXT",
            ),
            (
                r#"<w:textInput><w:checkBox><w:sizeAuto/></w:checkBox></w:textInput>"#,
                "FORMTEXT",
            ),
            (r#"<w:textInput><w:ffData/></w:textInput>"#, "FORMTEXT"),
        ] {
            let paragraph = parse_paragraph(&format!(
                r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData>{ff_data}</w:ffData></w:fldChar></w:r><w:r><w:instrText> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            ));
            let RunContent::Field(field) = &paragraph.runs[0].content[0] else {
                panic!("legacy field")
            };
            assert!(field.validate_legacy_form_owner().is_err(), "{ff_data}");
        }
    }

    #[test]
    fn empty_text_input_form_can_be_expanded_for_value_mutation() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput/></w:ffData></w:fldChar></w:r>"#,
            r#"<w:r><w:instrText> FORMTEXT </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>old</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("legacy field")
        };
        assert_eq!(
            field.legacy_form.as_ref().unwrap().value,
            LegacyFormFieldValue::Text("old".to_owned())
        );
        field
            .set_legacy_form_value(LegacyFormFieldValue::Text("new".to_owned()))
            .unwrap();
        let output = serialized_paragraph(&paragraph);
        assert!(
            output.contains(
                r#"<w:ffData><w:textInput><w:default xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="new"/></w:textInput></w:ffData>"#
            ),
            "{output}"
        );
    }

    #[test]
    fn split_run_keeps_properties_space_and_later_markers_in_order() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:bookmarkStart w:id="1" w:name="mark"/>"#,
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Hello world</w:t></w:r>"#,
            r#"<w:bookmarkEnd w:id="1"/><w:r><w:t>!</w:t></w:r>"#,
        ));
        paragraph.split_run(0, 6).unwrap();

        let texts = paragraph.runs.iter().map(CT_R::text).collect::<Vec<_>>();
        assert_eq!(texts, ["Hello ", "world", "!"]);
        assert_eq!(paragraph.runs[0].properties, paragraph.runs[1].properties);
        let output = serialized_paragraph(&paragraph);
        assert!(
            output.contains(r#"<w:t xml:space="preserve">Hello </w:t>"#),
            "{output}"
        );
        let world = output.find(">world<").unwrap();
        let bookmark_end = output.find("bookmarkEnd").unwrap();
        let exclamation = output.find(">!<").unwrap();
        assert!(
            world < bookmark_end && bookmark_end < exclamation,
            "{output}"
        );
    }

    #[test]
    fn split_run_keeps_both_parts_in_the_hyperlink_with_raw_children_in_order() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:hyperlink w:anchor="target"><w:r><w:t>ab</w:t><x:mark/><w:t>cd</w:t><x:end/></w:r></w:hyperlink>"#,
            r#"<w:r><w:t>after</w:t></w:r>"#,
        ));
        paragraph.split_run(0, 3).unwrap();

        let texts = paragraph.runs.iter().map(CT_R::text).collect::<Vec<_>>();
        assert_eq!(texts, ["abc", "d", "after"]);
        let hyperlink = &paragraph.hyperlinks[0];
        assert_eq!((hyperlink.run_start, hyperlink.run_end), (0, 2));
        let output = serialized_paragraph(&paragraph);
        assert_eq!(output.matches("<w:hyperlink").count(), 1, "{output}");
        assert!(
            output.contains(concat!(
                r#"<w:r><w:t>ab</w:t><x:mark/><w:t>c</w:t></w:r>"#,
                r#"<w:r><w:t>d</w:t><x:end/></w:r></w:hyperlink>"#,
            )),
            "{output}"
        );
    }

    #[test]
    fn split_run_rejects_missing_runs_edge_offsets_and_field_results_unchanged() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r><w:t>Page </w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText> PAGE </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>12</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        ));
        let before = serialized_paragraph(&paragraph);
        let field_run = paragraph
            .runs
            .iter()
            .position(|run| {
                run.content
                    .iter()
                    .any(|content| matches!(content, RunContent::Field(_)))
            })
            .unwrap();
        assert_eq!(paragraph.runs[field_run].text(), "12");
        assert_eq!(paragraph.split_run(field_run, 0), Ok(field_run));
        assert_eq!(
            paragraph.split_run(field_run, 1),
            Err(RunSplitError::OffsetOutOfRange { offset: 1, len: 0 })
        );
        assert_eq!(paragraph.split_run(0, 0), Ok(0));
        assert_eq!(paragraph.split_run(0, 5), Ok(1));
        assert_eq!(
            paragraph.split_run(0, 6),
            Err(RunSplitError::OffsetOutOfRange { offset: 6, len: 5 })
        );
        let run_count = paragraph.runs.len();
        assert_eq!(
            paragraph.split_run(run_count, 1),
            Err(RunSplitError::RunOutOfRange {
                run_index: run_count,
                run_count
            })
        );
        assert_eq!(serialized_paragraph(&paragraph), before);
    }

    /// #298: a paragraph that holds a picture, an object, a field or a break
    /// and no text is not empty.
    #[test]
    fn paragraphs_without_text_can_still_show_content() {
        let visible = [
            r#"<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/></w:drawing></w:r>"#,
            r#"<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"/></w:pict></w:r>"#,
            r#"<w:r><w:object><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office"/></w:object></w:r>"#,
            r#"<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent></w:r>"#,
            r#"<w:fldSimple w:instr=" PAGE "/>"#,
            r#"<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>"#,
            r#"<w:r><w:tab/></w:r>"#,
            r#"<w:r><w:br w:type="page"/></w:r>"#,
            r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#,
            r#"<w:sdt><w:sdtContent><w:r><w:pict/></w:r></w:sdtContent></w:sdt>"#,
            r#"<w:ins w:id="1" w:author="A"><w:r><w:br/></w:r></w:ins>"#,
            r#"<w:smartTag w:element="place"><w:r><w:tab/></w:r></w:smartTag>"#,
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>"#,
            r#"<w:bdo w:val="rtl"><w:r><w:t>abc</w:t></w:r></w:bdo>"#,
            r#"<w:dir w:val="rtl"><w:r><w:pict/></w:r></w:dir>"#,
        ];
        for xml in visible {
            assert!(parse_paragraph(xml).has_visible_content(), "{xml}");
        }
        let empty = [
            "",
            r#"<w:r><w:t xml:space="preserve">  </w:t></w:r>"#,
            r#"<w:r><w:lastRenderedPageBreak/></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:del w:id="1" w:author="A"><w:r><w:delText>gone</w:delText></w:r></w:del>"#,
            r#"<w:bookmarkStart w:id="0" w:name="b"/><w:bookmarkEnd w:id="0"/>"#,
            r#"<w:sdt><w:sdtContent><w:r><w:t> </w:t></w:r></w:sdtContent></w:sdt>"#,
            r#"<w:r><w:softHyphen/></w:r>"#,
            r#"<w:bdo w:val="rtl"><w:r><w:t> </w:t></w:r></w:bdo>"#,
        ];
        for xml in empty {
            assert!(!parse_paragraph(xml).has_visible_content(), "{xml}");
        }
    }

    #[test]
    fn symbols_and_special_characters_reopen_in_source_order() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:t>a</w:t><w:sym w:font="Wingdings" w:char="F0FC"/><w:cr/>"#,
            r#"<w:noBreakHyphen/><w:tab/><w:softHyphen/>"#,
            r#"<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="dot"/>"#,
            r#"<w:lastRenderedPageBreak/><w:br/><w:t>b</w:t></w:r>"#,
        ));
        let content = &paragraph.runs[0].content;
        assert_eq!(
            content[1],
            RunContent::Symbol {
                font: "Wingdings".to_owned(),
                char_code: 0xF0FC,
            }
        );
        assert_eq!(
            content[2],
            RunContent::SpecialCharacter(SpecialCharacter::CarriageReturn)
        );
        assert_eq!(
            content[3],
            RunContent::SpecialCharacter(SpecialCharacter::NoBreakHyphen)
        );
        assert_eq!(content[4], RunContent::Tab);
        assert_eq!(
            content[5],
            RunContent::SpecialCharacter(SpecialCharacter::SoftHyphen)
        );
        assert_eq!(
            content[6],
            RunContent::SpecialCharacter(SpecialCharacter::PositionalTab {
                alignment: ST_PTabAlignment::Right,
                relative_to: ST_PTabRelativeTo::Margin,
                leader: ST_PTabLeader::Dot,
            })
        );

        let output = serialized_paragraph(&paragraph);
        let ordered = [
            "<w:t>a</w:t>",
            r#"<w:sym w:font="Wingdings" w:char="F0FC"/>"#,
            "<w:cr/>",
            "<w:noBreakHyphen/>",
            "<w:tab/>",
            "<w:softHyphen/>",
            r#"<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="dot"/>"#,
            "<w:lastRenderedPageBreak/>",
            "<w:br/>",
            "<w:t>b</w:t>",
        ]
        .map(|needle| {
            output
                .find(needle)
                .unwrap_or_else(|| panic!("{needle} in {output}"))
        });
        assert!(ordered.windows(2).all(|pair| pair[0] < pair[1]), "{output}");
    }

    #[test]
    fn a_special_character_outside_the_modeled_shape_stays_raw() {
        let paragraph = parse_paragraph(concat!(
            r#"<w:r><w:sym w:font="Wingdings" w:char="zzzz"/>"#,
            r#"<w:sym w:font="Wingdings"/>"#,
            r#"<w:ptab w:alignment="right" w:leader="dot"/>"#,
            r#"<w:cr w:producerFlag="1"/></w:r>"#,
        ));
        assert!(paragraph.runs[0].content.is_empty());
        assert_eq!(paragraph.runs[0].extra_xml.len(), 4);

        let output = serialized_paragraph(&paragraph);
        for retained in [
            r#"w:char="zzzz""#,
            r#"<w:sym w:font="Wingdings"/>"#,
            r#"<w:ptab w:alignment="right" w:leader="dot"/>"#,
            r#"<w:cr w:producerFlag="1"/>"#,
        ] {
            assert!(
                output.contains(retained),
                "{retained} missing from {output}"
            );
        }
    }

    #[test]
    fn split_run_partitions_zero_width_mixed_content_without_reordering() {
        let mut paragraph = parse_paragraph(concat!(
            r#"<w:r xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:t>ab</w:t><w:tab/><w:br/><w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText>"#,
            r#"<mc:AlternateContent><mc:Choice Requires="wps">"#,
            r#"<w:drawing><wp:anchor behindDoc="0">"#,
            r#"<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>"#,
            r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
            r#"<wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Shape"/>"#,
            r#"<a:graphic><a:graphicData><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp></a:graphicData></a:graphic>"#,
            r#"</wp:anchor></w:drawing></mc:Choice>"#,
            r#"<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback></mc:AlternateContent>"#,
            r#"<w:sym w:font="Wingdings" w:char="F0FC"/><x:raw/><w:t>cd</w:t></w:r>"#,
        ));
        assert_eq!(paragraph.runs.len(), 1);
        assert_eq!(paragraph.runs[0].alt_drawings.len(), 1);

        assert_eq!(paragraph.split_run(0, 2), Ok(1));
        assert_eq!(paragraph.runs.len(), 2);
        assert_eq!(paragraph.runs[0].text(), "ab");
        assert_eq!(paragraph.runs[1].text(), "\t\ncd");
        assert!(paragraph.runs[0].alt_drawings.is_empty());
        assert_eq!(paragraph.runs[1].alt_drawings.len(), 1);

        let output = serialized_paragraph(&paragraph);
        assert_eq!(
            output.matches("<mc:AlternateContent").count(),
            1,
            "{output}"
        );
        let ordered = [
            "<w:t>ab</w:t>",
            "<w:tab/>",
            "<w:br/>",
            "fldCharType=\"begin\"",
            "<w:instrText> PAGE </w:instrText>",
            "<mc:AlternateContent",
            "<w:sym",
            "<x:raw/>",
            "<w:t>cd</w:t>",
        ]
        .map(|needle| output.find(needle).unwrap());
        assert!(ordered.windows(2).all(|pair| pair[0] < pair[1]), "{output}");
    }
    #[test]
    fn authored_simple_and_complex_fields_reopen_with_identical_semantics() {
        for form in [FieldForm::Simple, FieldForm::Complex] {
            let instruction = FieldInstruction::new(
                "REF",
                vec![FieldArgument::Text("target name".into())],
                vec![],
            )
            .unwrap();
            let field =
                Field::from_instruction(instruction, form, vec![CT_R::new("cached")]).unwrap();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(field_run(field, None));
            let output = serialized_paragraph(&paragraph);
            let reopened = CT_P::from_xml_fragment(
                output
                    .replacen(
                        "<w:p>",
                        &format!("<w:p xmlns:w=\"{}\">", crate::namespace::W_NS),
                        1,
                    )
                    .as_bytes(),
            )
            .unwrap();
            let field = parsed_field(&reopened, 0);
            assert_eq!(field.form(), form);
            assert_eq!(
                field.instruction.arguments,
                vec![FieldArgument::Text("target name".into())]
            );
            assert_eq!(field.cached_result, "cached");
        }
    }

    #[test]
    fn authored_nested_instruction_fields_preserve_operand_order() {
        let inner = Field::from_raw(
            "MERGEFIELD name",
            FieldForm::Complex,
            vec![CT_R::new("Ada")],
        )
        .unwrap();
        let instruction = FieldInstruction::new(
            "IF",
            vec![
                FieldArgument::Nested(Box::new(inner)),
                FieldArgument::Text("=".into()),
                FieldArgument::Text("Ada".into()),
                FieldArgument::Text("yes".into()),
                FieldArgument::Text("no".into()),
            ],
            vec![],
        )
        .unwrap();
        assert!(Field::from_instruction(instruction.clone(), FieldForm::Simple, vec![]).is_err());
        let field =
            Field::from_instruction(instruction, FieldForm::Complex, vec![CT_R::new("yes")])
                .unwrap();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(field_run(field, None));
        let output = serialized_paragraph(&paragraph);
        assert_eq!(output.matches("fldCharType=\"begin\"").count(), 2);
        let reopened = CT_P::from_xml_fragment(
            output
                .replacen(
                    "<w:p>",
                    &format!("<w:p xmlns:w=\"{}\">", crate::namespace::W_NS),
                    1,
                )
                .as_bytes(),
        )
        .unwrap();
        let field = parsed_field(&reopened, 0);
        assert_eq!(
            field.nested_fields_in_source_order()[0].instruction.name,
            "MERGEFIELD"
        );
        assert_eq!(field.cached_result, "yes");
    }

    #[test]
    fn field_lock_and_dirty_preserve_absent_false_and_true() {
        for form in [FieldForm::Simple, FieldForm::Complex] {
            for value in [None, Some(false), Some(true)] {
                let mut field = Field::from_raw("UNKNOWN", form, vec![CT_R::new("cache")]).unwrap();
                field.set_locked(value);
                field.dirty = value;
                let mut paragraph = CT_P::new();
                paragraph.runs.push(field_run(field, None));
                let output = serialized_paragraph(&paragraph);
                let reopened = CT_P::from_xml_fragment(
                    output
                        .replacen(
                            "<w:p>",
                            &format!("<w:p xmlns:w=\"{}\">", crate::namespace::W_NS),
                            1,
                        )
                        .as_bytes(),
                )
                .unwrap();
                let field = parsed_field(&reopened, 0);
                assert_eq!(field.locked(), value);
                assert_eq!(field.dirty, value);
            }
        }
    }

    #[test]
    fn typed_field_arguments_cannot_inject_instruction_tokens() {
        for value in [
            "",
            "\\h",
            "two words",
            "quoted\"name",
            "path\\name",
            "x\" \\h \"y",
        ] {
            let instruction =
                FieldInstruction::new("REF", vec![FieldArgument::Text(value.into())], vec![])
                    .unwrap();
            let field = Field::from_instruction(instruction, FieldForm::Simple, vec![]).unwrap();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(field_run(field, None));
            let reopened = CT_P::from_xml_fragment(
                serialized_paragraph(&paragraph)
                    .replacen(
                        "<w:p>",
                        &format!("<w:p xmlns:w=\"{}\">", crate::namespace::W_NS),
                        1,
                    )
                    .as_bytes(),
            )
            .unwrap();
            let field = parsed_field(&reopened, 0);
            assert_eq!(
                field.instruction.arguments,
                vec![FieldArgument::Text(value.into())]
            );
            assert!(field.instruction.switches.is_empty());
        }
        for raw in ["", " PAGE \"unterminated", "\\h", "PA\0GE"] {
            assert!(
                Field::from_raw(raw, FieldForm::Simple, vec![]).is_err(),
                "{raw:?}"
            );
        }
    }

    #[test]
    fn field_cache_runs_keep_controls_and_properties_in_order() {
        let mut first = CT_R::new("first");
        first.properties = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        first.content.push(RunContent::Tab);
        first.content.push(RunContent::Break(BreakType::Line));
        let second = CT_R::new("second");
        for form in [FieldForm::Simple, FieldForm::Complex] {
            let field =
                Field::from_raw("UNKNOWN", form, vec![first.clone(), second.clone()]).unwrap();
            assert_eq!(field.cached_result, "first\t\nsecond");
            let mut paragraph = CT_P::new();
            paragraph.runs.push(field_run(field, None));
            let output = serialized_paragraph(&paragraph);
            assert!(output.contains("<w:b"), "{output}");
            assert!(output.find("first").unwrap() < output.find("<w:tab").unwrap());
            assert!(output.find("<w:br").unwrap() < output.find("second").unwrap());
            let reopened = CT_P::from_xml_fragment(
                output
                    .replacen(
                        "<w:p>",
                        &format!("<w:p xmlns:w=\"{}\">", crate::namespace::W_NS),
                        1,
                    )
                    .as_bytes(),
            )
            .unwrap();
            assert_eq!(parsed_field(&reopened, 0).cached_result, "first\t\nsecond");
        }
    }

    #[test]
    fn typed_cache_replaces_simple_and_complex_results_without_losing_controls() {
        for xml in [
            r#"<w:fldSimple w:instr=" REF Target \f " w:dirty="1" xmlns:x="urn:producer" x:flag="kept"><w:r><w:t>OLD</w:t></w:r></w:fldSimple>"#,
            r#"<w:r xmlns:x="urn:producer"><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin" x:flag="kept"/></w:r><w:r><w:instrText xml:space="preserve"> REF Target \f </w:instrText></w:r><w:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="separate"/><w:t>OLD</w:t><w:fldChar w:fldCharType="end"/></w:r>"#,
        ] {
            let mut paragraph = parse_paragraph(xml);
            let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
                panic!("field");
            };
            let instruction = field.instruction.clone();
            let mut reference = CT_R::new("");
            reference.content = vec![RunContent::FootnoteRef {
                id: 37,
                custom_mark: None,
            }];
            reference.properties = Some(CT_RPr {
                bold: Some(true),
                ..Default::default()
            });
            field.set_cached_runs(vec![reference]).unwrap();
            assert_eq!(field.cached_result, "");
            assert_eq!(field.instruction, instruction);
            let output = serialized_paragraph(&paragraph);
            assert!(output.contains("x:flag=\"kept\""), "{output}");
            assert!(output.contains("footnoteReference w:id=\"37\""), "{output}");
            assert!(!output.contains("OLD"), "{output}");
            let reopened = parse_paragraph(
                output
                    .strip_prefix("<w:p>")
                    .unwrap()
                    .strip_suffix("</w:p>")
                    .unwrap(),
            );
            assert_eq!(parsed_field(&reopened, 0).cached_result, "");
            assert_eq!(serialized_paragraph(&reopened), output);
        }
    }

    #[test]
    fn typed_cache_shared_run_keeps_prefix_suffix_and_nested_result() {
        let mut paragraph = parse_paragraph(
            r#"<w:r xmlns:x="urn:producer" x:keep="yes"><w:rPr><w:b/></w:rPr><w:t>PRE</w:t><w:fldChar w:fldCharType="begin"/><w:instrText> REF Target \f </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>OLD</w:t><w:fldChar w:fldCharType="end"/><w:t>POST</w:t></w:r>"#,
        );
        let field = paragraph
            .runs
            .iter_mut()
            .flat_map(|run| &mut run.content)
            .find_map(|content| match content {
                RunContent::Field(field) => Some(field),
                _ => None,
            })
            .unwrap();
        let child =
            Field::from_raw("UNKNOWN", FieldForm::Complex, vec![CT_R::new("RICH")]).unwrap();
        field
            .set_cached_runs(vec![field_run(
                child,
                Some(CT_RPr {
                    italic: Some(true),
                    ..Default::default()
                }),
            )])
            .unwrap();
        let output = serialized_paragraph(&paragraph);
        assert!(
            output.contains("PRE") && output.contains("POST") && output.contains("RICH"),
            "{output}"
        );
        assert!(!output.contains("OLD"), "{output}");
        let reopened = parse_paragraph(
            output
                .strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(reopened.text(), "PRERICHPOST");
        assert_eq!(serialized_paragraph(&reopened), output);
    }

    #[test]
    fn typed_cache_nested_child_preserves_parent_and_namespace_bindings() {
        for outer in [FieldForm::Simple, FieldForm::Complex] {
            let child =
                Field::from_raw(r"REF Target \f", FieldForm::Complex, vec![CT_R::new("OLD")])
                    .unwrap();
            let root = Field::from_raw("UNKNOWN", outer, vec![field_run(child, None)]).unwrap();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(field_run(root, None));
            let source = serialized_paragraph(&paragraph);
            let mut parsed = parse_paragraph(
                source
                    .strip_prefix("<w:p>")
                    .unwrap()
                    .strip_suffix("</w:p>")
                    .unwrap(),
            );
            let RunContent::Field(parent) = &mut parsed.runs[0].content[0] else {
                panic!("parent");
            };
            let child = parent.cached_field_mut(0).unwrap();
            let mut reference = CT_R::new("kept ");
            reference.content.push(RunContent::EndnoteRef {
                id: 41,
                custom_mark: None,
            });
            child.set_cached_runs(vec![reference]).unwrap();
            parent.refresh_cached_field_projection();
            let output = serialized_paragraph(&parsed);
            assert!(output.contains("endnoteReference w:id=\"41\""), "{output}");
            assert!(!output.contains("OLD"), "{output}");
            assert!(output.contains("UNKNOWN"), "{output}");
            let reopened = parse_paragraph(
                output
                    .strip_prefix("<w:p>")
                    .unwrap()
                    .strip_suffix("</w:p>")
                    .unwrap(),
            );
            assert_eq!(parsed_field(&reopened, 0).cached_result, "kept ");
        }
        let mut parsed = CT_P::from_xml_fragment(br#"<q:p xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="urn:foreign"><q:fldSimple xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="urn:foreign" q:instr=" REF Target \f " xmlns:x="urn:producer" x:keep="yes"><q:r><q:t>OLD</q:t></q:r></q:fldSimple></q:p>"#).unwrap();
        let RunContent::Field(field) = &mut parsed.runs[0].content[0] else {
            panic!("alias field");
        };
        let mut reference = CT_R::new("");
        reference.content = vec![RunContent::CommentReference {
            id: 43,
            raw_before: 0,
        }];
        field.set_cached_runs(vec![reference]).unwrap();
        let output = serialized_paragraph(&parsed);
        assert!(output.contains("xmlns:w=\"urn:foreign\""), "{output}");
        assert!(output.contains("x:keep=\"yes\""), "{output}");
        assert!(output.contains("commentReference w:id=\"43\""), "{output}");
        assert!(!output.contains("OLD"), "{output}");
    }

    #[test]
    fn typed_cache_parse_keeps_empty_typed_results_and_original_shared_owner_bytes() {
        let source = r#"<w:r><w:rPr><w:b/></w:rPr><w:t>PRE</w:t><w:fldChar w:fldCharType="begin"/><w:instrText> REF A \f </w:instrText><w:fldChar w:fldCharType="separate"/><w:footnoteReference w:id="17"/><w:fldChar w:fldCharType="end"/><w:t>MID</w:t><w:fldChar w:fldCharType="begin"/><w:instrText> REF B \f </w:instrText><w:fldChar w:fldCharType="separate"/><w:endnoteReference w:id="19"/><w:fldChar w:fldCharType="end"/><w:t>POST</w:t></w:r>"#;
        let paragraph = parse_paragraph(source);
        assert_eq!(
            serialized_paragraph(&paragraph),
            format!("<w:p>{source}</w:p>")
        );
        let fields = paragraph
            .runs()
            .into_iter()
            .flat_map(|run| &run.content)
            .filter_map(|content| match content {
                RunContent::Field(field) => Some(field),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].cached_result, "");
        assert_eq!(fields[1].cached_result, "");
        assert!(
            fields[0]
                .cached_result_runs()
                .unwrap()
                .iter()
                .flat_map(|run| &run.content)
                .any(|content| matches!(content, RunContent::FootnoteRef { id: 17, .. }))
        );
        assert!(
            !fields[0]
                .cached_result_runs()
                .unwrap()
                .iter()
                .flat_map(|run| &run.content)
                .any(|content| matches!(content, RunContent::EndnoteRef { .. }))
        );
        assert!(
            fields[1]
                .cached_result_runs()
                .unwrap()
                .iter()
                .flat_map(|run| &run.content)
                .any(|content| matches!(content, RunContent::EndnoteRef { id: 19, .. }))
        );
        assert!(
            !fields[1]
                .cached_result_runs()
                .unwrap()
                .iter()
                .flat_map(|run| &run.content)
                .any(|content| matches!(content, RunContent::FootnoteRef { .. }))
        );
    }

    #[test]
    fn typed_cache_rejection_preserves_source_and_cached_state() {
        let mut field = Field::new(r"REF Target \f", "OLD");
        let before = format!("{field:?}");
        assert!(field.set_cached_runs(vec![CT_R::new("invalid\0")]).is_err());
        assert_eq!(field.cached_result, "OLD");
        assert_eq!(format!("{field:?}"), before);
        field.set_locked(Some(true));
        assert!(field.set_cached_runs(vec![CT_R::new("fresh")]).is_err());
        assert_eq!(field.cached_result, "OLD");
    }

    #[test]
    fn field_property_edits_preserve_unmodelled_xml_verbatim() {
        for xml in [
            r#"<a:fldSimple xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" a:instr="UNKNOWN" a:fldLock="1" xmlns:x="urn:producer"><x:keep note='yes'> exact </x:keep><a:r><a:t>cache</a:t></a:r></a:fldSimple>"#,
            r#"<a:r xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer"><a:fldChar a:fldCharType="begin" a:fldLock="1"><x:keep note='yes'> exact </x:keep></a:fldChar></a:r><a:r xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><a:instrText> UNKNOWN </a:instrText></a:r><a:r xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><a:fldChar a:fldCharType="separate"/><a:t>cache</a:t><a:fldChar a:fldCharType="end"/></a:r>"#,
        ] {
            let mut paragraph = parse_paragraph(xml);
            let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
                panic!("field not typed");
            };
            assert_eq!(field.locked(), Some(true));
            field.set_locked(Some(false));
            let output = serialized_paragraph(&paragraph);
            assert!(
                output.contains("<x:keep note='yes'> exact </x:keep>"),
                "{output}"
            );
            let reopened = CT_P::from_xml_fragment(
                output
                    .replacen(
                        "<w:p>",
                        &format!("<w:p xmlns:w=\"{}\">", crate::namespace::W_NS),
                        1,
                    )
                    .as_bytes(),
            )
            .unwrap();
            assert_eq!(parsed_field(&reopened, 0).locked(), Some(false));
        }
    }
    #[test]
    fn unknown_typed_switch_operands_reopen_without_changing_known_flags() {
        for form in [FieldForm::Simple, FieldForm::Complex] {
            let instruction = FieldInstruction::new(
                "PRODUCER",
                vec![],
                vec![FieldSwitch {
                    name: "q".into(),
                    argument: Some(FieldArgument::Text("operand".into())),
                }],
            )
            .unwrap();
            let field =
                Field::from_instruction(instruction, form, vec![CT_R::new("cache")]).unwrap();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(field_run(field, None));
            let reopened = parse_paragraph(
                &serialized_paragraph(&paragraph)
                    .replace("<w:p>", "")
                    .replace("</w:p>", ""),
            );
            assert_eq!(
                parsed_field(&reopened, 0).instruction.switches[0].argument,
                Some(FieldArgument::Text("operand".into()))
            );
            assert!(parsed_field(&reopened, 0).instruction.arguments.is_empty());
        }
        let nested = Field::from_raw(
            "MERGEFIELD name",
            FieldForm::Complex,
            vec![CT_R::new("Ada")],
        )
        .unwrap();
        let instruction = FieldInstruction::new(
            "PRODUCER",
            vec![],
            vec![FieldSwitch {
                name: "q".into(),
                argument: Some(FieldArgument::Nested(Box::new(nested))),
            }],
        )
        .unwrap();
        let field =
            Field::from_instruction(instruction, FieldForm::Complex, vec![CT_R::new("cache")])
                .unwrap();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(field_run(field, None));
        let output = serialized_paragraph(&paragraph);
        let reopened = parse_paragraph(&output.replace("<w:p>", "").replace("</w:p>", ""));
        let field = parsed_field(&reopened, 0);
        assert!(field.instruction.arguments.is_empty());
        assert!(
            matches!(&field.instruction.switches[0].argument, Some(FieldArgument::Nested(nested)) if nested.cached_result == "Ada")
        );
        let raw = r#"<w:fldSimple w:instr='REF \h "target name"'><w:r><w:t>cache</w:t></w:r></w:fldSimple>"#;
        let original = parse_paragraph(raw);
        let instruction = &parsed_field(&original, 0).instruction;
        assert_eq!(
            instruction.arguments,
            vec![FieldArgument::Text("target name".into())]
        );
        assert_eq!(instruction.switches[0].argument, None);
        assert!(serialized_paragraph(&original).contains(raw));
        assert!(
            FieldInstruction::new(
                "REF",
                vec![],
                vec![FieldSwitch {
                    name: "h".into(),
                    argument: Some(FieldArgument::Text("target name".into()))
                }]
            )
            .is_err()
        );
        let producer_raw = r#"<w:fldSimple w:instr='PRODUCER \q "value"'><w:r><w:t>cache</w:t></w:r></w:fldSimple>"#;
        assert!(serialized_paragraph(&parse_paragraph(producer_raw)).contains(producer_raw));
    }

    #[test]
    fn generated_table_operands_parse_bare_and_quoted_without_flag_collisions() {
        for raw in [
            r#"INDEX \z 1033 \f a \h A \e " | " \l "; " \g " to " \k " => " \s Chapter \r"#,
            r#"INDEX \z "1033" \f "a" \h "A" \e " | " \l "; " \g " to " \k " => " \s "Chapter" \r"#,
        ] {
            let field = Field::from_raw(raw, FieldForm::Complex, vec![]).unwrap();
            assert!(field.instruction.arguments.is_empty());
            for (name, operand) in [
                ("z", "1033"),
                ("f", "a"),
                ("h", "A"),
                ("e", " | "),
                ("l", "; "),
                ("g", " to "),
                ("k", " => "),
                ("s", "Chapter"),
            ] {
                assert_eq!(
                    field
                        .instruction
                        .switches
                        .iter()
                        .find(|switch| switch.name == name)
                        .unwrap()
                        .argument,
                    Some(FieldArgument::Text(operand.into()))
                );
            }
            assert_eq!(field.instruction.switches.last().unwrap().argument, None);
            assert_eq!(field.effective_instruction_text(), raw);
        }
        for (raw, operands, flags) in [
            (
                r#"XE "Alpha:Beta" \f a \r Range \t "See Alpha" \b \i"#,
                vec![("f", "a"), ("r", "Range"), ("t", "See Alpha")],
                vec!["b", "i"],
            ),
            (
                r#"TA \l "Long citation" \s Short \c 16 \r Range \b \i"#,
                vec![
                    ("l", "Long citation"),
                    ("s", "Short"),
                    ("c", "16"),
                    ("r", "Range"),
                ],
                vec!["b", "i"],
            ),
            (
                r#"TOA \c 1 \e " | " \l "; " \g " to " \h \p"#,
                vec![("c", "1"), ("e", " | "), ("l", "; "), ("g", " to ")],
                vec!["h", "p"],
            ),
            (r#"TOC \c Figure \h"#, vec![("c", "Figure")], vec!["h"]),
            (
                r#"TOC \a "Custom label" \h"#,
                vec![("a", "Custom label")],
                vec!["h"],
            ),
        ] {
            let field = Field::from_raw(raw, FieldForm::Simple, vec![]).unwrap();
            for (name, operand) in operands {
                assert_eq!(
                    field
                        .instruction
                        .switches
                        .iter()
                        .find(|switch| switch.name == name)
                        .unwrap()
                        .argument,
                    Some(FieldArgument::Text(operand.into()))
                );
            }
            for name in flags {
                assert_eq!(
                    field
                        .instruction
                        .switches
                        .iter()
                        .find(|switch| switch.name == name)
                        .unwrap()
                        .argument,
                    None
                );
            }
        }
        let reference = parse_field_instruction(r#"REF \h "target name""#);
        assert_eq!(
            reference.arguments,
            vec![FieldArgument::Text("target name".into())]
        );
        assert_eq!(reference.switches[0].argument, None);
        let sequence = parse_field_instruction(r#"SEQ Chapter \r 7"#);
        assert_eq!(
            sequence.switches[0].argument,
            Some(FieldArgument::Text("7".into()))
        );
        let malformed = parse_field_instruction(r#"TOA \c "1"#);
        assert!(!malformed.quotes_are_balanced());
    }

    #[test]
    fn equal_text_nested_replacement_writes_its_new_cache_properties() {
        for depth in [1, 2] {
            for switch_operand in [false, true] {
                let instruction_prefix = if switch_operand {
                    "PRODUCER \\q "
                } else {
                    "PRODUCER "
                };
                let source = format!(
                    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>{instruction_prefix}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD name</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Ada</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cache</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
                );
                let source = if depth == 2 {
                    format!(
                        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>OUTER </w:instrText></w:r>{source}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>outer cache</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
                    )
                } else {
                    source
                };
                let mut paragraph = parse_paragraph(&source);
                let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
                    panic!("outer field");
                };
                let field = if depth == 2 {
                    let FieldArgument::Nested(middle) = &mut field.instruction.arguments[0] else {
                        panic!("middle field");
                    };
                    middle.as_mut()
                } else {
                    field
                };
                let mut cached = CT_R::new("Ada");
                cached.properties = Some(CT_RPr {
                    bold: Some(true),
                    ..Default::default()
                });
                let nested =
                    Field::from_raw("MERGEFIELD name", FieldForm::Complex, vec![cached]).unwrap();
                if switch_operand {
                    field.instruction.switches[0].argument =
                        Some(FieldArgument::Nested(Box::new(nested)));
                } else {
                    field.instruction.arguments[0] = FieldArgument::Nested(Box::new(nested));
                }
                let output = serialized_paragraph(&paragraph);
                assert!(output.contains("<w:b"), "{output}");
                let reopened = parse_paragraph(&output.replace("<w:p>", "").replace("</w:p>", ""));
                let mut nested = parsed_field(&reopened, 0).nested_fields_in_source_order()[0];
                if depth == 2 {
                    nested = nested.nested_fields_in_source_order()[0];
                }
                assert_eq!(nested.cached_result, "Ada");
                assert_eq!(
                    nested.cached_display_segments()[0].1.unwrap().bold,
                    Some(true)
                );
            }
        }
    }

    #[test]
    fn cached_opaque_namespaces_preserve_foreign_lookalikes_and_inherited_word_names() {
        for raw in [
            br#"<p:fldChar xmlns:p="urn:producer" note='exact'/>"#.as_slice(),
            b"<w:lastRenderedPageBreak/>",
        ] {
            let mut cached = CT_R::new("cache");
            cached.extra_xml.push(raw.to_vec());
            let field = Field::from_raw("UNKNOWN", FieldForm::Simple, vec![cached]).unwrap();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(field_run(field, None));
            let output = serialized_paragraph(&paragraph);
            assert!(output.contains(std::str::from_utf8(raw).unwrap()));
            let reopened = parse_paragraph(&output.replace("<w:p>", "").replace("</w:p>", ""));
            assert!(serialized_paragraph(&reopened).contains(std::str::from_utf8(raw).unwrap()));
        }
        for raw in [br#"<a:fldChar xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" a:fldCharType="begin"/>"#.as_slice(), b"<w:fldChar w:fldCharType=\"begin\"/>", b"<w:unknown>"] {
            let mut cached = CT_R::new("cache");
            cached.extra_xml.push(raw.to_vec());
            assert!(Field::from_raw("UNKNOWN", FieldForm::Simple, vec![cached]).is_err());
        }
    }
    #[test]
    fn nested_cached_fields_preserve_both_wire_forms() {
        for outer_form in [FieldForm::Simple, FieldForm::Complex] {
            for child_form in [FieldForm::Simple, FieldForm::Complex] {
                let mut result = CT_R::new("OLD");
                result.properties = Some(CT_RPr {
                    italic: Some(true),
                    ..CT_RPr::default()
                });
                let child = Field::from_raw("PAGE", child_form, vec![result]).unwrap();
                let nested = field_run(child, None);
                let outer = Field::from_raw(
                    "UNKNOWN",
                    outer_form,
                    vec![CT_R::new("prefix "), nested, CT_R::new(" suffix")],
                )
                .unwrap();
                let mut authored = CT_P::new();
                authored.runs.push(field_run(outer, None));
                let original = serialized_paragraph(&authored);
                let mut reopened =
                    parse_paragraph(&original.replace("<w:p>", "").replace("</w:p>", ""));
                assert_eq!(serialized_paragraph(&reopened), original);
                let RunContent::Field(outer) = &mut reopened.runs[0].content[0] else {
                    panic!("outer");
                };
                assert_eq!(outer.form(), outer_form);
                assert_eq!(
                    outer.cached_result, "prefix OLD suffix",
                    "{outer_form:?}/{child_form:?} {original}"
                );
                assert_eq!(
                    outer.cached_fields_in_source_order().len(),
                    1,
                    "{outer_form:?}/{child_form:?}"
                );
                let display = outer.cached_display_segments();
                assert_eq!(
                    display.iter().map(|(text, _)| *text).collect::<String>(),
                    "prefix OLD suffix"
                );
                assert!(display.iter().any(|(text, properties)| *text == "OLD"
                    && properties.is_some_and(|properties| properties.italic == Some(true))));
                let child = outer.cached_field_mut(0).unwrap();
                assert_eq!(child.form(), child_form);
                child.cached_result = "2".into();
                outer.refresh_cached_field_projection();
                let display = outer.cached_display_segments();
                assert_eq!(
                    display.iter().map(|(text, _)| *text).collect::<String>(),
                    "prefix 2 suffix"
                );
                assert!(display.iter().any(|(text, properties)| *text == "2"
                    && properties.is_some_and(|properties| properties.italic == Some(true))));
                let edited = serialized_paragraph(&reopened);
                assert!(edited.contains("<w:i"));
                let final_model =
                    parse_paragraph(&edited.replace("<w:p>", "").replace("</w:p>", ""));
                let field = parsed_field(&final_model, 0);
                assert_eq!(field.cached_result, "prefix 2 suffix");
                let children = field.cached_fields_in_source_order();
                assert_eq!(children.len(), 1);
                assert_eq!(children[0].form(), child_form);
                assert_eq!(children[0].cached_result, "2");
                assert_eq!(
                    children[0].cached_display_segments()[0].1.unwrap().italic,
                    Some(true)
                );
            }
        }
    }

    #[test]
    fn cached_result_fields_retain_typed_identity_and_producer_bytes() {
        let raw = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>UNKNOWN</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r data="cache"><w:rPr><w:b/></w:rPr><w:t>prefix </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>OLD</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t> suffix</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
        let mut paragraph = parse_paragraph(raw);
        assert!(serialized_paragraph(&paragraph).contains(raw));
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("outer");
        };
        assert_eq!(field.cached_fields_in_source_order().len(), 1);
        let child = field.cached_field_mut(0).unwrap();
        child.cached_result = "2".into();
        child.dirty = Some(false);
        field.refresh_cached_field_projection();
        assert_eq!(field.cached_result, "prefix 2 suffix");
        let output = serialized_paragraph(&paragraph);
        assert!(
            output.contains(r#"<w:r data="cache"><w:rPr><w:b/></w:rPr><w:t>prefix </w:t></w:r>"#)
        );
        assert_eq!(output.matches("fldCharType=\"begin\"").count(), 2);
        let reopened = parse_paragraph(&output.replace("<w:p>", "").replace("</w:p>", ""));
        let field = parsed_field(&reopened, 0);
        assert_eq!(field.cached_result, "prefix 2 suffix");
        assert_eq!(field.cached_fields_in_source_order()[0].cached_result, "2");
        assert_eq!(
            field.cached_fields_in_source_order()[0].cached_display_segments()[0]
                .1
                .unwrap()
                .italic,
            Some(true)
        );
    }
    #[test]
    fn replacement_child_cannot_silently_discard_a_parsed_simple_cache() {
        let raw = br#"<w:fldSimple xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:instr="UNKNOWN"><w:fldSimple w:instr="PAGE"><w:r><w:t>OLD</w:t></w:r></w:fldSimple></w:fldSimple>"#;
        let mut field = parse_simple_field(raw, &[]).unwrap().unwrap();
        *field.cached_field_mut(0).unwrap() = Field::new("PAGE", "replacement");
        let mut writer = Writer::new(Vec::new());
        let error = write_field(&mut writer, &field, None).unwrap_err();
        assert!(error.to_string().contains("source identity"));
    }
    #[test]
    fn typed_annotation_cache_reopens_as_field_with_sibling_range_markers() {
        for xml in [
            r#"<w:fldSimple w:instr=" REF Target \f " producer="keep"><w:r><w:t>OLD</w:t></w:r></w:fldSimple>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin" producer="keep"/></w:r><w:r><w:instrText xml:space="preserve"> REF Target \f </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/><w:t>OLD</w:t><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"<w:r producer="keep"><w:rPr><w:b/></w:rPr><w:t>PRE</w:t><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF Target \f </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>OLD</w:t><w:fldChar w:fldCharType="end"/><w:t>POST</w:t></w:r>"#,
        ] {
            let mut paragraph = parse_paragraph(xml);
            let field = paragraph
                .runs
                .iter_mut()
                .flat_map(|run| &mut run.content)
                .find_map(|content| {
                    if let RunContent::Field(field) = content {
                        Some(field)
                    } else {
                        None
                    }
                })
                .unwrap();
            field
                .set_cached_runs_with_comment_ranges(Vec::new(), Vec::new())
                .unwrap();
            assert_eq!(
                field.cached_result, "",
                "empty marker payload keeps the ordinary typed result contract"
            );
            let mut reference = CT_R::new("");
            reference.content = vec![RunContent::CommentReference {
                id: 17,
                raw_before: 0,
            }];
            let ranges = vec![
                CommentRangeMarker::Start {
                    id: 17,
                    run_index: 0,
                    raw_before: 0,
                    has_child_content: false,
                },
                CommentRangeMarker::End {
                    id: 17,
                    run_index: 1,
                    raw_before: 0,
                    has_child_content: false,
                },
            ];
            field
                .set_cached_runs_with_comment_ranges(vec![CT_R::new("LITERAL"), reference], ranges)
                .unwrap();
            let output = serialized_paragraph(&paragraph);
            assert!(output.contains("producer=\"keep\""));
            assert!(!output.contains("OLD"));
            let reopened = parse_paragraph(
                output
                    .strip_prefix("<w:p>")
                    .unwrap()
                    .strip_suffix("</w:p>")
                    .unwrap(),
            );
            let field = reopened
                .runs
                .iter()
                .flat_map(|run| &run.content)
                .find_map(|content| {
                    if let RunContent::Field(field) = content {
                        Some(field)
                    } else {
                        None
                    }
                })
                .expect("complex annotation cache remains typed");
            assert_eq!(field.cached_result, "LITERAL");
            assert_eq!(field.effective_instruction().raw, r"REF Target \f");
            assert_eq!(field.cached_result_comment_ranges().len(), 2);
            assert!(
                field
                    .cached_result_runs()
                    .unwrap()
                    .iter()
                    .flat_map(|run| &run.content)
                    .any(|content| matches!(content, RunContent::CommentReference { id: 17, .. }))
            );
            assert!(
                reopened.comment_ranges.is_empty(),
                "cache markers are owned by the field, never moved outside it"
            );
            assert_eq!(serialized_paragraph(&reopened), output);
            let mut reader = NsReader::from_reader(output.as_bytes());
            let mut stack = Vec::<Vec<u8>>::new();
            let mut buffer = Vec::new();
            loop {
                match reader.read_event_into(&mut buffer).unwrap() {
                    Event::Start(element) => stack.push(element.local_name().as_ref().to_vec()),
                    Event::End(_) => {
                        stack.pop();
                    }
                    Event::Empty(element)
                        if matches!(
                            element.local_name().as_ref(),
                            b"commentRangeStart" | b"commentRangeEnd"
                        ) =>
                    {
                        assert_ne!(stack.last().map(Vec::as_slice), Some(b"r".as_slice()))
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buffer.clear();
            }
        }
    }

    #[test]
    fn typed_annotation_marker_failures_leave_field_unchanged() {
        let mut reference = CT_R::new("");
        reference.content = vec![RunContent::CommentReference {
            id: 17,
            raw_before: 0,
        }];
        for ranges in [
            vec![CommentRangeMarker::Start {
                id: 17,
                run_index: 0,
                raw_before: 0,
                has_child_content: false,
            }],
            vec![
                CommentRangeMarker::Start {
                    id: 18,
                    run_index: 0,
                    raw_before: 0,
                    has_child_content: false,
                },
                CommentRangeMarker::End {
                    id: 18,
                    run_index: 1,
                    raw_before: 0,
                    has_child_content: false,
                },
            ],
            vec![
                CommentRangeMarker::Start {
                    id: 17,
                    run_index: 0,
                    raw_before: 1,
                    has_child_content: false,
                },
                CommentRangeMarker::End {
                    id: 17,
                    run_index: 1,
                    raw_before: 0,
                    has_child_content: false,
                },
            ],
            vec![
                CommentRangeMarker::Start {
                    id: 17,
                    run_index: 0,
                    raw_before: 0,
                    has_child_content: true,
                },
                CommentRangeMarker::End {
                    id: 17,
                    run_index: 1,
                    raw_before: 0,
                    has_child_content: false,
                },
            ],
            vec![
                CommentRangeMarker::Start {
                    id: 17,
                    run_index: 3,
                    raw_before: 0,
                    has_child_content: false,
                },
                CommentRangeMarker::End {
                    id: 17,
                    run_index: 3,
                    raw_before: 0,
                    has_child_content: false,
                },
            ],
        ] {
            let mut field = Field::new(r"REF Target \f", "OLD");
            let before = format!("{field:?}");
            assert!(
                field
                    .set_cached_runs_with_comment_ranges(
                        vec![CT_R::new("literal"), reference.clone()],
                        ranges
                    )
                    .is_err()
            );
            assert_eq!(format!("{field:?}"), before);
        }
        let mut field = Field::new(r"REF Target \f", "OLD");
        let before = format!("{field:?}");
        let mut raw = CT_R::new("literal");
        raw.extra_xml
            .push(br#"<w:commentRangeStart w:id="17"/>"#.to_vec());
        assert!(
            field
                .set_cached_runs_with_comment_ranges(vec![raw, reference], vec![])
                .is_err()
        );
        assert_eq!(format!("{field:?}"), before);
    }

    #[test]
    fn typed_annotation_nested_cache_keeps_raw_markers_and_parent_controls() {
        for outer in [FieldForm::Simple, FieldForm::Complex] {
            for inner in [FieldForm::Simple, FieldForm::Complex] {
                let child =
                    Field::from_raw(r"REF Target \f", inner, vec![CT_R::new("OLD")]).unwrap();
                let parent = Field::from_raw(
                    "UNKNOWN",
                    outer,
                    vec![CT_R::new("kept "), field_run(child, None)],
                )
                .unwrap();
                let mut paragraph = CT_P::new();
                paragraph.runs.push(field_run(parent, None));
                let original = serialized_paragraph(&paragraph);
                let mut paragraph = parse_paragraph(
                    original
                        .strip_prefix("<w:p>")
                        .unwrap()
                        .strip_suffix("</w:p>")
                        .unwrap(),
                );
                let RunContent::Field(parent) = &mut paragraph.runs[0].content[0] else {
                    panic!("parent");
                };
                let child = parent.cached_field_mut(0).unwrap();
                let mut reference = CT_R::new("");
                reference.content = vec![RunContent::CommentReference {
                    id: 29,
                    raw_before: 0,
                }];
                child
                    .set_cached_runs_with_comment_ranges(
                        vec![CT_R::new("ANNOTATED"), reference],
                        vec![
                            CommentRangeMarker::Start {
                                id: 29,
                                run_index: 0,
                                raw_before: 0,
                                has_child_content: false,
                            },
                            CommentRangeMarker::End {
                                id: 29,
                                run_index: 1,
                                raw_before: 0,
                                has_child_content: false,
                            },
                        ],
                    )
                    .unwrap();
                parent.refresh_cached_field_projection();
                let output = serialized_paragraph(&paragraph);
                assert!(!output.contains("OLD"));
                assert!(output.contains("UNKNOWN"));
                let output = output.replace(
                    "<w:commentRangeStart ",
                    "<w:commentRangeStart producer='verbatim' ",
                );
                let reopened = parse_paragraph(
                    output
                        .strip_prefix("<w:p>")
                        .unwrap()
                        .strip_suffix("</w:p>")
                        .unwrap(),
                );
                let parent = parsed_field(&reopened, 0);
                assert_eq!(parent.cached_result, "kept ANNOTATED");
                let child = parent.cached_fields_in_source_order()[0];
                assert_eq!(child.cached_result_comment_ranges().len(), 2);
                assert_eq!(child.cached_result, "ANNOTATED");
                assert!(reopened.comment_ranges.is_empty());
                assert_eq!(serialized_paragraph(&reopened), output);
            }
        }
    }

    #[test]
    fn comment_reference_carriers_preserve_source_and_follow_typed_ownership() {
        for paired in [false, true] {
            let reference = if paired {
                r#"<q:commentReference q:id="7" x:id="foreign" xmlns:w="urn:foreign"><x:owned/><!--owned--><?owned exact?></q:commentReference>"#
            } else {
                r#"<q:commentReference q:id="7" x:id="foreign" xmlns:w="urn:foreign"/>"#
            };
            let raw = format!(
                r#"<w:p xmlns:w="{}" xmlns:q="{}" xmlns:x="urn:opaque"><w:r x:run="kept"><!--before--><?before exact?><w:t>AB</w:t>{reference}<!--between--><w:t>CD</w:t><q:commentReference q:id="8" x:id="second"/><?after exact?></w:r></w:p>"#,
                crate::namespace::W_NS,
                crate::namespace::W_NS
            );
            let paragraph = CT_P::from_xml_fragment(raw.as_bytes()).unwrap();
            assert_eq!(paragraph.accepted_text(), "ABCD");
            assert_eq!(paragraph.accepted_run_paths().len(), 1);
            let mut run = paragraph.runs[0].clone();
            let serialize = |run: &CT_R| {
                let mut writer = Writer::new(Vec::new());
                run.to_xml(&mut writer).unwrap();
                String::from_utf8(writer.into_inner()).unwrap()
            };
            let output = serialize(&run);
            assert_eq!(output.matches("x:id=\"foreign\"").count(), 1, "{output}");
            assert_eq!(output.matches("x:id=\"second\"").count(), 1, "{output}");
            assert!(output.find("<!--before-->").unwrap() < output.find("<w:t>AB").unwrap());
            assert!(
                output.find("<!--between-->").unwrap() > output.find("x:id=\"foreign\"").unwrap()
            );
            assert!(output.contains("<?after exact?>"));
            if paired {
                assert!(output.contains("<x:owned/><!--owned--><?owned exact?>"));
            }
            run.ensure_properties().bold = Some(true);
            let output = serialize(&run);
            assert!(output.find("<w:rPr>").unwrap() < output.find("x:id=\"foreign\"").unwrap());
            let mut changed = run.clone();
            let RunContent::CommentReference { id, .. } = &mut changed.content[1] else {
                panic!("reference");
            };
            *id = 9;
            let changed = serialize(&changed);
            assert!(!changed.contains("x:id=\"foreign\""), "{changed}");
            assert!(changed.contains("w:id=\"9\""));
            let mut removed = run.clone();
            assert!(removed.remove_comment_references(&[7]));
            let removed = serialize(&removed);
            assert!(
                !removed.contains("foreign") && !removed.contains("<x:owned"),
                "{removed}"
            );
            assert!(removed.contains("<!--between-->") && removed.contains("x:id=\"second\""));
            let mut replaced = run.clone();
            replaced.replace_content(vec![RunContent::Text(CT_Text::new("NEW"))]);
            let replaced = serialize(&replaced);
            assert!(!replaced.contains("commentReference"), "{replaced}");
            assert!(replaced.contains("<!--before-->") && replaced.contains("<?after exact?>"));
            let mut head = run.clone();
            let tail = head.split_off_at(2).unwrap();
            assert!(!serialize(&head).contains("commentReference"));
            assert_eq!(serialize(&tail).matches("commentReference ").count(), 2);
            let mut segmented = Writer::new(Vec::new());
            write_run_content_segment(&mut segmented, &run, 0, 1, true, false, None).unwrap();
            write_run_content_segment(&mut segmented, &run, 1, 3, false, false, None).unwrap();
            write_run_content_segment(&mut segmented, &run, 3, 4, false, true, None).unwrap();
            let segmented = String::from_utf8(segmented.into_inner()).unwrap();
            assert_eq!(
                segmented.matches("x:id=\"foreign\"").count(),
                1,
                "{segmented}"
            );
            assert_eq!(
                segmented.matches("x:id=\"second\"").count(),
                1,
                "{segmented}"
            );
            let mut unordered = run.clone();
            unordered.extra_xml_positions.clear();
            // Without private boundary provenance, raw references are independent sources.
            let unordered = serialize(&unordered);
            assert_eq!(
                unordered.matches("commentReference ").count(),
                4,
                "{unordered}"
            );
            assert!(unordered.contains("x:id=\"foreign\""));
            assert!(unordered.contains("x:id=\"second\""));
        }
    }

    #[test]
    fn unordered_raw_comment_references_remain_independent_sources() {
        let raw = format!(
            r#"<w:commentReference xmlns:w="{}" w:id="7" xmlns:x="urn:x" x:keep="opaque"/>"#,
            crate::namespace::W_NS
        );
        let mut failures = Vec::new();
        for typed in [None, Some(7), Some(8)] {
            let run = CT_R {
                properties: None,
                content: typed
                    .map(|id| RunContent::CommentReference { id, raw_before: 1 })
                    .into_iter()
                    .collect(),
                extra_xml: vec![raw.as_bytes().to_vec()],
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            };
            let mut writer = Writer::new(Vec::new());
            run.to_xml(&mut writer).unwrap();
            let output = String::from_utf8(writer.into_inner()).unwrap();
            if !output.contains(&raw)
                || output.matches("commentReference ").count() != 1 + usize::from(typed.is_some())
            {
                failures.push(format!("typed={typed:?}: {output}"));
            }
        }
        assert!(
            failures.is_empty(),
            "independent raw references were lost: {failures:?}"
        );
    }
}
