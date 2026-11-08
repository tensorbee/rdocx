//! Ruby phonetic guides (`w:ruby`).
//!
//! `w:ruby` is a sibling of `w:r` inside `w:p`, not a run child. It carries a
//! phonetic line in `w:rt` and a base line in `w:rubyBase`, and both hold
//! ordinary runs, so [`CT_R`] is reused rather than a second run grammar
//! being introduced.
//!
//! The base runs live in `CT_P::runs` and the annotation records the
//! half-open span they occupy, exactly as `w:hyperlink` already does. That is
//! what makes text extraction, search, redaction and layout see the base text
//! as ordinary paragraph content, while the phonetic line, held here in
//! [`CT_Ruby::ruby_text`], stays an annotation that no text extraction
//! returns.
//!
//! The typed model admits only the shapes whose projected geometry it retains.
//! Retained raw fragments gain only necessary root namespace declarations at
//! initial capture, completing their original context against canonical Word
//! publication. Their QName, existing attribute and descendant payload bytes
//! remain intact. A
//! `w:ruby` carrying an attribute, a `w:rt` or `w:rubyBase` child that is not
//! a run, non-whitespace character data between children, or an empty base
//! line is left as preserved raw XML by the paragraph parser instead.
//! Whitespace between children is layout indentation that this crate's own
//! writer produces, so it is read and dropped rather than rejected.
//!
//! The one exception is an unmodelled `w:rubyPr` element child, which is
//! retained and written after the modelled ones, because `w:rubyPr` is a
//! closed sequence and such a child is already off-schema wherever it sat.

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::numbering::word_prefixes_at;
use crate::properties::{get_word_val_attr, is_word_element, parse_word_toggle};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::text::{CT_R, parse_run_raw};
use crate::units::HalfPoint;

/// `ST_RubyAlign` — the horizontal distribution of the phonetic line over the
/// base line, on `w:rubyAlign/@w:val`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_RubyAlign {
    Center,
    DistributeLetter,
    DistributeSpace,
    Left,
    RightVertical,
    Right,
    /// A token outside the ECMA-376 inventory, retained verbatim.
    Other(String),
}

impl ST_RubyAlign {
    pub fn from_str(value: &str) -> Self {
        match value {
            "center" => ST_RubyAlign::Center,
            "distributeLetter" => ST_RubyAlign::DistributeLetter,
            "distributeSpace" => ST_RubyAlign::DistributeSpace,
            "left" => ST_RubyAlign::Left,
            "rightVertical" => ST_RubyAlign::RightVertical,
            "right" => ST_RubyAlign::Right,
            other => ST_RubyAlign::Other(other.to_owned()),
        }
    }

    pub fn as_str(&self) -> &str {
        match self {
            ST_RubyAlign::Center => "center",
            ST_RubyAlign::DistributeLetter => "distributeLetter",
            ST_RubyAlign::DistributeSpace => "distributeSpace",
            ST_RubyAlign::Left => "left",
            ST_RubyAlign::RightVertical => "rightVertical",
            ST_RubyAlign::Right => "right",
            ST_RubyAlign::Other(value) => value.as_str(),
        }
    }
}

/// `CT_RubyPr` — the geometry of one phonetic guide.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub struct CT_RubyPr {
    /// Distribution of the phonetic line over the base (`w:rubyAlign`).
    pub align: Option<ST_RubyAlign>,
    /// Phonetic line size in half-points (`w:hps`).
    pub hps: Option<HalfPoint>,
    /// Phonetic baseline raise above the base baseline in half-points
    /// (`w:hpsRaise`).
    pub hps_raise: Option<HalfPoint>,
    /// Base line size in half-points (`w:hpsBaseText`).
    pub hps_base_text: Option<HalfPoint>,
    /// Language identifier for the phonetic line (`w:lid/@w:val`).
    pub language: Option<String>,
    /// Producer-set dirty flag (`w:dirty`).
    pub dirty: Option<bool>,
    /// Unmodelled `w:rubyPr` children, written after the modelled ones.
    /// Initial capture may add needed root xmlns declarations to retain their
    /// original namespace context under canonical Word publication.
    #[doc(hidden)]
    pub raw_xml: Vec<Vec<u8>>,
}

/// Shared Start/Empty action of the actual Ruby root reader.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum RubyChildAction {
    Properties,
    Text,
    Base,
    Raw,
    Reject,
}
impl RubyChildAction {
    pub(crate) fn for_names(local: &[u8], word: bool, empty: bool) -> Self {
        if !word {
            return Self::Raw;
        }
        match local {
            b"rubyPr" => Self::Properties,
            b"rt" if !empty => Self::Text,
            b"rubyBase" if !empty => Self::Base,
            b"rt" | b"rubyBase" => Self::Reject,
            _ => Self::Raw,
        }
    }
}

/// Actual run-container action, shared by its reader and source qualification.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum RubyContentAction {
    Run,
    Reject,
}

impl RubyContentAction {
    pub(crate) fn for_names(local: &[u8], word: bool) -> Self {
        if word && local == b"r" {
            Self::Run
        } else {
            Self::Reject
        }
    }
}

/// Character events that force an exact raw fallback at the current Ruby reader.
pub(crate) fn rejects_direct_characters(event: &Event<'_>) -> bool {
    match event {
        Event::Text(text) => !text.iter().all(u8::is_ascii_whitespace),
        Event::CData(_) => true,
        _ => false,
    }
}

impl CT_RubyPr {
    /// Parse `w:rubyPr` from its captured source.
    ///
    /// `None` reports non-whitespace character data this model cannot write
    /// back exactly, so the whole annotation stays raw.
    fn from_raw(raw: &[u8], word_prefixes: &[String]) -> Result<Option<Self>> {
        let mut reader = Reader::from_reader(raw);
        reader.config_mut().trim_text(false);
        let mut properties = CT_RubyPr::default();
        let mut word_prefixes = word_prefixes.to_vec();
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        loop {
            let mut child = None;
            match reader.read_event_into(&mut buffer)? {
                Event::Start(element) => {
                    if depth == 0 {
                        word_prefixes = word_prefixes_at(&element, &word_prefixes)?.into_owned();
                        depth = 1;
                    } else {
                        let prefixes = word_prefixes_at(&element, &word_prefixes)?;
                        let raw = capture_element(&mut reader, &element)?;
                        child = Some((element.into_owned(), prefixes, Some(raw)));
                    }
                }
                Event::Empty(element) => {
                    if depth == 0 {
                        return Ok(Some(properties));
                    }
                    let prefixes = word_prefixes_at(&element, &word_prefixes)?;
                    child = Some((element.into_owned(), prefixes, None));
                }
                event if depth == 1 && rejects_direct_characters(&event) => return Ok(None),
                Event::End(_) => depth = depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
            if let Some((element, prefixes, raw)) = child
                && !properties.read_child(&element, &prefixes)?
            {
                properties.raw_xml.push(match raw {
                    Some(raw) => raw,
                    None => capture_empty_element(&element)?,
                });
            }
            buffer.clear();
        }
        complete_fragments(&mut properties.raw_xml, &word_prefixes, raw.len())?;
        Ok(Some(properties))
    }

    /// Record one modelled `w:rubyPr` child, reporting whether it was modelled.
    ///
    /// Every child but `w:dirty` carries a required `w:val`. One without it
    /// names no value, so it is reported unmodelled and kept as raw rather
    /// than normalised away, which is the policy `w:kern` already follows in
    /// the run property module.
    fn read_child(&mut self, element: &BytesStart<'_>, prefixes: &[String]) -> Result<bool> {
        match RubyPropertyChild::read(element, prefixes)? {
            RubyPropertyChild::Align(value) => self.align = Some(value),
            RubyPropertyChild::Size(value) => self.hps = Some(value),
            RubyPropertyChild::Raise(value) => self.hps_raise = Some(value),
            RubyPropertyChild::BaseSize(value) => self.hps_base_text = Some(value),
            RubyPropertyChild::Language(value) => self.language = Some(value),
            RubyPropertyChild::Dirty(value) => self.dirty = Some(value),
            RubyPropertyChild::Raw => return Ok(false),
        }
        Ok(true)
    }

    /// Write `w:rubyPr` in `xsd:sequence` order with the fixed `w:` prefix.
    fn to_xml<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<()> {
        let mut start = crate::text::FieldXmlStart::new("w:rubyPr", emission)?;
        let prefix = if role == crate::text::XmlWriteRole::Source {
            start.qualify(&[], role, source_context)?
        } else {
            "w".to_owned()
        };
        start.write(writer, false)?;
        if let Some(ref align) = self.align {
            let mut element = crate::text::FieldXmlStart::word(&prefix, "rubyAlign", emission)?;
            element.push_word(&prefix, "val", align.as_str())?;
            element.write(writer, true)?;
        }
        let mut buffer = itoa::Buffer::new();
        for (tag, value) in [
            ("hps", self.hps),
            ("hpsRaise", self.hps_raise),
            ("hpsBaseText", self.hps_base_text),
        ] {
            if let Some(value) = value {
                let mut element = crate::text::FieldXmlStart::word(&prefix, tag, emission)?;
                element.push_word(&prefix, "val", buffer.format(value.0))?;
                element.write(writer, true)?;
            }
        }
        if let Some(ref language) = self.language {
            let mut element = crate::text::FieldXmlStart::word(&prefix, "lid", emission)?;
            element.push_word(&prefix, "val", language.as_str())?;
            element.write(writer, true)?;
        }
        if let Some(dirty) = self.dirty {
            let mut element = crate::text::FieldXmlStart::new("w:dirty", emission)?;
            let dirty_prefix = if role == crate::text::XmlWriteRole::Source {
                element.qualify(&[], role, source_context)?
            } else {
                "w".to_owned()
            };
            if !dirty {
                element.push_word(&dirty_prefix, "val", "false")?;
            }
            element.write(writer, true)?;
        }
        for raw in &self.raw_xml {
            if role == crate::text::XmlWriteRole::Source {
                let mut destination = source_context.unwrap_or(&[]).to_vec();
                destination.retain(|(known, _)| *known != prefix);
                destination.push((&prefix, crate::namespace::W_NS));
                let original = ["w".to_owned()];
                let budget = crate::text::retained_owner_work(raw.len(), &original, 0)?;
                writer
                    .get_mut()
                    .write_all(&crate::text::retained_raw_roots(
                        raw,
                        &original,
                        &destination,
                        budget,
                    )?)?;
            } else {
                writer.get_mut().write_all(raw)?;
            }
        }
        if let Some(emission) = emission {
            emission.write_word_end(writer, &prefix, "rubyPr")?;
        } else {
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:rubyPr"))))?;
        }
        Ok(())
    }
}

/// Report whether the element carries an attribute other than a namespace
/// declaration.
///
/// `w:ruby`, `w:rt` and `w:rubyBase` have no attributes of their own, so one
/// that is not `xmlns` is content this model cannot write back exactly.
pub(crate) fn has_foreign_attributes(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes() {
        let key = attribute?.key;
        let key = key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            return Ok(true);
        }
    }
    Ok(false)
}

#[cfg(test)]
thread_local! {
    pub(crate) static RUBY_PROPERTY_READS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static RUBY_PROPERTY_LANGUAGE_BYTES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// Actual property scalar outcome, shared with source constructor qualification.
/// Started children reach this only after their complete capture succeeds.
pub(crate) enum RubyPropertyChild {
    Align(ST_RubyAlign),
    Size(HalfPoint),
    Raise(HalfPoint),
    BaseSize(HalfPoint),
    Language(String),
    Dirty(bool),
    Raw,
}

impl RubyPropertyChild {
    pub(crate) fn read(element: &BytesStart<'_>, prefixes: &[String]) -> Result<Self> {
        #[cfg(test)]
        RUBY_PROPERTY_READS.set(RUBY_PROPERTY_READS.get() + 1);
        let name = element.name();
        let name = name.as_ref();
        Ok(if is_word_element(name, b"rubyAlign", prefixes) {
            get_word_val_attr(element, prefixes)?
                .map(|value| Self::Align(ST_RubyAlign::from_str(&value)))
                .unwrap_or(Self::Raw)
        } else if is_word_element(name, b"hps", prefixes) {
            half_point_value(element, prefixes)?
                .map(Self::Size)
                .unwrap_or(Self::Raw)
        } else if is_word_element(name, b"hpsRaise", prefixes) {
            half_point_value(element, prefixes)?
                .map(Self::Raise)
                .unwrap_or(Self::Raw)
        } else if is_word_element(name, b"hpsBaseText", prefixes) {
            half_point_value(element, prefixes)?
                .map(Self::BaseSize)
                .unwrap_or(Self::Raw)
        } else if is_word_element(name, b"lid", prefixes) {
            get_word_val_attr(element, prefixes)?
                .map(|value| {
                    #[cfg(test)]
                    RUBY_PROPERTY_LANGUAGE_BYTES
                        .set(RUBY_PROPERTY_LANGUAGE_BYTES.get() + value.len());
                    Self::Language(value)
                })
                .unwrap_or(Self::Raw)
        } else if is_word_element(name, b"dirty", prefixes) {
            Self::Dirty(parse_word_toggle(element, prefixes)?)
        } else {
            Self::Raw
        })
    }
}

/// Read a `w:val` half-point attribute, leaving the slot empty when the
/// required attribute is absent or is not a number.
fn half_point_value(element: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<HalfPoint>> {
    Ok(get_word_val_attr(element, prefixes)?
        .and_then(|value| value.parse().ok())
        .map(HalfPoint))
}

/// `CT_Ruby` — one phonetic guide over a span of paragraph runs.
#[derive(Debug, Clone, Default, PartialEq)]
#[allow(non_camel_case_types)]
pub struct CT_Ruby {
    /// The phonetic geometry (`w:rubyPr`).
    pub properties: Option<CT_RubyPr>,
    /// The phonetic runs (`w:rt`). No text extraction returns these.
    pub ruby_text: Vec<CT_R>,
    /// Index of the first base run in the owning paragraph (inclusive).
    pub base_start: usize,
    /// Index one past the last base run in the owning paragraph (exclusive).
    pub base_end: usize,
    /// Unmodelled `w:ruby` children, written after `w:rubyBase`.
    /// Initial capture may add needed root xmlns declarations to retain their
    /// original namespace context under canonical Word publication.
    #[doc(hidden)]
    pub raw_xml: Vec<Vec<u8>>,
}

impl CT_Ruby {
    /// Parse `w:ruby` from its captured source.
    ///
    /// Returns the annotation, with `base_start` and `base_end` left at zero
    /// for the owning paragraph to fill, together with the base runs the
    /// paragraph splices into its own run list. `None` means the source
    /// carries something this model cannot write back exactly, so the caller
    /// keeps the raw subtree instead.
    pub(crate) fn from_raw(
        raw: &[u8],
        word_prefixes: &[String],
    ) -> Result<Option<(Self, Vec<CT_R>)>> {
        Self::from_raw_with_texts(raw, word_prefixes, None, None)
    }

    pub(crate) fn from_raw_with_texts(
        raw: &[u8],
        word_prefixes: &[String],
        input: Option<&crate::document::WordFieldInput>,
        text_facts: Option<&mut crate::text::WordParagraphTextFacts>,
    ) -> Result<Option<(Self, Vec<CT_R>)>> {
        let mut reader = Reader::from_reader(raw);
        reader.config_mut().trim_text(false);
        let mut ruby = CT_Ruby::default();
        let cursor = input.map(|input| input.cursor(raw)).transpose()?;
        let mut base_facts = Vec::new();
        let mut base_runs = Vec::new();
        let mut seen_base = false;
        let mut word_prefixes = word_prefixes.to_vec();
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        loop {
            let event_start = reader.buffer_position();
            match reader.read_event_into(&mut buffer)? {
                Event::Start(element) => {
                    if depth == 0 {
                        if has_foreign_attributes(&element)? {
                            return Ok(None);
                        }
                        word_prefixes = word_prefixes_at(&element, &word_prefixes)?.into_owned();
                        depth = 1;
                    } else {
                        let prefixes = word_prefixes_at(&element, &word_prefixes)?;
                        let name = element.name().as_ref().to_vec();
                        let child = capture_element(&mut reader, &element)?;
                        let local = element.local_name();
                        let action = RubyChildAction::for_names(
                            local.as_ref(),
                            is_word_element(&name, local.as_ref(), &prefixes),
                            false,
                        );
                        if action == RubyChildAction::Properties {
                            let Some(properties) = CT_RubyPr::from_raw(&child, &prefixes)? else {
                                return Ok(None);
                            };
                            ruby.properties = Some(properties);
                        } else if action == RubyChildAction::Text {
                            let Some(runs) = parse_ruby_content(&child, &prefixes, None, None)?
                            else {
                                return Ok(None);
                            };
                            ruby.ruby_text = runs;
                        } else if action == RubyChildAction::Base {
                            let mapped = cursor.as_ref().and_then(|cursor| {
                                cursor.field_input(
                                    &reader,
                                    event_start..reader.buffer_position(),
                                    &child,
                                )
                            });
                            let mut facts = Vec::new();
                            let Some(runs) = parse_ruby_content(
                                &child,
                                &prefixes,
                                mapped.as_ref(),
                                text_facts.as_ref().map(|_| &mut facts),
                            )?
                            else {
                                return Ok(None);
                            };
                            base_facts = facts;
                            seen_base = true;
                            base_runs = runs;
                        } else {
                            ruby.raw_xml.push(child);
                        }
                    }
                }
                Event::Empty(element) => {
                    if depth == 0 {
                        // `<w:ruby/>` carries no base line to annotate.
                        return Ok(None);
                    }
                    let prefixes = word_prefixes_at(&element, &word_prefixes)?;
                    let name = element.name();
                    let local = element.local_name();
                    let action = RubyChildAction::for_names(
                        local.as_ref(),
                        is_word_element(name.as_ref(), local.as_ref(), &prefixes),
                        true,
                    );
                    if action == RubyChildAction::Properties {
                        ruby.properties = Some(CT_RubyPr::default());
                    } else if action == RubyChildAction::Reject {
                        // An empty base line has nothing to annotate, and an
                        // empty phonetic line would not survive the write
                        // back as an empty element.
                        return Ok(None);
                    } else {
                        ruby.raw_xml.push(capture_empty_element(&element)?);
                    }
                }
                event if depth == 1 && rejects_direct_characters(&event) => return Ok(None),
                Event::End(_) => depth = depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        if !seen_base || base_runs.is_empty() {
            return Ok(None);
        }
        complete_fragments(&mut ruby.raw_xml, &word_prefixes, raw.len())?;
        if let Some(output) = text_facts {
            *output = base_facts;
        }
        Ok(Some((ruby, base_runs)))
    }

    /// The half-open range of owning-paragraph runs this annotation covers.
    pub fn base_range(&self) -> std::ops::Range<usize> {
        self.base_start..self.base_end
    }

    /// Write `w:ruby` up to and including the `w:rubyBase` start tag.
    ///
    /// The base runs are paragraph runs, so the owning paragraph writes them
    /// between this call and [`Self::write_end`].
    pub(crate) fn write_start<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<crate::text::RootBindingScope> {
        let mut start = crate::text::FieldXmlStart::new("w:ruby", emission)?;
        let prefix = if role == crate::text::XmlWriteRole::Source {
            start.qualify(&[], role, source_context)?
        } else {
            start.restore_word_binding()?;
            "w".to_owned()
        };
        let scope = crate::text::emitted_binding_scope(start.element(), false, false)?;
        start.write(writer, false)?;
        let mut context = source_context.unwrap_or(&[]).to_vec();
        context.retain(|(known, _)| *known != prefix);
        context.push((&prefix, crate::namespace::W_NS));
        let child_context =
            (role == crate::text::XmlWriteRole::Source).then_some(context.as_slice());
        if let Some(properties) = &self.properties {
            properties.to_xml(writer, role, child_context, emission)?;
        }
        crate::text::FieldXmlStart::word(&prefix, "rt", emission)?.write(writer, false)?;
        for run in &self.ruby_text {
            run.write_xml_role(writer, None, role, child_context, emission, None)?;
        }
        if let Some(emission) = emission {
            emission.write_word_end(writer, &prefix, "rt")?;
        } else {
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:rt"))))?;
        }
        crate::text::FieldXmlStart::word(&prefix, "rubyBase", emission)?.write(writer, false)?;
        Ok(scope)
    }

    /// Close the same actual Ruby root and locally complete only the raw-vector
    /// contract. Retained Drawing payloads keep their independent original scope.
    pub(crate) fn write_end<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<()> {
        let prefix = if let Some(emission) = emission {
            let plan = if role == crate::text::XmlWriteRole::Source {
                crate::text::generated_word_name(source_context.unwrap_or(&[]))?
            } else {
                crate::text::GeneratedWordName::fixed("w")
            };
            emission.write_word_end_planned(writer, plan, "rubyBase")?
        } else {
            let prefix = if role == crate::text::XmlWriteRole::Source {
                crate::text::generated_word_start(
                    BytesStart::new("w:ruby"),
                    &[],
                    role,
                    source_context,
                )?
                .1
            } else {
                "w".to_owned()
            };
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:rubyBase"))))?;
            prefix
        };
        let mut destination = source_context.unwrap_or(&[]).to_vec();
        destination.retain(|(known, _)| *known != prefix);
        destination.push((&prefix, crate::namespace::W_NS));
        for raw in &self.raw_xml {
            if role == crate::text::XmlWriteRole::Source {
                let original = ["w".to_owned()];
                let budget = crate::text::retained_owner_work(raw.len(), &original, 0)?;
                writer
                    .get_mut()
                    .write_all(&crate::text::retained_raw_roots(
                        raw,
                        &original,
                        &destination,
                        budget,
                    )?)?;
            } else {
                writer.get_mut().write_all(raw)?;
            }
        }
        if let Some(emission) = emission {
            emission.write_word_end(writer, &prefix, "ruby")?;
        } else {
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:ruby"))))?;
        }
        Ok(())
    }
}

/// Parse the runs of a `w:rt` or `w:rubyBase`, or report that the element
/// carries something the typed model cannot write back exactly.
// Capture completion is independent of current publication globals. Canonical
// Word is the existing consumer contract for these public retained vectors.
fn complete_fragments(
    fragments: &mut [Vec<u8>],
    original: &[String],
    source_bytes: usize,
) -> Result<()> {
    let budget = crate::text::retained_owner_work(source_bytes, original, 0)?;
    let destination = [("w", crate::namespace::W_NS)];
    let mut projected = source_bytes;
    let mut plans = Vec::with_capacity(fragments.len());
    for fragment in fragments.iter() {
        let plan = crate::text::retained_root_plan(fragment, original, &destination, budget)?;
        projected = projected
            .checked_add(plan.added_bytes(fragment))
            .filter(|bytes| *bytes <= budget)
            .ok_or_else(|| {
                crate::OxmlError::InvalidValue(
                    "Ruby namespace completion exceeds affected-source work budget".into(),
                )
            })?;
        plans.push(plan);
    }
    for (fragment, plan) in fragments.iter_mut().zip(plans) {
        if plan.added_bytes(fragment) != 0 {
            *fragment = crate::text::emit_retained_roots(fragment, plan)?.into_owned();
        }
    }
    Ok(())
}

fn parse_ruby_content(
    raw: &[u8],
    word_prefixes: &[String],
    input: Option<&crate::document::WordFieldInput>,
    text_facts: Option<&mut crate::text::WordParagraphTextFacts>,
) -> Result<Option<Vec<CT_R>>> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let cursor = input.map(|input| input.cursor(raw)).transpose()?;
    let mut facts = Vec::new();
    let mut runs = Vec::new();
    let mut word_prefixes = word_prefixes.to_vec();
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let event_start = reader.buffer_position();
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                if depth == 0 {
                    if has_foreign_attributes(&element)? {
                        return Ok(None);
                    }
                    word_prefixes = word_prefixes_at(&element, &word_prefixes)?.into_owned();
                    depth = 1;
                } else {
                    let prefixes = word_prefixes_at(&element, &word_prefixes)?;
                    if RubyContentAction::for_names(
                        element.local_name().as_ref(),
                        is_word_element(element.name().as_ref(), b"r", &prefixes),
                    ) == RubyContentAction::Reject
                    {
                        return Ok(None);
                    }
                    let child = capture_element(&mut reader, &element)?;
                    let mut parsed_texts = Vec::new();
                    runs.push(parse_run_raw(
                        &child,
                        &prefixes,
                        text_facts.as_ref().map(|_| &mut parsed_texts),
                    )?);
                    if text_facts.is_some() {
                        facts.push(
                            cursor
                                .as_ref()
                                .and_then(|cursor| {
                                    cursor.captured_map(
                                        &reader,
                                        event_start..reader.buffer_position(),
                                        &child,
                                    )
                                })
                                .map(|source_map| crate::text::ParsedWordRunTexts {
                                    source: 0..child.len(),
                                    source_map: Some(source_map),
                                    texts: parsed_texts,
                                }),
                        );
                    }
                }
            }
            Event::Empty(element) => {
                if depth == 0 {
                    return Ok(None);
                }
                let prefixes = word_prefixes_at(&element, &word_prefixes)?;
                if RubyContentAction::for_names(
                    element.local_name().as_ref(),
                    is_word_element(element.name().as_ref(), b"r", &prefixes),
                ) == RubyContentAction::Reject
                {
                    return Ok(None);
                }
                runs.push(CT_R::from_empty_root(&element, &prefixes)?);
                if text_facts.is_some() {
                    facts.push(None);
                }
            }
            event if depth == 1 && rejects_direct_characters(&event) => return Ok(None),
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if let Some(output) = text_facts {
        *output = facts;
    }
    Ok(Some(runs))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::namespace::W_NS;

    fn parse(inner: &str) -> Option<(CT_Ruby, Vec<CT_R>)> {
        let source = format!(r#"<w:ruby xmlns:w="{W_NS}">{inner}</w:ruby>"#);
        CT_Ruby::from_raw(source.as_bytes(), &["w".to_owned()]).unwrap()
    }

    #[test]
    fn a_full_ruby_parses_its_properties_its_phonetic_runs_and_its_base_runs() {
        let (ruby, base) = parse(concat!(
            r#"<w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/>"#,
            r#"<w:hpsRaise w:val="22"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/>"#,
            r#"<w:dirty w:val="false"/></w:rubyPr>"#,
            r#"<w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt>"#,
            r#"<w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase>"#,
        ))
        .expect("a well formed ruby is modelled");
        let properties = ruby.properties.expect("w:rubyPr is modelled");
        assert_eq!(properties.align, Some(ST_RubyAlign::DistributeSpace));
        assert_eq!(properties.hps, Some(HalfPoint(10)));
        assert_eq!(properties.hps_raise, Some(HalfPoint(22)));
        assert_eq!(properties.hps_base_text, Some(HalfPoint(21)));
        assert_eq!(properties.language.as_deref(), Some("ja-JP"));
        assert_eq!(properties.dirty, Some(false));
        assert_eq!(ruby.ruby_text.len(), 1);
        assert_eq!(ruby.ruby_text[0].text(), "かんじ");
        assert_eq!(base.len(), 1);
        assert_eq!(base[0].text(), "漢字");
    }

    #[test]
    fn a_prefix_aliased_ruby_reads_through_its_own_alias() {
        let source = format!(
            concat!(
                r#"<q:ruby xmlns:q="{ns}"><q:rt><q:r><q:t>よみ</q:t></q:r></q:rt>"#,
                r#"<q:rubyBase><q:r><q:t>読</q:t></q:r></q:rubyBase></q:ruby>"#,
            ),
            ns = W_NS
        );
        let (ruby, base) = CT_Ruby::from_raw(source.as_bytes(), &[])
            .unwrap()
            .expect("an aliased ruby is modelled");
        assert_eq!(ruby.ruby_text[0].text(), "よみ");
        assert_eq!(base[0].text(), "読");
    }

    #[test]
    fn a_ruby_without_a_base_line_is_not_modelled() {
        assert!(parse(r#"<w:rt><w:r><w:t>よみ</w:t></w:r></w:rt>"#).is_none());
        assert!(parse(r#"<w:rt><w:r><w:t>よみ</w:t></w:r></w:rt><w:rubyBase/>"#).is_none());
    }

    #[test]
    fn a_ruby_content_child_that_is_not_a_run_is_not_modelled() {
        assert!(
            parse(concat!(
                r#"<w:rt><w:r><w:t>よみ</w:t></w:r></w:rt>"#,
                r#"<w:rubyBase><w:sdt><w:sdtContent/></w:sdt></w:rubyBase>"#,
            ))
            .is_none()
        );
    }

    /// A `w:rubyPr` child without its required `w:val` names no value, so it
    /// is kept verbatim rather than normalised into an empty modelled slot.
    #[test]
    fn a_ruby_property_child_without_its_required_value_stays_raw() {
        let (ruby, _) = parse(concat!(
            r#"<w:rubyPr><w:hps/><w:rubyAlign/></w:rubyPr>"#,
            r#"<w:rt><w:r><w:t>よみ</w:t></w:r></w:rt>"#,
            r#"<w:rubyBase><w:r><w:t>読</w:t></w:r></w:rubyBase>"#,
        ))
        .expect("a valueless property child does not reject the ruby");
        let properties = ruby.properties.as_ref().expect("w:rubyPr is modelled");
        assert_eq!(properties.hps, None);
        assert_eq!(properties.align, None);
        assert_eq!(properties.raw_xml.len(), 2);
    }

    #[test]
    fn unmodelled_ruby_property_children_are_retained_and_written_back() {
        let (ruby, _) = parse(concat!(
            r#"<w:rubyPr><w:hps w:val="10"/><w:producerOnly w:val="1"/></w:rubyPr>"#,
            r#"<w:rt><w:r><w:t>よみ</w:t></w:r></w:rt>"#,
            r#"<w:rubyBase><w:r><w:t>読</w:t></w:r></w:rubyBase>"#,
        ))
        .expect("an unmodelled property child does not reject the ruby");
        let properties = ruby.properties.as_ref().expect("w:rubyPr is modelled");
        assert_eq!(properties.raw_xml.len(), 1);

        let mut output = Vec::new();
        properties
            .to_xml(
                &mut Writer::new(&mut output),
                crate::text::XmlWriteRole::Publication,
                None,
                None,
            )
            .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(
            output.contains(r#"<w:producerOnly w:val="1"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"<w:hps w:val="10"/>"#), "{output}");
    }
}
