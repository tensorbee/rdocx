//! Structured template parsing and JSON evaluation.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};

use quick_xml::Reader;
use quick_xml::events::Event;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::namespace::matches_local_name;
use rdocx_oxml::placeholder;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, RunContent};
use serde_json::Value;

use crate::document::Document;
use crate::error::{Error, Result};

/// The explicitly selected stable template grammar.
pub const TEMPLATE_GRAMMAR_VERSION: u32 = 1;

/// Original physical paragraph and expanded scalar occurrence supplied to a formatter.
pub struct TemplateFormatContext<'a> {
    pub location: &'a crate::ContentLocation,
    pub path: &'a str,
    pub occurrence: usize,
}

// Source kind, source part, old id, destination part, new id.
type NoteCopy = (crate::StoryKind, String, i32, String, i32);

// Part name, original bytes, rendered bytes.
type PartRewrite = (String, Vec<u8>, Vec<u8>);
type ScopedPartRewrite = (String, Vec<u8>, rdocx_oxml::document::ScopedWordSource);

// Prepared grammar state for one operation, derived by the existing physical
// story inventory. Ranges remain original source coordinates, never authority
// for a later document or for an imported item's changed byte offsets.
#[derive(Clone)]
pub(crate) struct TemplateProgram {
    pub story: crate::StoryId,
    pub part: String,
    pub ranges: Vec<std::ops::Range<usize>>,
    pub raw: Vec<u8>,
    pub source: Option<std::sync::Arc<rdocx_oxml::document::ScopedWordSource>>,
    pub scope: BTreeMap<String, String>,
    pub required_paragraph: bool,
    pub table_rows: bool,
    pub items: Vec<TemplateProgramItem>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum TemplateItemKind {
    Source(crate::StoryItemKind),
    TableRow,
}

#[derive(Clone)]
pub(crate) struct TemplateProgramItem {
    pub location: crate::ContentLocation,
    pub range: std::ops::Range<usize>,
    pub kind: TemplateItemKind,
    pub field_owned: bool,
    pub marker_unsafe: bool,
    pub raw: Vec<u8>,
    pub scope: BTreeMap<String, String>,
    pub children: Vec<TemplateProgram>,
}

#[derive(Debug, Clone)]
enum PathStep {
    Key(String),
    Index(usize),
}

#[derive(Debug)]
enum Filter {
    Trim,
    Upper,
    Lower,
    Default(String),
    Join(String),
}

#[derive(Debug)]
struct Expression {
    spelling: String,
    path: Vec<PathStep>,
    filters: Vec<Filter>,
}

#[derive(Debug)]
enum ScalarValue {
    Expression(Expression),
    Literal(&'static str),
}

#[derive(Debug)]
struct ScalarToken {
    range: std::ops::Range<usize>,
    value: ScalarValue,
}

fn scalar_tokens(source: &str, versioned: bool) -> Result<Vec<ScalarToken>> {
    let mut tokens = Vec::new();
    let mut cursor = 0;
    while cursor < source.len() {
        if versioned {
            let escape = [
                ("{{{{", "{{"),
                ("{%{%", "{%"),
                ("}}}}", "}}"),
                ("%}%}", "%}"),
            ]
            .into_iter()
            .find(|(literal, _)| source[cursor..].starts_with(literal));
            if let Some((literal, value)) = escape {
                tokens.push(ScalarToken {
                    range: cursor..cursor + literal.len(),
                    value: ScalarValue::Literal(value),
                });
                cursor += literal.len();
                continue;
            }
        }
        if source[cursor..].starts_with("{{") {
            let start = cursor;
            cursor += 2;
            let inner_start = cursor;
            let mut quoted = false;
            let mut escaped = false;
            loop {
                let Some(character) = source[cursor..].chars().next() else {
                    return Err(template_error("unclosed scalar tag"));
                };
                if versioned && quoted {
                    if escaped {
                        escaped = false;
                    } else if character == '\\' {
                        escaped = true;
                    } else if character == '"' {
                        quoted = false;
                    }
                } else if versioned && character == '"' {
                    quoted = true;
                } else if source[cursor..].starts_with("}}") {
                    let inner = &source[inner_start..cursor];
                    let expression = if versioned {
                        parse_expression(inner)?
                    } else {
                        Expression {
                            spelling: inner.trim().into(),
                            path: parse_path(inner)?
                                .into_iter()
                                .map(|key| PathStep::Key(key.into()))
                                .collect(),
                            filters: Vec::new(),
                        }
                    };
                    cursor += 2;
                    tokens.push(ScalarToken {
                        range: start..cursor,
                        value: ScalarValue::Expression(expression),
                    });
                    break;
                } else if source[cursor..].starts_with("{{")
                    || source[cursor..].starts_with("{%")
                    || source[cursor..].starts_with("%}")
                {
                    return Err(template_error("nested or mismatched scalar tag"));
                }
                cursor += character.len_utf8();
            }
        } else if ["{%", "}}", "%}"]
            .iter()
            .any(|marker| source[cursor..].starts_with(marker))
        {
            return Err(template_error(
                "structural or unmatched delimiter in scalar text",
            ));
        } else {
            cursor += source[cursor..]
                .chars()
                .next()
                .expect("cursor is before end")
                .len_utf8();
        }
    }
    Ok(tokens)
}

struct ParsedProgram<'source> {
    source: &'source TemplateProgram,
    blocks: Vec<Block<usize>>,
    scalars: Vec<Vec<(usize, ScalarToken)>>,
    children: Vec<Vec<ParsedProgram<'source>>>,
}

// Stretch, token range, rendered text, formatting, actual borrowed operands,
// and the genuine callback delivery event carried through execution copies.
type PreparedTemplateScalar<'data> = (
    usize,
    std::ops::Range<usize>,
    String,
    Option<rdocx_oxml::properties::CT_RPr>,
    Vec<&'data Value>,
    Option<std::rc::Rc<TemplateCallbackEvent>>,
);

#[derive(Clone)]
struct ProgramExecution<'data> {
    blocks: Vec<ExecutedBlock<'data>>,
}

#[derive(Clone)]
enum ExecutedBlock<'data> {
    Item {
        index: usize,
        scalars: Vec<PreparedTemplateScalar<'data>>,
        children: Vec<ProgramExecution<'data>>,
    },
    Repeat {
        source_range: std::ops::Range<usize>,
        copies: Vec<Vec<ExecutedBlock<'data>>>,
    },
}

impl<'data> ProgramExecution<'data> {
    fn checked_clone(
        &self,
        operation: &mut rdocx_oxml::document::ScopedWordOperation<'data>,
    ) -> Result<Self> {
        fn reserve<'data>(
            blocks: &[ExecutedBlock<'data>],
            operation: &mut rdocx_oxml::document::ScopedWordOperation<'data>,
        ) -> Result<()> {
            for block in blocks {
                match block {
                    ExecutedBlock::Item {
                        scalars, children, ..
                    } => {
                        for (_, _, text, _, _, callback) in scalars {
                            if let Some(event) = callback {
                                operation.register_template_callback_input(text, &event.copy)?;
                            }
                            operation.reserve_template_text_copy(text.len())?;
                        }
                        for child in children {
                            reserve(&child.blocks, operation)?;
                        }
                    }
                    ExecutedBlock::Repeat { copies, .. } => {
                        for copy in copies {
                            reserve(copy, operation)?;
                        }
                    }
                }
            }
            Ok(())
        }
        reserve(&self.blocks, operation)?;
        Ok(self.clone())
    }
}

#[derive(Default)]
struct TemplateCallbackEvent {
    copy: std::cell::Cell<bool>,
    field: std::cell::Cell<bool>,
}

// Physical owner paths are captured from the actual staging clone input.
// The initial preparation traversal retains its existing order independently
// of those paths. No candidate addresses survive the clone producer.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
enum TemplateWordStep {
    Main,
    Footnote(usize),
    Comment(usize),
    Glossary(usize),
    Paragraph(usize),
    BodyItem(usize),
    TableRow(usize),
    TableControl(usize),
    RowCell(usize),
    RowControl(usize),
    CellItem(usize),
    ControlItem(usize),
    ParagraphControl(usize),
    ParagraphRevision(usize),
    ControlRevision(usize),
    RevisionParagraph,
    RevisionRun(usize),
    NestedRevision(usize),
}

enum TemplateParagraphOwner<'a> {
    Body(&'a [rdocx_oxml::document::BodyContent]),
    Paragraph(&'a CT_P),
    Table(&'a CT_Tbl),
    Row(&'a CT_Row),
    Cell(&'a CT_Tc),
    Control(&'a CT_Sdt),
    Revision(&'a rdocx_oxml::revision::CT_Revision),
    Run(&'a rdocx_oxml::text::CT_R),
}

impl<'a> TemplateParagraphOwner<'a> {
    fn children(&self) -> Vec<(TemplateWordStep, Self)> {
        use TemplateWordStep as Step;
        let mut children = Vec::new();
        match self {
            Self::Body(items) => {
                for (index, item) in items.iter().enumerate() {
                    let child = match item {
                        rdocx_oxml::document::BodyContent::Paragraph(p) => Self::Paragraph(p),
                        rdocx_oxml::document::BodyContent::Table(t) => Self::Table(t),
                        rdocx_oxml::document::BodyContent::ContentControl(c) => Self::Control(c),
                        rdocx_oxml::document::BodyContent::RawXml(_) => continue,
                    };
                    children.push((Step::BodyItem(index), child));
                }
            }
            Self::Paragraph(p) => {
                for (index, (_, _, revision)) in p.revisions.iter().enumerate() {
                    children.push((Step::ParagraphRevision(index), Self::Revision(revision)));
                }
                for (index, (_, _, _, control)) in p.content_controls.iter().enumerate() {
                    children.push((Step::ParagraphControl(index), Self::Control(control)));
                }
            }
            Self::Table(t) => {
                for (index, row) in t.rows.iter().enumerate() {
                    children.push((Step::TableRow(index), Self::Row(row)));
                }
                for (index, (_, _, control)) in t.content_controls.iter().enumerate() {
                    children.push((Step::TableControl(index), Self::Control(control)));
                }
            }
            Self::Row(r) => {
                for (index, cell) in r.cells.iter().enumerate() {
                    children.push((Step::RowCell(index), Self::Cell(cell)));
                }
                for (index, (_, _, control)) in r.content_controls.iter().enumerate() {
                    children.push((Step::RowControl(index), Self::Control(control)));
                }
            }
            Self::Cell(c) => {
                for (index, item) in c.content.iter().enumerate() {
                    let child = match item {
                        CellContent::Paragraph(p) => Self::Paragraph(p),
                        CellContent::Table(t) => Self::Table(t),
                        CellContent::ContentControl(c) => Self::Control(c),
                    };
                    children.push((Step::CellItem(index), child));
                }
            }
            Self::Control(c) => {
                for (index, (_, revision)) in c.revisions().iter().enumerate() {
                    children.push((Step::ControlRevision(index), Self::Revision(revision)));
                }
                for (index, item) in c.content.iter().enumerate() {
                    let child = match item {
                        SdtContent::Paragraph(p) => Self::Paragraph(p),
                        SdtContent::Table(t) => Self::Table(t),
                        SdtContent::Row(r) => Self::Row(r),
                        SdtContent::Cell(c) => Self::Cell(c),
                        SdtContent::ContentControl(c) => Self::Control(c),
                        SdtContent::Run(_) | SdtContent::RawXml(_) => continue,
                    };
                    children.push((Step::ControlItem(index), child));
                }
            }
            Self::Revision(revision) => {
                if let Some(paragraph) = revision.content_paragraph() {
                    children.push((Step::RevisionParagraph, Self::Paragraph(paragraph)));
                }
                if let rdocx_oxml::revision::RevisionContent::Runs(runs) = revision.content() {
                    for (index, run) in runs.iter().enumerate() {
                        children.push((Step::RevisionRun(index), Self::Run(run)));
                    }
                }
                for (index, (_, revision)) in revision.nested_revisions().iter().enumerate() {
                    children.push((Step::NestedRevision(index), Self::Revision(revision)));
                }
            }
            Self::Run(_) => {}
        }
        children
    }
}

// Borrowed producer pairs live only through one substantive writer call.
#[derive(Default)]
pub(crate) struct TemplateWordInputs<'current, 'frozen> {
    pub(crate) texts: Vec<(
        &'current rdocx_oxml::text::CT_Text,
        &'frozen rdocx_oxml::text::CT_Text,
    )>,
    pub(crate) fields: Vec<(
        &'current rdocx_oxml::text::Field,
        &'frozen rdocx_oxml::text::Field,
    )>,
    pub(crate) raw: Vec<(&'current [u8], &'frozen [u8])>,
}

impl<'current, 'frozen> TemplateWordInputs<'current, 'frozen> {
    fn run(
        &mut self,
        current: &'current rdocx_oxml::text::CT_R,
        frozen: &'frozen rdocx_oxml::text::CT_R,
    ) {
        for (current, frozen) in current.content.iter().zip(&frozen.content) {
            match (current, frozen) {
                (RunContent::Text(current), RunContent::Text(frozen))
                | (RunContent::DeletedText(current), RunContent::DeletedText(frozen)) => {
                    self.texts.push((current, frozen))
                }
                (RunContent::Field(current), RunContent::Field(frozen)) => {
                    self.fields.push((current, frozen))
                }
                (RunContent::Drawing(current), RunContent::Drawing(frozen)) => {
                    if let (Some(current), Some(frozen)) = (&current.inline, &frozen.inline)
                        && let (Some(current), Some(frozen)) = (&current.raw_xml, &frozen.raw_xml)
                    {
                        self.raw.push((current, frozen));
                    }
                    if let (Some(current), Some(frozen)) = (&current.anchor, &frozen.anchor)
                        && let (Some(current), Some(frozen)) = (&current.raw_xml, &frozen.raw_xml)
                    {
                        self.raw.push((current, frozen));
                    }
                }
                _ => {}
            }
        }
        for (current, frozen) in current.extra_xml.iter().zip(&frozen.extra_xml) {
            self.raw.push((current, frozen));
        }
    }
}

// Text, Field and retained byte slices are three current instantiations.
fn dedup_template_pairs<T: ?Sized>(pairs: &mut Vec<(&T, &T)>) -> bool {
    let mut seen = HashMap::new();
    let mut valid = true;
    pairs.retain(|(current, frozen)| {
        let current = (
            std::ptr::from_ref(*current).cast::<()>() as usize,
            std::mem::size_of_val(*current),
        );
        let frozen = (
            std::ptr::from_ref(*frozen).cast::<()>() as usize,
            std::mem::size_of_val(*frozen),
        );
        if let Some(previous) = seen.insert(current, frozen) {
            valid &= previous == frozen;
            false
        } else {
            true
        }
    });
    valid
}

pub(crate) struct TemplateWordLineage<'a> {
    pub(crate) records: Option<&'a crate::document::TemplateWordRecordInputs>,
    main_body: &'a [rdocx_oxml::document::BodyContent],
    frozen_notes: &'a [rdocx_oxml::footnotes::CT_Footnote],
    frozen_comments: Option<&'a rdocx_oxml::comments::CT_Comments>,
    original_main_part: Option<(&'a str, usize)>,
    original_parts: Vec<(&'a str, &'a std::sync::Arc<Vec<u8>>)>,
    preparation: Vec<Option<(Vec<TemplateWordStep>, &'a CT_P)>>,
    replaced_fields: HashSet<(
        Vec<TemplateWordStep>,
        rdocx_oxml::text::AcceptedRunPath,
        usize,
    )>,
    prepared: bool,
}

impl<'a> TemplateWordLineage<'a> {
    pub(crate) fn clone_for_staging(document: &'a Document) -> (Document, Self) {
        use TemplateParagraphOwner as Owner;
        use TemplateWordStep as Step;
        let mut pending = vec![(
            vec![Step::Main],
            Owner::Body(&document.document.body.content),
        )];
        for (note, model) in document.footnotes.footnotes.iter().enumerate() {
            for (index, paragraph) in model.paragraphs.iter().enumerate() {
                pending.push((
                    vec![Step::Footnote(note), Step::Paragraph(index)],
                    Owner::Paragraph(paragraph),
                ));
            }
        }
        if let Some(comments) = &document.comments {
            for (comment, model) in comments.comments.iter().enumerate() {
                for (index, paragraph) in model.paragraphs.iter().enumerate() {
                    pending.push((
                        vec![Step::Comment(comment), Step::Paragraph(index)],
                        Owner::Paragraph(paragraph),
                    ));
                }
            }
        }
        if let Some(glossary) = &document.glossary {
            for (index, part) in glossary.doc_parts.iter().enumerate() {
                pending.push((vec![Step::Glossary(index)], Owner::Body(&part.body.content)));
            }
        }
        let mut physical = HashMap::new();
        while let Some((path, owner)) = pending.pop() {
            if let Owner::Paragraph(paragraph) = owner {
                physical.insert(std::ptr::from_ref(paragraph), (path.clone(), paragraph));
            }
            for (step, child) in owner.children() {
                let mut child_path = path.clone();
                child_path.push(step);
                pending.push((child_path, child));
            }
        }
        let mut preparation = Vec::new();
        let mut record = |paragraph: &'a CT_P| {
            preparation.push(physical.get(&std::ptr::from_ref(paragraph)).cloned());
        };
        crate::document::visit_body_paragraphs(&document.document.body.content, false, &mut record);
        for note in &document.footnotes.footnotes {
            for paragraph in &note.paragraphs {
                record(paragraph);
            }
        }
        if let Some(comments) = &document.comments {
            for comment in &comments.comments {
                for paragraph in &comment.paragraphs {
                    record(paragraph);
                }
            }
        }
        if let Some(glossary) = &document.glossary {
            for part in &glossary.doc_parts {
                crate::document::visit_body_paragraphs(&part.body.content, false, &mut record);
            }
        }
        // This exact clone establishes the recorded path correspondence.
        let candidate = document.clone_for_staging();
        (
            candidate,
            Self {
                records: None,
                main_body: &document.document.body.content,
                frozen_notes: &document.footnotes.footnotes,
                frozen_comments: document.comments.as_ref(),
                original_main_part: document
                    .package
                    .get_part(&document.doc_part_name)
                    .map(|bytes| (document.doc_part_name.as_str(), bytes.len())),
                original_parts: document
                    .package
                    .parts
                    .iter()
                    .map(|(name, bytes)| (name.as_str(), bytes))
                    .collect(),
                preparation,
                replaced_fields: HashSet::new(),
                prepared: false,
            },
        )
    }

    pub(crate) fn original_part_source(
        &self,
        operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
        name: &str,
        bytes: &std::sync::Arc<Vec<u8>>,
    ) -> Result<Option<rdocx_oxml::document::ScopedWordSource>> {
        let Some((name, frozen)) = self
            .original_parts
            .iter()
            .find(|(part, original)| *part == name && std::sync::Arc::ptr_eq(original, bytes))
        else {
            return Ok(None);
        };
        Ok(Some(operation.source_part(
            name,
            0..frozen.len(),
            &Default::default(),
        )?))
    }

    pub(crate) fn initial_main_part(&self) -> Option<(&'a str, usize)> {
        (!self.prepared && self.replaced_fields.is_empty())
            .then_some(self.original_main_part)
            .flatten()
    }

    // Physical clone paths survive slot-only preparation. The actual caller
    // disables this mapping when a canonicalizer replaces the whole model.
    // Returned prepared Field roots carry their own operation-minted receipts.
    pub(crate) fn model_writer_inputs<'current>(
        &self,
        index: usize,
        item: &'current rdocx_oxml::document::BodyContent,
    ) -> Option<TemplateWordInputs<'current, 'a>> {
        use rdocx_oxml::document::BodyContent;
        let frozen = self.main_body.get(index)?;
        let mut inputs = TemplateWordInputs::default();
        let owners = match (item, frozen) {
            (BodyContent::Paragraph(current), BodyContent::Paragraph(frozen)) => (
                TemplateParagraphOwner::Paragraph(current),
                TemplateParagraphOwner::Paragraph(frozen),
            ),
            (BodyContent::Table(current), BodyContent::Table(frozen)) => (
                TemplateParagraphOwner::Table(current),
                TemplateParagraphOwner::Table(frozen),
            ),
            (BodyContent::ContentControl(current), BodyContent::ContentControl(frozen)) => (
                TemplateParagraphOwner::Control(current),
                TemplateParagraphOwner::Control(frozen),
            ),
            (BodyContent::RawXml(current), BodyContent::RawXml(frozen)) => {
                inputs.raw.push((current.as_slice(), frozen.as_slice()));
                return Some(inputs);
            }
            _ => return None,
        };
        self.collect_writer_inputs(
            owners,
            vec![TemplateWordStep::Main, TemplateWordStep::BodyItem(index)],
        )
    }

    pub(crate) fn note_writer_inputs<'current>(
        &self,
        index: usize,
        note: &'current rdocx_oxml::footnotes::CT_Footnote,
    ) -> Option<TemplateWordInputs<'current, 'a>> {
        let frozen = self.frozen_notes.get(index)?;
        if note.paragraphs.len() != frozen.paragraphs.len() {
            return None;
        }
        let mut output = TemplateWordInputs::default();
        for (paragraph, (current, frozen)) in
            note.paragraphs.iter().zip(&frozen.paragraphs).enumerate()
        {
            let child = self.collect_writer_inputs(
                (
                    TemplateParagraphOwner::Paragraph(current),
                    TemplateParagraphOwner::Paragraph(frozen),
                ),
                vec![
                    TemplateWordStep::Footnote(index),
                    TemplateWordStep::Paragraph(paragraph),
                ],
            )?;
            output.texts.extend(child.texts);
            output.fields.extend(child.fields);
            output.raw.extend(child.raw);
        }
        Some(output)
    }

    pub(crate) fn comments_writer_inputs<'current>(
        &self,
        comments: &'current rdocx_oxml::comments::CT_Comments,
    ) -> Option<TemplateWordInputs<'current, 'a>> {
        let frozen = self.frozen_comments?;
        if comments.comments.len() != frozen.comments.len()
            || comments.extra_xml.len() != frozen.extra_xml.len()
        {
            return None;
        }
        let mut output = TemplateWordInputs::default();
        output.raw.extend(
            comments
                .extra_xml
                .iter()
                .zip(&frozen.extra_xml)
                .map(|(current, frozen)| (current.1.as_slice(), frozen.1.as_slice())),
        );
        for (comment_index, (comment, frozen)) in
            comments.comments.iter().zip(&frozen.comments).enumerate()
        {
            if comment.paragraphs.len() != frozen.paragraphs.len()
                || comment.extra_xml.len() != frozen.extra_xml.len()
            {
                return None;
            }
            output.raw.extend(
                comment
                    .extra_xml
                    .iter()
                    .zip(&frozen.extra_xml)
                    .map(|(current, frozen)| (current.1.as_slice(), frozen.1.as_slice())),
            );
            for (paragraph_index, (current, frozen)) in comment
                .paragraphs
                .iter()
                .zip(&frozen.paragraphs)
                .enumerate()
            {
                let child = self.collect_writer_inputs(
                    (
                        TemplateParagraphOwner::Paragraph(current),
                        TemplateParagraphOwner::Paragraph(frozen),
                    ),
                    vec![
                        TemplateWordStep::Comment(comment_index),
                        TemplateWordStep::Paragraph(paragraph_index),
                    ],
                )?;
                output.texts.extend(child.texts);
                output.fields.extend(child.fields);
                output.raw.extend(child.raw);
            }
        }
        Some(output)
    }

    fn collect_writer_inputs<'current>(
        &self,
        owners: (TemplateParagraphOwner<'current>, TemplateParagraphOwner<'a>),
        root_path: Vec<TemplateWordStep>,
    ) -> Option<TemplateWordInputs<'current, 'a>> {
        let mut inputs = TemplateWordInputs::default();
        let mut pending = vec![(
            owners.0,
            owners.1,
            root_path,
            None::<Vec<TemplateWordStep>>,
            Vec::<rdocx_oxml::text::AcceptedRunPathSegment>::new(),
        )];
        while let Some((current, frozen, physical, mut paragraph_root, mut accepted)) =
            pending.pop()
        {
            if matches!(current, TemplateParagraphOwner::Paragraph(_))
                && !matches!(physical.last(), Some(TemplateWordStep::RevisionParagraph))
            {
                paragraph_root = Some(physical.clone());
                accepted.clear();
            }
            match (&current, &frozen) {
                (
                    TemplateParagraphOwner::Paragraph(current),
                    TemplateParagraphOwner::Paragraph(frozen),
                ) => {
                    // These direct slots remain actual clone/preparation
                    // handbacks even when another retained owner is opaque.
                    // Retained ruby children remain actual clone pairs.
                    for (current, frozen) in current.rubies.iter().zip(&frozen.rubies) {
                        for (current, frozen) in current.raw_xml.iter().zip(&frozen.raw_xml) {
                            inputs.raw.push((current, frozen));
                        }
                    }
                    for (current, frozen) in current.runs.iter().zip(&frozen.runs) {
                        inputs.run(current, frozen);
                    }
                    for ((_, current), (_, frozen)) in
                        current.extra_xml.iter().zip(&frozen.extra_xml)
                    {
                        inputs.raw.push((current, frozen));
                    }
                }
                (
                    TemplateParagraphOwner::Control(current),
                    TemplateParagraphOwner::Control(frozen),
                ) => {
                    for ((current_index, current), (frozen_index, frozen)) in current
                        .retained_resource_run_xml()
                        .zip(frozen.retained_resource_run_xml())
                    {
                        if current_index != frozen_index {
                            return None;
                        }
                        inputs.raw.push((current, frozen));
                    }
                    for (current, frozen) in current
                        .original_resource_runs()
                        .zip(frozen.original_resource_runs())
                    {
                        inputs.run(current, frozen);
                    }
                    for (current, frozen) in current.content.iter().zip(&frozen.content) {
                        match (current, frozen) {
                            (SdtContent::Run(current), SdtContent::Run(frozen)) => {
                                inputs.run(current, frozen);
                            }
                            (SdtContent::RawXml(current), SdtContent::RawXml(frozen)) => {
                                inputs.raw.push((current, frozen))
                            }
                            _ => {}
                        }
                    }
                }
                (
                    TemplateParagraphOwner::Revision(current),
                    TemplateParagraphOwner::Revision(frozen),
                ) => {
                    inputs
                        .raw
                        .push((current.retained_xml(), frozen.retained_xml()));
                }
                (TemplateParagraphOwner::Run(current), TemplateParagraphOwner::Run(frozen)) => {
                    inputs.run(current, frozen)
                }
                (TemplateParagraphOwner::Table(current), TemplateParagraphOwner::Table(frozen)) => {
                    for ((_, current), (_, frozen)) in
                        current.extra_xml.iter().zip(&frozen.extra_xml)
                    {
                        inputs.raw.push((current, frozen));
                    }
                }
                (TemplateParagraphOwner::Row(current), TemplateParagraphOwner::Row(frozen)) => {
                    for ((_, current), (_, frozen)) in
                        current.extra_xml.iter().zip(&frozen.extra_xml)
                    {
                        inputs.raw.push((current, frozen));
                    }
                    if let (Some(current), Some(frozen)) = (
                        current.table_property_exception.as_deref(),
                        frozen.table_property_exception.as_deref(),
                    ) {
                        inputs.raw.push((current, frozen));
                    }
                }
                (TemplateParagraphOwner::Cell(current), TemplateParagraphOwner::Cell(frozen)) => {
                    for ((_, current), (_, frozen)) in
                        current.extra_xml.iter().zip(&frozen.extra_xml)
                    {
                        inputs.raw.push((current, frozen));
                    }
                }
                _ => {}
            }
            // These edges were established by the clone itself, before any
            // candidate mutation. Step matching checks the structured path,
            // never the content value, pointer or serialized bytes.
            let left = current.children();
            let right = frozen.children();
            if left.len() != right.len() {
                return None;
            }
            for ((step, current), (frozen_step, frozen)) in left.into_iter().zip(right) {
                if step != frozen_step {
                    return None;
                }
                let mut path = physical.clone();
                path.push(step.clone());
                let mut prefix = accepted.clone();
                use rdocx_oxml::text::AcceptedRunPathSegment as A;
                match step {
                    TemplateWordStep::ParagraphRevision(index)
                    | TemplateWordStep::ControlRevision(index) => prefix.push(A::Revision(index)),
                    TemplateWordStep::ParagraphControl(index)
                    | TemplateWordStep::ControlItem(index) => prefix.push(A::ContentControl(index)),
                    _ => {}
                }
                if matches!(current, TemplateParagraphOwner::Revision(_))
                    && self.replaced_fields.iter().any(|(owner, changed, _)| {
                        paragraph_root.as_ref() == Some(owner)
                            && changed.segments().starts_with(&prefix)
                    })
                {
                    // This actual changed subtree has its own scoped reader
                    // companion. Unchanged sibling revisions keep clone pairs.
                    continue;
                }
                pending.push((current, frozen, path, paragraph_root.clone(), prefix));
            }
        }
        // Private Arc-backed views can reach the same real payload twice.
        // Deduplicate actual borrowed pairs, never discover origins by address.
        if !dedup_template_pairs(&mut inputs.texts)
            || !dedup_template_pairs(&mut inputs.fields)
            || !dedup_template_pairs(&mut inputs.raw)
        {
            return None;
        }
        Some(inputs)
    }

    pub(crate) fn initial_field(
        &self,
        paragraph: usize,
        path: &rdocx_oxml::text::AcceptedRunPath,
        slot: usize,
    ) -> Option<&'a rdocx_oxml::text::Field> {
        if self.prepared {
            return None;
        }
        let (owner, frozen) = self.preparation.get(paragraph)?.as_ref()?;
        let frozen = *frozen;
        if self
            .replaced_fields
            .contains(&(owner.clone(), path.clone(), slot))
        {
            return None;
        }
        match frozen.accepted_run(path)?.content.get(slot)? {
            RunContent::Field(field) => Some(field),
            _ => None,
        }
    }

    pub(crate) fn initial_paragraph(&self, paragraph: usize) -> Option<&'a CT_P> {
        if self.prepared {
            return None;
        }
        self.preparation
            .get(paragraph)?
            .as_ref()
            .map(|(_, frozen)| *frozen)
    }

    pub(crate) fn installed_field(
        &mut self,
        paragraph: usize,
        path: &rdocx_oxml::text::AcceptedRunPath,
        slot: usize,
    ) {
        if let Some(Some((owner, _))) = self.preparation.get(paragraph) {
            self.replaced_fields
                .insert((owner.clone(), path.clone(), slot));
        }
    }

    pub(crate) fn finish_preparation(&mut self) {
        self.prepared = true;
    }
}

pub(crate) struct TemplateEvaluation<'data, 'formatter> {
    data: &'data Value,
    pub(crate) operation: rdocx_oxml::document::ScopedWordOperation<'data>,
    pub(crate) lineage: TemplateWordLineage<'data>,
    pub(crate) sources: crate::document::TemplateMainSources,
    versioned: bool,
    occurrence: usize,
    // The callback lifetime stays on this field. A type alias would force 'static.
    #[allow(clippy::type_complexity)]
    formatter: Option<
        &'formatter mut dyn FnMut(
            &TemplateFormatContext<'_>,
            &mut crate::MailMergeFormattedText,
        ) -> Result<()>,
    >,
}

fn program_marker(sources: &[String], versioned: bool) -> Result<Option<Control>> {
    if !versioned {
        return marker_from_sources(sources);
    }
    let controls = sources
        .iter()
        .filter(|source| {
            source.trim_start().starts_with("{%") && !source.trim_start().starts_with("{%{%")
        })
        .collect::<Vec<_>>();
    if controls.is_empty() {
        return Ok(None);
    }
    if controls.len() != 1
        || sources
            .iter()
            .filter(|source| !source.trim().is_empty())
            .count()
            != 1
    {
        return Err(template_error(
            "a control marker must occupy one sibling paragraph or row",
        ));
    }
    let text = controls[0].trim();
    let inner = text
        .strip_prefix("{%")
        .and_then(|text| text.strip_suffix("%}"))
        .ok_or_else(|| template_error("malformed control marker"))?
        .trim();
    if inner == "endfor" {
        return Ok(Some(Control::End(EndKind::For)));
    }
    if inner == "endif" {
        return Ok(Some(Control::End(EndKind::If)));
    }
    let (name, path, is_loop) = if let Some(path) = inner.strip_prefix("if ") {
        ("", path, false)
    } else if let Some(rest) = inner.strip_prefix("for ") {
        let (name, path) = rest
            .split_once(" in ")
            .ok_or_else(|| template_error("for requires a variable and path"))?;
        validate_scope_name(name.trim())?;
        (name.trim(), path, true)
    } else {
        return Err(template_error("unsupported control marker"));
    };
    let expression = parse_expression(path)?;
    if !expression.filters.is_empty() {
        return Err(template_error(
            "structural controls require a path without scalar filters",
        ));
    }
    Ok(Some(if is_loop {
        Control::For {
            name: name.into(),
            path: vec![expression.spelling],
        }
    } else {
        Control::If {
            path: vec![expression.spelling],
        }
    }))
}

fn parse_program<'program>(
    source: &'program TemplateProgram,
    versioned: bool,
    inherited_field: bool,
    scoped: &mut Option<(
        &mut rdocx_oxml::document::ScopedWordOperation<'_>,
        &rdocx_oxml::document::ScopedWordSource,
    )>,
) -> Result<ParsedProgram<'program>> {
    let mut markers = Vec::new();
    let mut scalars = Vec::new();
    let mut children = Vec::new();
    for item in &source.items {
        let mut tokens = Vec::new();
        let marker = if item.kind == TemplateItemKind::Source(crate::StoryItemKind::Paragraph) {
            let texts = if let Some((operation, parent)) = scoped {
                let selected =
                    operation.compose_sources(&[(Some(*parent), item.range.clone(), &[])])?;
                let closed = operation.close_source_namespaces(selected, &item.scope)?;
                operation.template_replaceable_texts(&closed)?
            } else {
                let closed =
                    crate::document::close_content_fragment_namespaces(&item.raw, &item.scope)?;
                let paragraph = CT_P::from_xml_fragment(&closed)?;
                placeholder::replaceable_texts(&paragraph)
            };
            let marker = program_marker(&texts, versioned)?;
            if marker.is_some() && (item.field_owned || inherited_field) {
                return Err(template_error(
                    "structural marker overlaps a physical field owner",
                ));
            }
            if marker.is_none() {
                for (stretch, text) in texts.iter().enumerate() {
                    tokens.extend(
                        scalar_tokens(text, versioned)?
                            .into_iter()
                            .map(|token| (stretch, token)),
                    );
                }
            }
            marker
        } else if item.kind == TemplateItemKind::TableRow {
            let (direct, all, preserved_siblings) = if let Some((operation, parent)) = scoped {
                let selected =
                    operation.compose_sources(&[(Some(*parent), item.range.clone(), &[])])?;
                let closed = operation.close_source_namespaces(selected, &item.scope)?;
                let (buffers, direct, all, hazard) =
                    operation.template_row_replaceable_texts(&closed)?;
                (
                    direct
                        .into_iter()
                        .map(|index| {
                            operation.reserve_template_text_copy(buffers[index].len())?;
                            Ok(buffers[index].clone())
                        })
                        .collect::<Result<Vec<_>>>()?,
                    all.into_iter()
                        .map(|index| {
                            operation.reserve_template_text_copy(buffers[index].len())?;
                            Ok(buffers[index].clone())
                        })
                        .collect::<Result<Vec<_>>>()?,
                    hazard,
                )
            } else {
                let mut wrapped =
                    format!(r#"<w:tbl xmlns:w="{}">"#, rdocx_oxml::namespace::W_NS).into_bytes();
                wrapped.extend(crate::document::close_content_fragment_namespaces(
                    &item.raw,
                    &item.scope,
                )?);
                wrapped.extend_from_slice(b"</w:tbl>");
                let table = CT_Tbl::from_xml_fragment_checked(&wrapped)?;
                let row = table
                    .rows
                    .first()
                    .ok_or_else(|| template_error("prepared row source disappeared"))?;
                let mut direct = Vec::new();
                collect_row_marker_sources(row, &mut direct);
                let mut all = Vec::new();
                collect_row(row, &mut all);
                let hazard = row
                    .extra_xml
                    .iter()
                    .any(|(position, raw)| !CT_Row::raw_is_root_attributes(*position, raw))
                    || !row.content_controls.is_empty();
                (direct, all, hazard)
            };
            let marker = if direct.iter().filter(|text| !text.trim().is_empty()).count() == 1
                && program_marker(&direct, versioned)?.is_some()
            {
                program_marker(&all, versioned)?
            } else {
                None
            };
            if marker.is_some()
                && (item.marker_unsafe || item.field_owned || inherited_field || preserved_siblings)
            {
                return Err(template_error(
                    "a marker row cannot carry preserved siblings or ambiguous field ownership",
                ));
            }
            marker
        } else {
            None
        };
        markers.push(marker.clone());
        scalars.push(tokens);
        children.push(
            item.children
                .iter()
                .map(|child| {
                    if marker.is_some() && item.kind == TemplateItemKind::TableRow {
                        Ok(ParsedProgram {
                            source: child,
                            blocks: Vec::new(),
                            scalars: Vec::new(),
                            children: Vec::new(),
                        })
                    } else {
                        parse_program(
                            child,
                            versioned,
                            inherited_field || item.field_owned,
                            scoped,
                        )
                    }
                })
                .collect::<Result<Vec<_>>>()?,
        );
    }
    let indices = (0..source.items.len()).collect::<Vec<_>>();
    let blocks = parse_blocks(&indices, |index| Ok(markers[*index].clone()))?;
    Ok(ParsedProgram {
        source,
        blocks,
        scalars,
        children,
    })
}

fn deferred_expression(expression: &Expression, root: &Value, scopes: &[Scope]) -> bool {
    let Some(PathStep::Key(key)) = expression.path.first() else {
        return false;
    };
    if let Some(scope) = scopes.iter().rev().find(|scope| scope.name == *key) {
        return scope.deferred;
    }
    // An empty item cannot supply implicit properties yet. Known outer/root
    // properties must still be evaluated, including invalid intermediate kinds.
    scopes.iter().any(|scope| scope.deferred)
        && !root
            .as_object()
            .is_some_and(|value| value.contains_key(key))
        && !scopes.iter().any(|scope| {
            !scope.deferred
                && scope
                    .value
                    .and_then(Value::as_object)
                    .is_some_and(|value| value.contains_key(key))
        })
}

fn program_path_value<'a>(
    path: &[String],
    evaluation: &mut TemplateEvaluation<'a, '_>,
    scopes: &[Scope<'a>],
) -> Result<ResolvedValue<'a>> {
    if !evaluation.versioned {
        return resolve_value(evaluation.data, scopes, path, &mut evaluation.operation);
    }
    let expression = parse_expression(&path[0])?;
    if deferred_expression(&expression, evaluation.data, scopes) {
        return Ok(ResolvedValue::Deferred);
    }
    expression_value(
        &expression,
        evaluation.data,
        scopes,
        &mut evaluation.operation,
    )?
    .map(ResolvedValue::Value)
    .ok_or_else(|| template_error("missing structural template path"))
}

fn evaluate_program_blocks<'data>(
    program: &ParsedProgram<'_>,
    blocks: &[Block<usize>],
    scopes: &mut Vec<Scope<'data>>,
    evaluation: &mut TemplateEvaluation<'data, '_>,
    emit: bool,
) -> Result<Vec<ExecutedBlock<'data>>> {
    let mut output = Vec::new();
    for block in blocks {
        match block {
            Block::Item { value: index, .. } => {
                let mut scalars = Vec::new();
                for (stretch, token) in &program.scalars[*index] {
                    let mut dependencies = Vec::new();
                    let (text, properties) = match &token.value {
                        ScalarValue::Literal(text) => {
                            if !emit {
                                continue;
                            }
                            evaluation
                                .operation
                                .reserve_template_text_copy(text.len())?;
                            ((*text).to_owned(), None)
                        }
                        ScalarValue::Expression(expression) => {
                            if deferred_expression(expression, evaluation.data, scopes) {
                                continue;
                            }
                            if !emit && evaluation.versioned {
                                evaluate_expression(
                                    expression,
                                    evaluation.data,
                                    scopes,
                                    &mut evaluation.operation,
                                    None,
                                    false,
                                )?;
                                continue;
                            }
                            let text = if evaluation.versioned {
                                expression_text_in(
                                    expression,
                                    evaluation.data,
                                    scopes,
                                    &mut evaluation.operation,
                                    Some(&mut dependencies),
                                )?
                            } else {
                                let path = expression
                                    .path
                                    .iter()
                                    .map(|step| match step {
                                        PathStep::Key(key) => key.clone(),
                                        PathStep::Index(_) => {
                                            unreachable!("legacy path has only keys")
                                        }
                                    })
                                    .collect::<Vec<_>>();
                                match resolve_value(
                                    evaluation.data,
                                    scopes,
                                    &path,
                                    &mut evaluation.operation,
                                )? {
                                    ResolvedValue::Value(value) => {
                                        if !emit {
                                            require_scalar(value, &path)?;
                                            validate_expression_text(
                                                ValidationText::Scalar(value),
                                                &[],
                                            )?;
                                            continue;
                                        }
                                        dependencies.push(value);
                                        scalar_text(value, &path, &mut evaluation.operation)?
                                    }
                                    ResolvedValue::Deferred => continue,
                                }
                            };
                            oxml_core::xml::reject_non_xml_characters(
                                "template replacement",
                                &text,
                            )?;
                            (text, None)
                        }
                    };
                    if emit {
                        scalars.push((
                            *stretch,
                            token.range.clone(),
                            text,
                            properties,
                            dependencies,
                            None,
                        ));
                    }
                }
                let children = program.children[*index]
                    .iter()
                    .map(|child| {
                        evaluate_program_blocks(child, &child.blocks, scopes, evaluation, emit)
                            .map(|blocks| ProgramExecution { blocks })
                    })
                    .collect::<Result<Vec<_>>>()?;
                if emit {
                    output.push(ExecutedBlock::Item {
                        index: *index,
                        scalars,
                        children,
                    });
                }
            }
            Block::For {
                name,
                path,
                source_range,
                body,
            } => {
                let array = match program_path_value(path, evaluation, scopes)? {
                    ResolvedValue::Value(value @ Value::Array(_)) => Some(value),
                    ResolvedValue::Deferred => None,
                    _ => return Err(template_error("loop path must resolve to an array")),
                };
                let values = array
                    .and_then(Value::as_array)
                    .map_or(&[][..], Vec::as_slice);
                if values.is_empty() {
                    scopes.push(Scope {
                        name: name.clone(),
                        value: None,
                        deferred: true,
                    });
                    evaluate_program_blocks(program, body, scopes, evaluation, false)?;
                    scopes.pop();
                }
                let mut copies = Vec::new();
                for index in 0..values.len() {
                    let value = evaluation
                        .operation
                        .template_datum_index(
                            array.expect("nonempty loop has an actual input array"),
                            index,
                        )?
                        .expect("index belongs to the immutable selected array");
                    scopes.push(Scope {
                        name: name.clone(),
                        value: Some(value),
                        deferred: false,
                    });
                    let blocks = evaluate_program_blocks(program, body, scopes, evaluation, emit)?;
                    scopes.pop();
                    if emit {
                        copies.push(blocks);
                    }
                }
                if emit {
                    output.push(ExecutedBlock::Repeat {
                        source_range: source_range.clone(),
                        copies,
                    });
                }
            }
            Block::If { path, body } => {
                let enabled = matches!(program_path_value(path, evaluation, scopes)?, ResolvedValue::Value(value) if is_truthy(value));
                let blocks =
                    evaluate_program_blocks(program, body, scopes, evaluation, emit && enabled)?;
                if emit && enabled {
                    output.extend(blocks);
                }
            }
        }
    }
    Ok(output)
}

fn validate_paragraph_numbering(
    paragraph: &CT_P,
    document: &Document,
    repeated: bool,
) -> Result<()> {
    if !repeated {
        return Ok(());
    }
    let Some(properties) = &paragraph.properties else {
        return Ok(());
    };
    let (num_id, level) = match (properties.num_id, properties.num_ilvl) {
        (Some(num_id), level) => (num_id, level.unwrap_or(0)),
        // A `w:ilvl` without a `w:numId` takes the instance of the paragraph
        // style chain, as Word does, and numbers nothing without one.
        (None, Some(level)) => match crate::style::resolve_paragraph_properties(
            properties.style_id.as_deref(),
            &document.styles,
        )
        .num_id
        {
            Some(num_id) if num_id != 0 => (num_id, level),
            _ => return Ok(()),
        },
        (None, None) => return Ok(()),
    };
    if !document.template_numbering_reference_exists(num_id, level) {
        return Err(template_error(&format!(
            "a repeated paragraph references missing numbering numId {num_id} level {level}"
        )));
    }
    Ok(())
}

fn validate_program_numbering(
    program: &ParsedProgram<'_>,
    blocks: &[Block<usize>],
    document: &Document,
    repeated: bool,
    scoped: &mut Option<(
        &mut rdocx_oxml::document::ScopedWordOperation<'_>,
        &rdocx_oxml::document::ScopedWordSource,
    )>,
) -> Result<()> {
    for block in blocks {
        match block {
            Block::Item { value: index, .. } => {
                let item = &program.source.items[*index];
                if repeated
                    && item.kind == TemplateItemKind::Source(crate::StoryItemKind::Paragraph)
                {
                    let paragraph = if let Some((operation, parent)) = scoped {
                        let selected = operation.compose_sources(&[(
                            Some(*parent),
                            item.range.clone(),
                            &[],
                        )])?;
                        let closed = operation.close_source_namespaces(selected, &item.scope)?;
                        operation.source_paragraph(&closed)?
                    } else {
                        let closed = crate::document::close_content_fragment_namespaces(
                            &item.raw,
                            &item.scope,
                        )?;
                        CT_P::from_xml_fragment(&closed)?
                    };
                    validate_paragraph_numbering(&paragraph, document, true)?;
                }
                for child in &program.children[*index] {
                    validate_program_numbering(child, &child.blocks, document, repeated, scoped)?;
                }
            }
            Block::For { body, .. } => {
                validate_program_numbering(program, body, document, true, scoped)?
            }
            Block::If { body, .. } => {
                validate_program_numbering(program, body, document, repeated, scoped)?
            }
        }
    }
    Ok(())
}

fn flattened_execution_items<'borrow, 'data>(
    blocks: &'borrow [ExecutedBlock<'data>],
) -> Vec<&'borrow ExecutedBlock<'data>> {
    let mut items = Vec::new();
    for block in blocks {
        match block {
            ExecutedBlock::Item { .. } => items.push(block),
            ExecutedBlock::Repeat { copies, .. } => {
                for copy in copies {
                    items.extend(flattened_execution_items(copy));
                }
            }
        }
    }
    items
}

fn format_program_execution<'data>(
    program: &ParsedProgram<'_>,
    blocks: &mut [ExecutedBlock<'data>],
    current: &TemplateProgram,
    position: &mut usize,
    evaluation: &mut TemplateEvaluation<'data, '_>,
) -> Result<()> {
    for block in blocks {
        match block {
            ExecutedBlock::Repeat { copies, .. } => {
                for copy in copies {
                    format_program_execution(program, copy, current, position, evaluation)?;
                }
            }
            ExecutedBlock::Item {
                index,
                scalars,
                children,
            } => {
                let actual = current
                    .items
                    .get(*position)
                    .ok_or_else(|| template_error("expanded template source is missing"))?;
                *position += 1;
                if actual.kind != program.source.items[*index].kind
                    || scalars.len() != program.scalars[*index].len()
                    || children.len() != actual.children.len()
                {
                    return Err(template_error(
                        "expanded template source correspondence changed",
                    ));
                }
                for (prepared, (_, token)) in scalars.iter_mut().zip(&program.scalars[*index]) {
                    let ScalarValue::Expression(expression) = &token.value else {
                        continue;
                    };
                    if let Some(event) = &prepared.5 {
                        evaluation
                            .operation
                            .register_template_callback_input(&prepared.2, &event.copy)?;
                    }
                    evaluation
                        .operation
                        .reserve_template_text_copy(prepared.2.len())?;
                    let mut formatted = crate::MailMergeFormattedText {
                        text: prepared.2.clone(),
                        run_properties: prepared.3.clone(),
                    };
                    if let Some(callback) = evaluation.formatter.as_deref_mut() {
                        callback(
                            &TemplateFormatContext {
                                location: &program.source.items[*index].location,
                                path: &expression.spelling,
                                occurrence: evaluation.occurrence,
                            },
                            &mut formatted,
                        )?;
                    }
                    evaluation.occurrence += 1;
                    oxml_core::xml::reject_non_xml_characters(
                        "template replacement",
                        &formatted.text,
                    )?;
                    if formatted.run_properties.is_some() {
                        let mut paragraph = CT_P::new();
                        let mut run = rdocx_oxml::text::CT_R::new(&formatted.text);
                        run.properties = formatted.run_properties.clone();
                        paragraph.runs.push(run);
                        let mut writer = quick_xml::Writer::new(Vec::new());
                        paragraph.to_xml(&mut writer)?;
                        let closed = crate::document::close_content_fragment_namespaces(
                            &writer.into_inner(),
                            &program.source.items[*index].scope,
                        )?;
                        CT_P::from_xml_fragment_checked(&closed)?;
                    }
                    if evaluation.formatter.is_some() {
                        prepared.4.clear();
                        prepared.5 = Some(std::rc::Rc::new(TemplateCallbackEvent::default()));
                    }
                    prepared.2 = formatted.text;
                    prepared.3 = formatted.run_properties;
                }
                for ((source, execution), child) in program.children[*index]
                    .iter()
                    .zip(children)
                    .zip(&actual.children)
                {
                    let mut child_position = 0;
                    format_program_execution(
                        source,
                        &mut execution.blocks,
                        child,
                        &mut child_position,
                        evaluation,
                    )?;
                    if child_position != child.items.len()
                        && !(child_position == 0 && child.required_paragraph)
                    {
                        return Err(template_error(
                            "expanded child template source has extra items",
                        ));
                    }
                }
            }
        }
    }
    Ok(())
}

struct TemplateReplay {
    raw: Vec<u8>,
    source: Option<rdocx_oxml::document::ScopedWordSource>,
}

impl TemplateReplay {
    fn from_source(
        source: rdocx_oxml::document::ScopedWordSource,
        operation: &mut rdocx_oxml::document::ScopedWordOperation<'_>,
    ) -> Result<Self> {
        let retained = operation.repeat_source(&source)?;
        let raw = operation.finish_source(retained)?;
        Ok(Self {
            raw,
            source: Some(source),
        })
    }
}

fn replay_program_blocks<'data>(
    program: &ParsedProgram<'_>,
    blocks: &[ExecutedBlock<'data>],
    current: &TemplateProgram,
    offset: usize,
    candidate: &mut Document,
    mut note_copies: Option<&mut Vec<NoteCopy>>,
    pending_drawings: &mut HashSet<u32>,
    evaluation: &mut TemplateEvaluation<'data, '_>,
) -> Result<TemplateReplay> {
    let mut output = Vec::new();
    let mut output_sources = Vec::<TemplateReplay>::new();
    let structural = note_copies.is_some();
    let mut position = 0;
    let blocks = if structural {
        blocks.iter().collect::<Vec<_>>()
    } else {
        flattened_execution_items(blocks)
    };
    for block in blocks {
        match block {
            ExecutedBlock::Item {
                index,
                scalars,
                children,
            } => {
                let local = if structural {
                    index.checked_sub(offset).ok_or_else(|| {
                        template_error("imported template item escaped its original range")
                    })?
                } else {
                    let index = position;
                    position += 1;
                    index
                };
                let item = current
                    .items
                    .get(local)
                    .ok_or_else(|| template_error("imported template item is missing"))?;
                if item.kind != program.source.items[*index].kind
                    || children.len() != program.children[*index].len()
                    || children.len() != item.children.len()
                {
                    return Err(template_error("import changed template source topology"));
                }
                let owner_start = current.ranges[0].start;
                let mut raw_source = current
                    .source
                    .as_ref()
                    .map(|parent| {
                        evaluation.operation.compose_sources(&[(
                            Some(parent.as_ref()),
                            item.range.start - owner_start..item.range.end - owner_start,
                            &[],
                        )])
                    })
                    .transpose()?;
                let mut raw = if let Some(source) = &raw_source {
                    let retained = evaluation.operation.repeat_source(source)?;
                    evaluation.operation.finish_source(retained)?
                } else {
                    item.raw.clone()
                };
                let mut nested = Vec::new();
                for ((source, execution), actual) in program.children[*index]
                    .iter()
                    .zip(children)
                    .zip(&item.children)
                {
                    if source.source.story.kind() != actual.story.kind()
                        || source.source.ranges.len() != actual.ranges.len()
                    {
                        return Err(template_error("import changed template child ownership"));
                    }
                    for (alias, range) in actual.ranges.iter().enumerate() {
                        let relative =
                            range.start.checked_sub(item.range.start).ok_or_else(|| {
                                template_error("template child begins outside its item")
                            })?
                                ..range.end.checked_sub(item.range.start).ok_or_else(|| {
                                    template_error("template child ends outside its item")
                                })?;
                        let alias_source;
                        let original_alias;
                        let selected = if alias == 0 {
                            original_alias = None;
                            actual
                        } else {
                            let source_bytes = raw.get(relative.clone()).ok_or_else(|| {
                                template_error("paired template child source is stale")
                            })?;
                            original_alias = Some(source_bytes.to_vec());
                            alias_source = if let Some(parent) = &raw_source {
                                let selected = evaluation.operation.compose_sources(&[(
                                    Some(parent),
                                    relative.clone(),
                                    &[],
                                )])?;
                                Document::template_fragment_program_from_source(
                                    &current.part,
                                    selected,
                                    &item.scope,
                                    actual.story.kind(),
                                    &mut evaluation.operation,
                                )?
                            } else {
                                Document::template_fragment_program(
                                    &current.part,
                                    source_bytes,
                                    &item.scope,
                                    actual.story.kind(),
                                )?
                            };
                            &alias_source
                        };
                        let mut replacement = replay_program(
                            source,
                            &execution.blocks,
                            selected,
                            0,
                            candidate,
                            note_copies.as_deref_mut(),
                            pending_drawings,
                            false,
                            evaluation,
                        )?;
                        if let Some(original_alias) = original_alias.as_deref() {
                            // Parsing the copy declares inherited namespaces on
                            // its root. The stored copy keeps the root it had.
                            if let (Some(parent), Some(rendered)) =
                                (&raw_source, replacement.source.take())
                            {
                                let original_end = fragment_root_start_tag_end(original_alias)?;
                                let rendered_end =
                                    fragment_root_start_tag_end(rendered.as_bytes())?;
                                let restored = evaluation.operation.compose_sources(&[
                                    (
                                        Some(parent),
                                        relative.start..relative.start + original_end,
                                        &[],
                                    ),
                                    (
                                        Some(&rendered),
                                        rendered_end..rendered.as_bytes().len(),
                                        &[],
                                    ),
                                ])?;
                                replacement = TemplateReplay::from_source(
                                    restored,
                                    &mut evaluation.operation,
                                )?;
                            } else {
                                let restored =
                                    restore_paired_copy_root(original_alias, &replacement.raw)?;
                                replacement = TemplateReplay {
                                    raw: restored,
                                    source: None,
                                };
                            }
                        }
                        nested.push((relative, replacement));
                    }
                }
                nested.sort_by_key(|(range, _)| range.start);
                if nested
                    .windows(2)
                    .any(|pair| pair[0].0.end > pair[1].0.start)
                {
                    return Err(template_error("template child sources overlap"));
                }
                if let Some(source) = raw_source.take() {
                    let mut pieces = Vec::new();
                    let mut at = 0;
                    for (range, replacement) in &nested {
                        pieces.push((Some(&source), at..range.start, &[][..]));
                        if let Some(child) = &replacement.source {
                            pieces.push((Some(child), 0..child.as_bytes().len(), &[][..]));
                        } else {
                            pieces.push((None, 0..0, replacement.raw.as_slice()));
                        }
                        at = range.end;
                    }
                    pieces.push((Some(&source), at..source.as_bytes().len(), &[][..]));
                    let joined = evaluation.operation.compose_sources(&pieces)?;
                    let retained = evaluation.operation.repeat_source(&joined)?;
                    raw = evaluation.operation.finish_source(retained)?;
                    raw_source = Some(joined);
                } else {
                    for (range, replacement) in nested.into_iter().rev() {
                        raw.splice(range, replacement.raw);
                    }
                }
                if !structural && !scalars.is_empty() {
                    if let Some(selected) = raw_source.take() {
                        let selected = evaluation
                            .operation
                            .close_source_namespaces(selected, &item.scope)?;
                        for (_, _, text, _, _, callback) in scalars {
                            if let Some(event) = callback {
                                evaluation
                                    .operation
                                    .register_template_callback_input(text, &event.copy)?;
                            }
                        }
                        let operands = scalars
                            .iter()
                            .map(
                                |(stretch, range, text, properties, dependencies, callback)| {
                                    (
                                        *stretch,
                                        range.clone(),
                                        text.as_str(),
                                        properties.as_ref(),
                                        callback.as_ref().map(|event| &event.field),
                                        dependencies.as_slice(),
                                    )
                                },
                            )
                            .collect::<Vec<_>>();
                        let formatted = evaluation.operation.format_paragraph(
                            selected,
                            Vec::new(),
                            &operands,
                        )?;
                        let retained = evaluation.operation.repeat_source(&formatted)?;
                        raw = evaluation.operation.finish_source(retained)?;
                        raw_source = Some(formatted);
                    } else {
                        let closed =
                            crate::document::close_content_fragment_namespaces(&raw, &item.scope)?;
                        let mut paragraph = CT_P::from_xml_fragment(&closed)?;
                        for (_, _, text, _, _, callback) in scalars {
                            if let Some(event) = callback {
                                evaluation
                                    .operation
                                    .register_template_callback_input(text, &event.copy)?;
                            }
                            evaluation
                                .operation
                                .reserve_template_text_copy(text.len())?;
                        }
                        let legacy = scalars
                            .iter()
                            .map(|(stretch, range, text, properties, _, _)| {
                                (*stretch, range.clone(), text.clone(), properties.clone())
                            })
                            .collect::<Vec<_>>();
                        placeholder::replace_prepared_occurrences(&mut paragraph, &legacy)?;
                        let mut writer = quick_xml::Writer::new(Vec::new());
                        paragraph.to_xml(&mut writer)?;
                        raw = writer.into_inner();
                    }
                }
                let gap_start = if local == 0 {
                    item.range.start
                } else {
                    current.items[local - 1].range.end
                };
                let owner_start = current.ranges[0].start;
                let gap = current
                    .raw
                    .get(gap_start - owner_start..item.range.start - owner_start)
                    .ok_or_else(|| template_error("template source gap is stale"))?;
                if let Some(parent) = &current.source {
                    let gap_source = evaluation.operation.compose_sources(&[(
                        Some(parent.as_ref()),
                        gap_start - owner_start..item.range.start - owner_start,
                        &[],
                    )])?;
                    output_sources.push(TemplateReplay {
                        raw: Vec::new(),
                        source: Some(gap_source),
                    });
                    output_sources.push(TemplateReplay {
                        raw,
                        source: raw_source,
                    });
                } else {
                    output.extend_from_slice(gap);
                    output.extend(raw);
                }
            }
            ExecutedBlock::Repeat {
                source_range,
                copies,
            } => {
                if source_range.is_empty() {
                    continue;
                }
                let first = source_range
                    .start
                    .checked_sub(offset)
                    .ok_or_else(|| template_error("repeated source escaped imported range"))?;
                let end = source_range
                    .end
                    .checked_sub(offset)
                    .ok_or_else(|| template_error("repeated source escaped imported range"))?;
                let selected = current
                    .items
                    .get(first..end)
                    .ok_or_else(|| template_error("repeated source items are missing"))?;
                let mut fragment_sources = Vec::new();
                if let Some(parent) = &current.source {
                    for (position, item) in selected.iter().enumerate() {
                        if position > 0 {
                            let start = selected[position - 1].range.end - current.ranges[0].start;
                            let end = item.range.start - current.ranges[0].start;
                            fragment_sources.push(evaluation.operation.compose_sources(&[(
                                Some(parent.as_ref()),
                                start..end,
                                &[],
                            )])?);
                        }
                        let selected = evaluation.operation.compose_sources(&[(
                            Some(parent.as_ref()),
                            item.range.start - current.ranges[0].start
                                ..item.range.end - current.ranges[0].start,
                            &[],
                        )])?;
                        fragment_sources.push(
                            evaluation
                                .operation
                                .close_source_namespaces(selected, &item.scope)?,
                        );
                    }
                }
                let fragment_source = if current.source.is_some() {
                    let pieces = fragment_sources
                        .iter()
                        .map(|source| (Some(source), 0..source.as_bytes().len(), &[][..]))
                        .collect::<Vec<_>>();
                    Some(evaluation.operation.compose_sources(&pieces)?)
                } else {
                    None
                };
                let mut fragment = Vec::new();
                if fragment_source.is_none() {
                    for (position, item) in selected.iter().enumerate() {
                        if position > 0 {
                            let start = selected[position - 1].range.end - current.ranges[0].start;
                            let end = item.range.start - current.ranges[0].start;
                            fragment.extend_from_slice(&current.raw[start..end]);
                        }
                        fragment.extend(crate::document::close_content_fragment_namespaces(
                            &item.raw,
                            &item.scope,
                        )?);
                    }
                }
                for copy in copies {
                    if copy.is_empty() {
                        continue;
                    }
                    let selection = if let Some(source) = &fragment_source {
                        crate::document::TemplateFragmentSelection::Source(Box::new(
                            evaluation.operation.repeat_source(source)?,
                        ))
                    } else {
                        crate::document::TemplateFragmentSelection::Bytes(fragment.clone())
                    };
                    let imported = candidate.template_import_repeated_blocks(
                        current,
                        selection,
                        note_copies.as_deref_mut().expect("structural import"),
                        pending_drawings,
                        evaluation,
                    )?;
                    if imported.items.len() != source_range.len()
                        || imported
                            .items
                            .iter()
                            .zip(selected)
                            .any(|(a, b)| a.kind != b.kind)
                    {
                        return Err(template_error(
                            "whole-range importer changed direct item correspondence",
                        ));
                    }
                    let rendered = replay_program(
                        program,
                        copy,
                        &imported,
                        source_range.start,
                        candidate,
                        note_copies.as_deref_mut(),
                        pending_drawings,
                        false,
                        evaluation,
                    )?;
                    if current.table_rows {
                        let first = imported
                            .items
                            .first()
                            .ok_or_else(|| template_error("imported row range is empty"))?;
                        let last = imported.items.last().unwrap();
                        let start = first.range.start - imported.ranges[0].start;
                        let suffix =
                            imported.raw.len() - (last.range.end - imported.ranges[0].start);
                        let range = start..rendered.raw.len() - suffix;
                        if current.source.is_some() {
                            let selected = rendered
                                .source
                                .as_ref()
                                .map(|source| {
                                    evaluation.operation.compose_sources(&[(
                                        Some(source),
                                        range.clone(),
                                        &[],
                                    )])
                                })
                                .transpose()?;
                            output_sources.push(TemplateReplay {
                                raw: rendered.raw[range].to_vec(),
                                source: selected,
                            });
                        } else {
                            output.extend_from_slice(&rendered.raw[range]);
                        }
                        continue;
                    }
                    let open = format!(r#"<w:document xmlns:w="{}">"#, rdocx_oxml::namespace::W_NS);
                    if let Some(source) = rendered.source {
                        let wrapped = evaluation.operation.compose_sources(&[
                            (None, 0..0, open.as_bytes()),
                            (Some(&source), 0..source.as_bytes().len(), &[]),
                            (None, 0..0, b"</w:document>"),
                        ])?;
                        let fragment = crate::document::package_authoritative_body_source(
                            &mut evaluation.operation,
                            &wrapped,
                            false,
                            &BTreeMap::new(),
                        )?;
                        output_sources.push(TemplateReplay::from_source(
                            fragment,
                            &mut evaluation.operation,
                        )?);
                    } else {
                        let mut wrapped = open.into_bytes();
                        wrapped.extend(rendered.raw);
                        wrapped.extend_from_slice(b"</w:document>");
                        let fragment = crate::document::package_authoritative_body_fragment(
                            &wrapped,
                            false,
                            &BTreeMap::new(),
                        )?;
                        if current.source.is_some() {
                            output_sources.push(TemplateReplay {
                                raw: fragment,
                                source: None,
                            });
                        } else {
                            output.extend(fragment);
                        }
                    }
                }
            }
        }
    }
    if current.source.is_some() {
        let pieces = output_sources
            .iter()
            .map(|piece| match &piece.source {
                Some(source) => (Some(source), 0..source.as_bytes().len(), &[][..]),
                None => (None, 0..0, piece.raw.as_slice()),
            })
            .collect::<Vec<_>>();
        let source = evaluation.operation.compose_sources(&pieces)?;
        TemplateReplay::from_source(source, &mut evaluation.operation)
    } else {
        Ok(TemplateReplay {
            raw: output,
            source: None,
        })
    }
}

/// Drop namespace declarations added so a paired copy could be parsed alone.
/// The copy's original root start tag stays, and the rendered interior stays.
fn restore_paired_copy_root(original: &[u8], rendered: &[u8]) -> Result<Vec<u8>> {
    let original_end = fragment_root_start_tag_end(original)?;
    let rendered_end = fragment_root_start_tag_end(rendered)?;
    let mut output = Vec::with_capacity(original_end + rendered.len().saturating_sub(rendered_end));
    output.extend_from_slice(&original[..original_end]);
    output.extend_from_slice(&rendered[rendered_end..]);
    Ok(output)
}

fn fragment_root_start_tag_end(xml: &[u8]) -> Result<usize> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_) | Event::Empty(_)) => return Ok(reader.buffer_position() as usize),
            Ok(Event::Eof) => {
                return Err(template_error("paired template copy has no root element"));
            }
            Ok(_) => {}
            Err(error) => return Err(template_xml_error(error)),
        }
        buffer.clear();
    }
}

fn replay_program<'data>(
    program: &ParsedProgram<'_>,
    blocks: &[ExecutedBlock<'data>],
    current: &TemplateProgram,
    offset: usize,
    candidate: &mut Document,
    note_copies: Option<&mut Vec<NoteCopy>>,
    pending_drawings: &mut HashSet<u32>,
    refresh_root: bool,
    evaluation: &mut TemplateEvaluation<'data, '_>,
) -> Result<TemplateReplay> {
    if current.items.is_empty()
        || (note_copies.is_none()
            && current.required_paragraph
            && flattened_execution_items(blocks).is_empty())
    {
        return if let Some(source) = &current.source {
            let source = evaluation.operation.repeat_source(source)?;
            TemplateReplay::from_source(source, &mut evaluation.operation)
        } else {
            Ok(TemplateReplay {
                raw: current.raw.clone(),
                source: None,
            })
        };
    }
    let output = replay_program_blocks(
        program,
        blocks,
        current,
        offset,
        candidate,
        note_copies,
        pending_drawings,
        evaluation,
    )?;
    // Allocation canonicalizes bookmarks and relationships before publication.
    // Splice the prepared expansion into that fresh owner, preserving the
    // original callback locations already paired onto the root program.
    let base = if refresh_root {
        Cow::Owned(canonical_root_base(
            candidate,
            current,
            pending_drawings,
            evaluation,
        )?)
    } else {
        Cow::Borrowed(current)
    };
    let start = base.items.first().expect("nonempty source").range.start - base.ranges[0].start;
    let end = base.items.last().expect("nonempty source").range.end - base.ranges[0].start;
    if let Some(source) = &base.source {
        let mut pieces = vec![(Some(source.as_ref()), 0..start, &[][..])];
        if let Some(rendered) = &output.source {
            pieces.push((Some(rendered), 0..rendered.as_bytes().len(), &[][..]));
        } else {
            pieces.push((None, 0..0, output.raw.as_slice()));
        }
        pieces.push((Some(source.as_ref()), end..source.as_bytes().len(), &[][..]));
        let source = evaluation.operation.compose_sources(&pieces)?;
        let source = if current.required_paragraph {
            candidate.template_require_cell_source(
                source,
                &current.scope,
                &mut evaluation.operation,
            )?
        } else {
            source
        };
        TemplateReplay::from_source(source, &mut evaluation.operation)
    } else {
        let mut raw = base.raw.clone();
        raw.splice(start..end, output.raw);
        if current.required_paragraph {
            candidate.template_require_cell_paragraph(&mut raw, &current.scope)?;
        }
        Ok(TemplateReplay { raw, source: None })
    }
}

// Chart text has no authentic Word paragraph location. Prepare built-ins from
// original text operands, and refuse formatter mode before any callback.
fn prepare_chart_edits<'a>(
    document: &Document,
    data: &'a Value,
    versioned: bool,
    has_formatter: bool,
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<(Vec<PartRewrite>, usize)> {
    let mut edits = Vec::new();
    let mut count = 0;
    for (part, original) in document.template_chart_parts() {
        let mut reader = Reader::from_reader(original.as_slice());
        let mut replacements = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(element))
                    if matches_local_name(element.name().as_ref(), b"t")
                        || matches_local_name(element.name().as_ref(), b"v") =>
                {
                    let start = reader.buffer_position() as usize;
                    let text = oxml_core::xml_text::read_element_text(&mut reader, element.name());
                    let end = reader.buffer_position() as usize;
                    let closing = original[start..end]
                        .iter()
                        .rposition(|byte| *byte == b'<')
                        .ok_or_else(|| template_error("chart text closing source is missing"))?
                        + start;
                    let tokens = scalar_tokens(&text, versioned)?;
                    if tokens.is_empty() {
                        continue;
                    }
                    let mut rendered = text;
                    for token in tokens.into_iter().rev() {
                        let value = match token.value {
                            ScalarValue::Literal(value) => value.to_owned(),
                            ScalarValue::Expression(expression) => {
                                if has_formatter {
                                    return Err(template_error(&format!(
                                        "chart scalar in {part} has no original Word paragraph context for a formatter; use render_template_versioned"
                                    )));
                                }
                                count += 1;
                                if versioned {
                                    expression_text(&expression, data, &[], operation)?
                                } else {
                                    let path = expression
                                        .path
                                        .iter()
                                        .map(|step| match step {
                                            PathStep::Key(key) => key.clone(),
                                            PathStep::Index(_) => unreachable!(),
                                        })
                                        .collect::<Vec<_>>();
                                    match resolve_value(data, &[], &path, operation)? {
                                        ResolvedValue::Value(value) => {
                                            scalar_text(value, &path, operation)?
                                        }
                                        ResolvedValue::Deferred => unreachable!(),
                                    }
                                }
                            }
                        };
                        oxml_core::xml::reject_non_xml_characters(
                            "chart template replacement",
                            &value,
                        )?;
                        rendered.replace_range(token.range, &value);
                    }
                    replacements.push((
                        start..closing,
                        quick_xml::escape::escape(&rendered)
                            .into_owned()
                            .into_bytes(),
                    ));
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(error) => return Err(template_xml_error(error)),
            }
        }
        let mut rendered = original.clone();
        for (range, value) in replacements.into_iter().rev() {
            rendered.splice(range, value);
        }
        if rendered != original {
            edits.push((part, original, rendered));
        }
    }
    Ok((edits, count))
}

fn canonical_root_base(
    candidate: &mut Document,
    current: &TemplateProgram,
    pending_drawings: &HashSet<u32>,
    evaluation: &mut TemplateEvaluation<'_, '_>,
) -> Result<TemplateProgram> {
    // Dependency allocation changes identifier occupancy. Canonicalize before
    // reading the owner that publication will check again.
    candidate.prepare_staged_template_package(
        &mut evaluation.operation,
        &mut evaluation.lineage,
        &mut evaluation.sources,
    )?;
    // That prepare keeps only drawings already in the package. Repeated copies
    // still live in the unpublished replacement, so their ids stay reserved.
    candidate.retain_pending_drawing_ids(pending_drawings);
    let mut fresh = candidate
        .template_programs_scoped(
            &mut evaluation.operation,
            &evaluation.lineage,
            &mut evaluation.sources,
        )?
        .into_iter()
        .find(|program| {
            program.part == current.part
                && program.story.kind() == current.story.kind()
                && program.story.owner_index() == current.story.owner_index()
        })
        .ok_or_else(|| template_error("template physical owner disappeared during staging"))?;
    retain_original_template_locations(current, &mut fresh)?;
    Ok(fresh)
}

fn retain_original_template_locations(
    original: &TemplateProgram,
    current: &mut TemplateProgram,
) -> Result<()> {
    if original.part != current.part
        || original.story.kind() != current.story.kind()
        || original.ranges.len() != current.ranges.len()
        || original.table_rows != current.table_rows
        || original.items.len() != current.items.len()
    {
        return Err(template_error(
            "source normalization changed template owner correspondence",
        ));
    }
    for (before, after) in original.items.iter().zip(&mut current.items) {
        if before.kind != after.kind || before.children.len() != after.children.len() {
            return Err(template_error(
                "source normalization changed template item correspondence",
            ));
        }
        after.location = before.location.clone();
        for (before, after) in before.children.iter().zip(&mut after.children) {
            retain_original_template_locations(before, after)?;
        }
    }
    Ok(())
}

fn prepare_unowned_textbox_edits<'a>(
    document: &Document,
    programs: &[TemplateProgram],
    data: &'a Value,
    versioned: bool,
    has_formatter: bool,
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
    lineage: &TemplateWordLineage<'a>,
    main_sources: &mut crate::document::TemplateMainSources,
) -> Result<(Vec<ScopedPartRewrite>, usize)> {
    fn exclusions(
        program: &TemplateProgram,
        parts: &mut BTreeMap<String, Vec<std::ops::Range<usize>>>,
    ) {
        let physical_textbox = program.story.kind() == crate::StoryKind::TextBox
            && matches!(Reader::from_reader(program.raw.as_slice()).read_event(), Ok(Event::Start(ref start) | Event::Empty(ref start)) if matches_local_name(start.name().as_ref(),b"txbxContent"));
        if physical_textbox {
            parts
                .entry(program.part.clone())
                .or_default()
                .extend(program.ranges.iter().cloned());
        }
        for item in &program.items {
            for child in &item.children {
                exclusions(child, parts);
            }
        }
    }
    let mut excluded = BTreeMap::new();
    for program in programs {
        exclusions(program, &mut excluded);
    }
    let mut edits = Vec::new();
    let mut count = 0;
    for (part, original) in
        document.template_raw_sources_scoped(operation, lineage, main_sources)?
    {
        let ranges = excluded.entry(part.clone()).or_default();
        // Word writes one text box twice, and a box can contain another.
        // Those physical ranges overlap. The raw walker accepts only a
        // sorted list whose members do not overlap, so each overlap becomes
        // one covering range. A range that merely touches the previous end
        // stays separate.
        ranges.retain(|range| range.start < range.end);
        ranges.sort_by_key(|range| (range.start, std::cmp::Reverse(range.end)));
        let mut covered: Vec<std::ops::Range<usize>> = Vec::new();
        for range in ranges.iter().cloned() {
            if let Some(last) = covered.last_mut()
                && range.start < last.end
            {
                last.end = last.end.max(range.end);
            } else {
                covered.push(range);
            }
        }
        *ranges = covered;
        let (discovered, _, sources) =
            operation.replace_original_template_literals(&original, ranges, &[])?;
        let mut operands = BTreeMap::<String, (String, Vec<&'a Value>)>::new();
        for source in sources {
            for token in scalar_tokens(&source, versioned)? {
                let mut dependencies = Vec::new();
                let value = match token.value {
                    ScalarValue::Literal(value) => value.to_owned(),
                    ScalarValue::Expression(expression) => {
                        if has_formatter {
                            return Err(template_error(&format!(
                                "raw textbox scalar in {part} has no checked original Word paragraph context for a formatter; use render_template_versioned"
                            )));
                        }
                        count += 1;
                        if versioned {
                            expression_text_in(
                                &expression,
                                data,
                                &[],
                                operation,
                                Some(&mut dependencies),
                            )?
                        } else {
                            let path = expression
                                .path
                                .iter()
                                .map(|step| match step {
                                    PathStep::Key(key) => key.clone(),
                                    PathStep::Index(_) => unreachable!(),
                                })
                                .collect::<Vec<_>>();
                            match resolve_value(data, &[], &path, operation)? {
                                ResolvedValue::Value(value) => {
                                    dependencies.push(value);
                                    scalar_text(value, &path, operation)?
                                }
                                ResolvedValue::Deferred => unreachable!(),
                            }
                        }
                    }
                };
                oxml_core::xml::reject_non_xml_characters(
                    "raw textbox template replacement",
                    &value,
                )?;
                operands.insert(source[token.range].to_owned(), (value, dependencies));
            }
        }
        if operands.is_empty() {
            document.template_retain_raw_source(&part, discovered, main_sources)?;
            continue;
        }
        let operands = operands
            .iter()
            .map(|(key, (value, dependencies))| {
                (key.as_str(), value.as_str(), dependencies.as_slice())
            })
            .collect::<Vec<_>>();
        let (rendered, _, _) =
            operation.replace_original_template_literals(&discovered, ranges, &operands)?;
        if rendered.as_bytes() != original.as_bytes() {
            let retained = operation.repeat_source(&original)?;
            let original = operation.finish_source(retained)?;
            edits.push((part, original, rendered));
        } else {
            document.template_retain_raw_source(&part, rendered, main_sources)?;
        }
    }
    Ok((edits, count))
}

#[allow(clippy::type_complexity)] // Same non-static formatter contract as the evaluation field.
pub(crate) fn render_grammar(
    document: &mut Document,
    data: &Value,
    versioned: bool,
    formatter: Option<
        &mut dyn FnMut(
            &TemplateFormatContext<'_>,
            &mut crate::MailMergeFormattedText,
        ) -> Result<()>,
    >,
) -> Result<usize> {
    document.validate_model_resource_limits()?;
    let records = document.template_word_record_inputs();
    let mut current_paragraphs = Vec::new();
    crate::document::visit_body_paragraphs(
        &document.document.body.content,
        false,
        &mut |paragraph| {
            current_paragraphs.push(paragraph);
        },
    );
    for note in &document.footnotes.footnotes {
        for paragraph in &note.paragraphs {
            crate::document::visit_paragraph(paragraph, false, &mut |paragraph| {
                current_paragraphs.push(paragraph);
            });
        }
    }
    if let Some(glossary) = &document.glossary {
        for part in &glossary.doc_parts {
            crate::document::visit_body_paragraphs(&part.body.content, false, &mut |paragraph| {
                current_paragraphs.push(paragraph);
            });
        }
    }
    if let Some(comments) = &document.comments {
        for comment in &comments.comments {
            current_paragraphs.extend(comment.paragraphs.iter());
        }
    }
    let mut current_bodies = vec![&document.document.body];
    let mut current_word_sources = Vec::new();
    if let Some(glossary) = &document.glossary {
        current_word_sources.push(glossary.retained_xml());
        for part in &glossary.doc_parts {
            current_bodies.push(&part.body);
            current_word_sources.push(part.retained_xml());
        }
    }
    if let Some(comments) = &document.comments {
        current_word_sources.extend(comments.extra_xml.iter().map(|(_, raw)| raw.as_slice()));
        for comment in &comments.comments {
            current_word_sources.extend(comment.extra_xml.iter().map(|(_, raw)| raw.as_slice()));
        }
    }
    current_word_sources.extend(records.records());
    let mut operation = rdocx_oxml::document::ScopedWordOperation::with_template_word_inputs(
        document
            .package
            .parts
            .iter()
            .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
            .collect(),
        current_paragraphs,
        current_bodies,
        Vec::new(),
        Vec::new(),
        current_word_sources,
        data,
    )?;
    let (mut candidate, mut lineage) = TemplateWordLineage::clone_for_staging(document);
    lineage.records = Some(&records);
    let mut main_sources = crate::document::TemplateMainSources::default();
    let originals =
        candidate.template_programs_scoped(&mut operation, &lineage, &mut main_sources)?;
    candidate.prepare_staged_template_package(&mut operation, &mut lineage, &mut main_sources)?;
    let mut sources =
        candidate.template_programs_scoped(&mut operation, &lineage, &mut main_sources)?;
    if originals.len() != sources.len() {
        return Err(template_error(
            "source normalization changed template root correspondence",
        ));
    }
    for (original, current) in originals.iter().zip(&mut sources) {
        retain_original_template_locations(original, current)?;
    }
    let (raw_textboxes, raw_count) = prepare_unowned_textbox_edits(
        &candidate,
        &sources,
        data,
        versioned,
        formatter.is_some(),
        &mut operation,
        &lineage,
        &mut main_sources,
    )?;
    let raw_changed = !raw_textboxes.is_empty();
    for (part, original, rendered) in raw_textboxes {
        candidate.template_publish_scalar_source(
            &part,
            &original,
            rendered,
            &mut operation,
            &lineage,
            &mut main_sources,
        )?;
    }
    // Reacquire fresh edit ranges after original-only raw scalar edits. Retain
    // original callback locations through the same checked item correspondence.
    if raw_changed {
        let mut refreshed =
            candidate.template_programs_scoped(&mut operation, &lineage, &mut main_sources)?;
        if originals.len() != refreshed.len() {
            return Err(template_error(
                "raw scalar publication changed checked template roots",
            ));
        }
        for (original, current) in originals.iter().zip(&mut refreshed) {
            retain_original_template_locations(original, current)?;
        }
        sources = refreshed;
    }
    let (charts, chart_count) = prepare_chart_edits(
        &candidate,
        data,
        versioned,
        formatter.is_some(),
        &mut operation,
    )?;
    let main_part = candidate.doc_part_name.clone();
    let programs = sources
        .iter()
        .map(|source| {
            let actual_source = if source.part == main_part {
                main_sources.source.as_ref()
            } else {
                candidate
                    .package
                    .get_part_shared(&source.part)
                    .and_then(|bytes| main_sources.related_source(&source.part, bytes))
            };
            let mut scoped = actual_source.map(|source| (&mut operation, source.as_ref()));
            parse_program(source, versioned, false, &mut scoped)
        })
        .collect::<Result<Vec<_>>>()?;
    for program in &programs {
        let actual_source = if program.source.part == main_part {
            main_sources.source.as_ref()
        } else {
            candidate
                .package
                .get_part_shared(&program.source.part)
                .and_then(|bytes| main_sources.related_source(&program.source.part, bytes))
        };
        let mut scoped = actual_source.map(|source| (&mut operation, source.as_ref()));
        validate_program_numbering(program, &program.blocks, &candidate, false, &mut scoped)?;
    }
    let mut evaluation = TemplateEvaluation {
        data,
        operation,
        lineage,
        sources: main_sources,
        versioned,
        occurrence: 0,
        formatter: None,
    };
    let mut executions = Vec::new();
    for program in &programs {
        executions.push(ProgramExecution {
            blocks: evaluate_program_blocks(
                program,
                &program.blocks,
                &mut Vec::new(),
                &mut evaluation,
                true,
            )?,
        });
    }
    // Structure is prepared without touching scalar operands. Import allocation
    // facts connect copied physical notes to already evaluated original programs.
    let key = |source: &TemplateProgram| {
        (
            source.story.kind(),
            source.part.clone(),
            source.story.owner_index(),
        )
    };
    let mut bindings = programs
        .iter()
        .enumerate()
        .zip(&executions)
        .map(|((index, program), execution)| {
            Ok((
                key(program.source),
                (index, execution.checked_clone(&mut evaluation.operation)?),
            ))
        })
        .collect::<Result<HashMap<_, _>>>()?;
    let mut note_origins = HashMap::new();
    for (kind, part, id, owner) in
        candidate.template_note_identities(Some(&mut crate::document::TemplateMainWriter {
            operation: &mut evaluation.operation,
            lineage: &evaluation.lineage,
            sources: &mut evaluation.sources,
        }))?
    {
        let (index, _) = bindings
            .get(&(kind, part.clone(), owner))
            .ok_or_else(|| template_error("original note has no prepared template owner"))?;
        note_origins.insert((kind, part, id), *index);
    }
    let mut allocations = Vec::new();
    let mut pending_drawings = HashSet::new();
    let mut structural_edits = Vec::new();
    for (program, execution) in programs.iter().zip(&executions) {
        let rendered = replay_program(
            program,
            &execution.blocks,
            program.source,
            0,
            &mut candidate,
            Some(&mut allocations),
            &mut pending_drawings,
            true,
            &mut evaluation,
        )?;
        let fresh = canonical_root_base(
            &mut candidate,
            program.source,
            &pending_drawings,
            &mut evaluation,
        )?;
        structural_edits.push((fresh, rendered));
    }
    let mut allocated = 0;
    while allocated < allocations.len() {
        let end = allocations.len();
        let current_sources = candidate
            .template_programs_scoped(
                &mut evaluation.operation,
                &evaluation.lineage,
                &mut evaluation.sources,
            )?
            .into_iter()
            .map(|source| (key(&source), source))
            .collect::<HashMap<_, _>>();
        let identities = candidate
            .template_note_identities(Some(&mut crate::document::TemplateMainWriter {
                operation: &mut evaluation.operation,
                lineage: &evaluation.lineage,
                sources: &mut evaluation.sources,
            }))?
            .into_iter()
            .map(|(kind, part, id, owner)| ((kind, part, id), owner))
            .collect::<HashMap<_, _>>();
        // Register the complete allocation batch first. Nested references use
        // this same actual old-to-new map, never numeric or hash-map order.
        for (kind, source_part, old_id, destination_part, new_id) in &allocations[allocated..end] {
            let origin = *note_origins
                .get(&(*kind, source_part.clone(), *old_id))
                .ok_or_else(|| template_error("copied note has no original prepared provenance"))?;
            note_origins.insert((*kind, destination_part.clone(), *new_id), origin);
        }
        // Detach this batch. Replaying a copied note appends further allocations.
        let batch = allocations[allocated..end].to_vec();
        for (kind, _, _, destination_part, new_id) in batch {
            let origin = note_origins[&(kind, destination_part.clone(), new_id)];
            let owner = identities
                .get(&(kind, destination_part.clone(), new_id))
                .ok_or_else(|| template_error("allocated note has no checked physical owner"))?;
            let current = current_sources
                .get(&(kind, destination_part.clone(), *owner))
                .ok_or_else(|| template_error("allocated note source is missing"))?;
            let execution = executions[origin].checked_clone(&mut evaluation.operation)?;
            let rendered = replay_program(
                &programs[origin],
                &execution.blocks,
                current,
                0,
                &mut candidate,
                Some(&mut allocations),
                &mut pending_drawings,
                true,
                &mut evaluation,
            )?;
            if bindings.insert(key(current), (origin, execution)).is_some() {
                return Err(template_error("allocated note source was prepared twice"));
            }
            let fresh =
                canonical_root_base(&mut candidate, current, &pending_drawings, &mut evaluation)?;
            structural_edits.push((fresh, rendered));
        }
        allocated = end;
    }
    let mut changed = !charts.is_empty() || raw_changed;
    for (source, rendered) in structural_edits {
        let source =
            canonical_root_base(&mut candidate, &source, &pending_drawings, &mut evaluation)?;
        changed |= rendered.raw != source.raw;
        if let Some(produced) = rendered.source {
            candidate.template_publish_program_source(
                &source,
                produced,
                &mut evaluation.operation,
                &evaluation.lineage,
                &mut evaluation.sources,
            )?;
        } else {
            evaluation.sources.invalidate_model();
            candidate.template_publish_program(&source, rendered.raw)?;
        }
    }
    evaluation.formatter = formatter;
    let expanded = candidate.template_programs_scoped(
        &mut evaluation.operation,
        &evaluation.lineage,
        &mut evaluation.sources,
    )?;
    for current in expanded {
        let (origin, mut execution) = bindings
            .remove(&key(&current))
            .ok_or_else(|| template_error("expanded owner has no prepared source provenance"))?;
        let mut position = 0;
        format_program_execution(
            &programs[origin],
            &mut execution.blocks,
            &current,
            &mut position,
            &mut evaluation,
        )?;
        if position != current.items.len() && !(position == 0 && current.required_paragraph) {
            return Err(template_error("expanded owner has unbound source items"));
        }
        let rendered = replay_program(
            &programs[origin],
            &execution.blocks,
            &current,
            0,
            &mut candidate,
            None,
            &mut pending_drawings,
            false,
            &mut evaluation,
        )?;
        changed |= rendered.raw != current.raw;
        if let Some(produced) = rendered.source {
            candidate.template_publish_program_source(
                &current,
                produced,
                &mut evaluation.operation,
                &evaluation.lineage,
                &mut evaluation.sources,
            )?;
        } else {
            evaluation.sources.invalidate_model();
            candidate.template_publish_program(&current, rendered.raw)?;
        }
    }
    if !bindings.is_empty() {
        return Err(template_error("prepared template owner disappeared"));
    }
    let count = evaluation.occurrence + chart_count + raw_count;
    if !changed {
        return Ok(count);
    }
    for (part, original, rendered) in charts {
        candidate.template_publish_scalar_part(&part, &original, rendered)?;
    }
    let prepared = candidate.prepare_and_reopen_staged()?;
    drop(evaluation);
    document.commit_template(prepared);
    Ok(count)
}

fn json_string_at(source: &str, cursor: &mut usize) -> Result<String> {
    let start = *cursor;
    if source.as_bytes().get(start) != Some(&b'"') {
        return Err(template_error(
            "filter arguments and quoted keys require JSON strings",
        ));
    }
    *cursor += 1;
    let mut escaped = false;
    while let Some(&byte) = source.as_bytes().get(*cursor) {
        *cursor += 1;
        if escaped {
            escaped = false;
        } else if byte == b'\\' {
            escaped = true;
        } else if byte == b'"' {
            return serde_json::from_str(&source[start..*cursor])
                .map_err(|_| template_error("invalid JSON string in template"));
        }
    }
    Err(template_error("unclosed JSON string in template"))
}

fn skip_space(source: &str, cursor: &mut usize) {
    while let Some(character) = source[*cursor..].chars().next() {
        if !character.is_whitespace() {
            break;
        }
        *cursor += character.len_utf8();
    }
}

fn parse_expression(source: &str) -> Result<Expression> {
    let source = source.trim();
    let mut cursor = 0;
    let mut path = Vec::new();
    let mut need_key = true;
    loop {
        if source.as_bytes().get(cursor) == Some(&b'[') && !need_key {
            cursor += 1;
            skip_space(source, &mut cursor);
            let step = if source.as_bytes().get(cursor) == Some(&b'"') {
                PathStep::Key(json_string_at(source, &mut cursor)?)
            } else {
                let start = cursor;
                while source
                    .as_bytes()
                    .get(cursor)
                    .is_some_and(u8::is_ascii_digit)
                {
                    cursor += 1;
                }
                if start == cursor {
                    return Err(template_error(
                        "brackets require an array index or quoted object key",
                    ));
                }
                PathStep::Index(
                    source[start..cursor]
                        .parse()
                        .map_err(|_| template_error("array index exceeds supported range"))?,
                )
            };
            skip_space(source, &mut cursor);
            if source.as_bytes().get(cursor) != Some(&b']') {
                return Err(template_error("unclosed path bracket"));
            }
            cursor += 1;
            path.push(step);
        } else if need_key && source.as_bytes().get(cursor) == Some(&b'[') && path.is_empty() {
            // A quoted root key admits punctuation without an invented root name.
            need_key = false;
            continue;
        } else if need_key {
            let start = cursor;
            while let Some(character) = source[cursor..].chars().next() {
                if character.is_whitespace()
                    || matches!(
                        character,
                        '.' | '[' | ']' | '|' | '"' | '\'' | '{' | '}' | '%' | '(' | ')' | '\0'
                    )
                {
                    break;
                }
                cursor += character.len_utf8();
            }
            if cursor == start {
                return Err(template_error("expected a template property name"));
            }
            path.push(PathStep::Key(source[start..cursor].into()));
            need_key = false;
        } else if source.as_bytes().get(cursor) == Some(&b'.') {
            cursor += 1;
            need_key = true;
        } else {
            break;
        }
    }
    if need_key || path.is_empty() {
        return Err(template_error("incomplete template path"));
    }
    let spelling = source[..cursor].to_owned();
    let mut filters = Vec::new();
    loop {
        skip_space(source, &mut cursor);
        if cursor == source.len() {
            break;
        }
        if source.as_bytes().get(cursor) != Some(&b'|') {
            return Err(template_error("unexpected text after template path"));
        }
        cursor += 1;
        skip_space(source, &mut cursor);
        let start = cursor;
        while source
            .as_bytes()
            .get(cursor)
            .is_some_and(u8::is_ascii_alphabetic)
        {
            cursor += 1;
        }
        let name = &source[start..cursor];
        skip_space(source, &mut cursor);
        let filter = match name {
            "trim" => Filter::Trim,
            "upper" => Filter::Upper,
            "lower" => Filter::Lower,
            "default" | "join" => {
                if source.as_bytes().get(cursor) != Some(&b'(') {
                    return Err(template_error("filter requires one JSON string argument"));
                }
                cursor += 1;
                skip_space(source, &mut cursor);
                let argument = json_string_at(source, &mut cursor)?;
                skip_space(source, &mut cursor);
                if source.as_bytes().get(cursor) != Some(&b')') {
                    return Err(template_error(
                        "filter requires exactly one JSON string argument",
                    ));
                }
                cursor += 1;
                if name == "default" {
                    Filter::Default(argument)
                } else {
                    Filter::Join(argument)
                }
            }
            _ => return Err(template_error("unknown template filter")),
        };
        filters.push(filter);
    }
    Ok(Expression {
        spelling,
        path,
        filters,
    })
}

fn traverse_expression<'a>(
    mut value: &'a Value,
    path: &[PathStep],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<Option<&'a Value>> {
    for step in path {
        let next = match step {
            PathStep::Key(key) => {
                value
                    .as_object()
                    .ok_or_else(|| template_error("template path crosses a non-object value"))?;
                operation.template_datum_key(value, key)?
            }
            PathStep::Index(index) => {
                value
                    .as_array()
                    .ok_or_else(|| template_error("template path crosses a non-array value"))?;
                operation.template_datum_index(value, *index)?
            }
        };
        let Some(next) = next else {
            return Ok(None);
        };
        value = next;
    }
    Ok(Some(value))
}

fn expression_value<'a>(
    expression: &Expression,
    root: &'a Value,
    scopes: &[Scope<'a>],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<Option<&'a Value>> {
    if let Some(PathStep::Key(first)) = expression.path.first() {
        for scope in scopes.iter().rev() {
            if &scope.name == first {
                return scope.value.map_or(Ok(None), |value| {
                    traverse_expression(value, &expression.path[1..], operation)
                });
            }
        }
        for scope in scopes.iter().rev() {
            if scope
                .value
                .and_then(Value::as_object)
                .is_some_and(|object| object.contains_key(first))
            {
                return traverse_expression(
                    scope.value.expect("selected non-deferred scope"),
                    &expression.path,
                    operation,
                );
            }
        }
    }
    traverse_expression(root, &expression.path, operation)
}

fn expression_text<'a>(
    expression: &Expression,
    root: &'a Value,
    scopes: &[Scope<'a>],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<String> {
    expression_text_in(expression, root, scopes, operation, None)
}

// A non-emitting expression keeps only actual borrowed inputs and its filter
// suffix. It never registers scalar copy credit or constructs replacement text.
enum ValidationText<'data, 'expression> {
    Scalar(&'data Value),
    Literal(&'expression str),
    Joined(&'data [Value], &'expression str),
}

enum ExpressionValue<'data, 'expression> {
    Borrowed(&'data Value),
    Emitted(String),
    Validation(ValidationText<'data, 'expression>, usize),
}

fn expression_text_in<'a>(
    expression: &Expression,
    root: &'a Value,
    scopes: &[Scope<'a>],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
    inputs: Option<&mut Vec<&'a Value>>,
) -> Result<String> {
    Ok(
        evaluate_expression(expression, root, scopes, operation, inputs, true)?
            .expect("emitting expression returns its text"),
    )
}

fn evaluate_expression<'a>(
    expression: &Expression,
    root: &'a Value,
    scopes: &[Scope<'a>],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
    mut inputs: Option<&mut Vec<&'a Value>>,
    emit: bool,
) -> Result<Option<String>> {
    let selected = expression_value(expression, root, scopes, operation)?;
    if emit
        && let Some(inputs) = inputs.as_deref_mut()
        && let Some(value) = selected
        && !value.is_array()
        && !value.is_object()
    {
        inputs.push(value);
    }
    let mut value = selected.map(ExpressionValue::Borrowed);
    for (filter_index, filter) in expression.filters.iter().enumerate() {
        match filter {
            Filter::Default(fallback) => {
                if value
                    .as_ref()
                    .is_none_or(|value| matches!(value, ExpressionValue::Borrowed(Value::Null)))
                {
                    if emit {
                        if let Some(inputs) = inputs.as_deref_mut() {
                            inputs.clear();
                        }
                        operation.reserve_template_text_copy(fallback.len())?;
                        value = Some(ExpressionValue::Emitted(fallback.clone()));
                    } else {
                        value = Some(ExpressionValue::Validation(
                            ValidationText::Literal(fallback),
                            filter_index + 1,
                        ));
                    }
                }
            }
            Filter::Join(separator) => {
                let array = match value.as_ref() {
                    Some(ExpressionValue::Borrowed(array @ Value::Array(_))) => *array,
                    _ => return Err(template_error("join requires an array of scalar values")),
                };
                let Value::Array(values) = array else {
                    unreachable!("array kind was selected")
                };
                if emit {
                    if let Some(inputs) = inputs.as_deref_mut() {
                        inputs.clear();
                    }
                    let items = (0..values.len())
                        .map(|index| {
                            let input = operation
                                .template_datum_index(array, index)?
                                .expect("index belongs to the immutable selected array");
                            if let Some(inputs) = inputs.as_deref_mut() {
                                inputs.push(input);
                            }
                            scalar_text(input, &[], operation)
                        })
                        .collect::<Result<Vec<_>>>()?;
                    let bytes = separator
                        .len()
                        .checked_mul(items.len().saturating_sub(1))
                        .and_then(|separators| {
                            items
                                .iter()
                                .try_fold(separators, |bytes, text| bytes.checked_add(text.len()))
                        })
                        .ok_or_else(|| template_error("template joined text length overflow"))?;
                    operation.reserve_template_text_copy(bytes)?;
                    value = Some(ExpressionValue::Emitted(items.join(separator)));
                } else {
                    // Validate every scalar kind before checking the final text.
                    // An earlier invalid character cannot hide a later type error.
                    for index in 0..values.len() {
                        let input = operation
                            .template_datum_index(array, index)?
                            .expect("index belongs to the immutable selected array");
                        require_scalar(input, &[])?;
                    }
                    value = Some(ExpressionValue::Validation(
                        ValidationText::Joined(values, separator),
                        filter_index + 1,
                    ));
                }
            }
            Filter::Trim | Filter::Upper | Filter::Lower => {
                let current = value
                    .take()
                    .ok_or_else(|| template_error("missing template path"))?;
                if !emit {
                    let (source, start) = match current {
                        ExpressionValue::Borrowed(input) => {
                            require_scalar(input, &[])?;
                            (ValidationText::Scalar(input), filter_index)
                        }
                        ExpressionValue::Validation(source, start) => (source, start),
                        ExpressionValue::Emitted(_) => unreachable!("validation never emits text"),
                    };
                    value = Some(ExpressionValue::Validation(source, start));
                    continue;
                }
                let text = match current {
                    ExpressionValue::Emitted(text) => Cow::Owned(text),
                    ExpressionValue::Borrowed(input @ Value::String(text)) => {
                        operation.template_scalar_input_bytes(input)?;
                        Cow::Borrowed(text.as_str())
                    }
                    ExpressionValue::Borrowed(input) => {
                        Cow::Owned(scalar_text(input, &[], operation)?)
                    }
                    ExpressionValue::Validation(_, _) => {
                        unreachable!("emitting expression has no validation projection")
                    }
                };
                if matches!(&text, Cow::Borrowed(_))
                    || matches!(filter, Filter::Upper | Filter::Lower)
                {
                    let bytes = match filter {
                        Filter::Trim => Some(text.trim().len()),
                        Filter::Upper => text
                            .chars()
                            .flat_map(char::to_uppercase)
                            .try_fold(0usize, |bytes, c| bytes.checked_add(c.len_utf8())),
                        Filter::Lower => text
                            .chars()
                            .flat_map(char::to_lowercase)
                            .try_fold(0usize, |bytes, c| bytes.checked_add(c.len_utf8())),
                        _ => unreachable!(),
                    }
                    .ok_or_else(|| template_error("template case text length overflow"))?;
                    operation.reserve_template_text_copy(bytes)?;
                }
                let text = match (filter, text) {
                    (Filter::Trim, Cow::Borrowed(text)) => text.trim().to_owned(),
                    (Filter::Trim, Cow::Owned(mut text)) => {
                        let end = text.trim_end().len();
                        text.truncate(end);
                        let start = text.len() - text.trim_start().len();
                        text.drain(..start);
                        text
                    }
                    (Filter::Upper, text) => text.to_uppercase(),
                    (Filter::Lower, text) => text.to_lowercase(),
                    _ => unreachable!(),
                };
                value = Some(ExpressionValue::Emitted(text));
            }
        }
    }
    match value {
        Some(ExpressionValue::Emitted(text)) => Ok(Some(text)),
        Some(ExpressionValue::Borrowed(value)) if emit => {
            scalar_text(value, &[], operation).map(Some)
        }
        Some(ExpressionValue::Borrowed(value)) => {
            require_scalar(value, &[])?;
            validate_expression_text(ValidationText::Scalar(value), &[])?;
            Ok(None)
        }
        Some(ExpressionValue::Validation(source, start)) => {
            validate_expression_text(source, &expression.filters[start..])?;
            Ok(None)
        }
        None => Err(template_error("missing template path")),
    }
}

// Case conversion preserves whitespace. Moving trim to this validation sink
// therefore preserves the final scalar count and invalid-character position.
// Contextual lowercasing changes final sigma's identity, but not either fact.
struct ExpressionTextValidator<'expression> {
    filters: &'expression [Filter],
    pending: Vec<(usize, char)>,
    trim: bool,
    characters: usize,
    whitespace: usize,
    whitespace_invalid: Option<(usize, char)>,
    first_invalid: Option<(usize, char)>,
}

impl ExpressionTextValidator<'_> {
    fn character(&mut self, character: char) {
        self.pending.push((0, character));
        while let Some((index, character)) = self.pending.pop() {
            if let Some(filter) = self.filters.get(index) {
                match filter {
                    Filter::Upper => {
                        for character in character.to_uppercase().rev() {
                            self.pending.push((index + 1, character));
                        }
                    }
                    Filter::Lower => {
                        for character in character.to_lowercase().rev() {
                            self.pending.push((index + 1, character));
                        }
                    }
                    Filter::Trim | Filter::Default(_) => {
                        self.pending.push((index + 1, character));
                    }
                    Filter::Join(_) => unreachable!("join cannot follow string conversion"),
                }
            } else {
                self.final_character(character);
            }
        }
    }

    fn final_character(&mut self, character: char) {
        if self.trim && character.is_whitespace() {
            if self.characters != 0 {
                self.whitespace += 1;
                if self.whitespace_invalid.is_none()
                    && !oxml_core::xml::is_xml_1_0_character(character)
                {
                    self.whitespace_invalid = Some((self.whitespace, character));
                }
            }
            return;
        }
        if let Some((position, invalid)) = self.whitespace_invalid.take()
            && self.first_invalid.is_none()
        {
            self.first_invalid = Some((self.characters + position, invalid));
        }
        self.characters += self.whitespace + 1;
        self.whitespace = 0;
        if self.first_invalid.is_none() && !oxml_core::xml::is_xml_1_0_character(character) {
            self.first_invalid = Some((self.characters, character));
        }
    }

    fn scalar(&mut self, value: &Value) -> Result<()> {
        match value {
            Value::String(text) => {
                std::fmt::Write::write_str(self, text)
                    .expect("character validation sink does not fail formatting");
            }
            Value::Number(value) => {
                std::fmt::write(self, format_args!("{value}"))
                    .map_err(|_| template_error("canonical number formatting failed"))?;
            }
            Value::Bool(value) => {
                std::fmt::Write::write_str(self, if *value { "true" } else { "false" })
                    .expect("character validation sink does not fail formatting");
            }
            Value::Null => {}
            Value::Array(_) | Value::Object(_) => {
                unreachable!("scalar kind was checked before character validation")
            }
        }
        Ok(())
    }
}

impl std::fmt::Write for ExpressionTextValidator<'_> {
    fn write_str(&mut self, text: &str) -> std::fmt::Result {
        for character in text.chars() {
            self.character(character);
        }
        Ok(())
    }
}

fn validate_expression_text(source: ValidationText<'_, '_>, filters: &[Filter]) -> Result<()> {
    let mut validator = ExpressionTextValidator {
        filters,
        pending: Vec::new(),
        trim: filters.iter().any(|filter| matches!(filter, Filter::Trim)),
        characters: 0,
        whitespace: 0,
        whitespace_invalid: None,
        first_invalid: None,
    };
    match source {
        ValidationText::Scalar(value) => validator.scalar(value)?,
        ValidationText::Literal(text) => {
            std::fmt::Write::write_str(&mut validator, text)
                .expect("character validation sink does not fail formatting");
        }
        ValidationText::Joined(values, separator) => {
            for (index, value) in values.iter().enumerate() {
                if index != 0 {
                    std::fmt::Write::write_str(&mut validator, separator)
                        .expect("character validation sink does not fail formatting");
                }
                validator.scalar(value)?;
            }
        }
    }
    if let Some((position, character)) = validator.first_invalid {
        return Err(oxml_core::OxmlError::InvalidValue(format!(
            "template replacement holds U+{:04X} at character {}, a character XML 1.0 cannot carry",
            u32::from(character),
            position,
        ))
        .into());
    }
    Ok(())
}

#[cfg(test)]
#[test]
fn versioned_template_expression_paths_and_filters_are_typed() {
    let data = serde_json::json!({"person":{"name":"  Straße  ","values":[1,false,null,"é"]},"a}}|b":"literal", "zero":0, "empty":"", "wrong":3, "greek":"ΟΣ", "turkish":"İ"});
    for (source, expected) in [
        ("person.name | trim | upper", "STRASSE"),
        ("person.name | upper | trim | trim", "STRASSE"),
        ("greek | lower", "ος"),
        ("turkish | lower", "i\u{307}"),
        (r#"["a}}|b"]"#, "literal"),
        ("person.values[1]", "false"),
        (r#"person.values | join("/")"#, "1/false//é"),
        (r#"absent | default("fallback")"#, "fallback"),
        (r#"zero | default("fallback")"#, "0"),
        (r#"empty | default("fallback")"#, ""),
    ] {
        let word = format!(
            r#"<w:p xmlns:w="{}"><w:r><w:t>{source}</w:t></w:r></w:p>"#,
            rdocx_oxml::namespace::W_NS,
        );
        let mut operation = expression_test_operation(&data, word.as_bytes());
        assert_eq!(
            expression_text(
                &parse_expression(source).unwrap(),
                &data,
                &[],
                &mut operation
            )
            .unwrap(),
            expected,
            "{source}"
        );
    }
    for source in [
        r#"wrong.child | default("fallback")"#,
        r#"person[0] | default("fallback")"#,
        r#"person.values | join(1)"#,
        "person.name | trim()",
        "person.name | unknown",
        "person.values[-1]",
        "person..name",
    ] {
        assert!(
            parse_expression(source)
                .and_then(|expression| expression_text(
                    &expression,
                    &data,
                    &[],
                    &mut rdocx_oxml::document::ScopedWordOperation::with_template_inputs(
                        vec![],
                        vec![],
                        &data
                    )
                    .unwrap()
                ))
                .is_err(),
            "{source}"
        );
    }
}

#[test]
fn template_join_reserves_the_final_copy_before_allocating() {
    let data = serde_json::json!({"values": vec![Value::Null; 1000]});
    let expression = parse_expression(r#"values | join(",")"#).unwrap();
    for remaining in [998, 999] {
        let mut operation =
            rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
                .unwrap();
        let values = operation
            .template_datum_key(&data, "values")
            .unwrap()
            .unwrap();
        for index in 0..1000 {
            let input = operation
                .template_datum_index(values, index)
                .unwrap()
                .unwrap();
            assert_eq!(operation.template_scalar_input_bytes(input).unwrap(), 0);
        }
        let allowance = 1000 * 64 * (rdocx_oxml::namespace::W_NS.len() + 64);
        operation
            .reserve_template_text_copy(allowance - remaining)
            .unwrap();
        let result = expression_text(&expression, &data, &[], &mut operation);
        if remaining == 998 {
            assert!(
                result.is_err(),
                "the final 999-byte joined copy must be reserved"
            );
        } else {
            assert_eq!(result.unwrap(), ",".repeat(999));
            assert!(
                operation.reserve_template_text_copy(1).is_err(),
                "the exact final copy consumed the remaining allowance"
            );
        }
    }
}

#[test]
fn template_case_filters_reserve_owned_outputs_but_trim_reuses_them() {
    let data = serde_json::json!({"value": "x".repeat(10_000)});
    for (source, remaining, succeeds) in [
        ("value | upper | lower", 15_000, false),
        ("value | lower | upper", 15_000, false),
        ("value | upper | trim | trim", 10_000, true),
    ] {
        let mut operation =
            rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
                .unwrap();
        let input = operation
            .template_datum_key(&data, "value")
            .unwrap()
            .unwrap();
        assert_eq!(
            operation.template_scalar_input_bytes(input).unwrap(),
            10_000
        );
        let allowance = (64 + 10_000) * (rdocx_oxml::namespace::W_NS.len() + 64);
        operation
            .reserve_template_text_copy(allowance - remaining)
            .unwrap();
        let result = expression_text(
            &parse_expression(source).unwrap(),
            &data,
            &[],
            &mut operation,
        );
        assert_eq!(result.is_ok(), succeeds, "{source}");
        if succeeds {
            assert_eq!(result.unwrap(), "X".repeat(10_000));
            assert!(operation.reserve_template_text_copy(1).is_err());
        }
    }
}

#[test]
fn template_scalar_copy_producer_shares_one_input_allowance() {
    let data = serde_json::json!({"big": "x".repeat(10_000), "unused": "z".repeat(1024 * 1024)});
    let mut operation =
        rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
            .unwrap();
    let expression = parse_expression("big").unwrap();
    let mut completed = 0;
    for _ in 0..1000 {
        match expression_text(&expression, &data, &[], &mut operation) {
            Ok(text) => {
                assert_eq!(text.len(), 10_000);
                completed += 1;
            }
            Err(error) => {
                assert!(
                    error
                        .to_string()
                        .contains("template text copy exceeds selected input work")
                );
                assert!(completed > 0 && completed < 1000);
                return;
            }
        }
    }
    panic!("repeated real scalar copies must consume one shared input allowance");
}

#[test]
fn template_repeated_json_copies_refuse_atomically_without_unused_sibling_credit() {
    let data =
        serde_json::json!({"big": "x".repeat(10_000), "unused": "z".repeat(16 * 1024 * 1024)});
    let mut document = Document::new();
    document.add_paragraph(&"{{ big }}".repeat(200));
    let before = document.to_bytes().unwrap();
    let error = document
        .render_template_versioned(&data, 1)
        .unwrap_err()
        .to_string();
    assert!(
        error.contains("template text copy exceeds selected input work"),
        "{error}"
    );
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn template_empty_json_operands_join_in_the_real_native_pipeline() {
    for input in [Value::Null, Value::String(String::new())] {
        let data = serde_json::json!({"values": vec![input; 1000]});
        let mut document = Document::new();
        document.add_paragraph(r#"{{ values | join(",") }}"#);
        document.render_template_versioned(&data, 1).unwrap();
        assert_eq!(document.paragraphs()[0].text(), ",".repeat(999));
    }
}

#[test]
fn template_ingress_preserves_recursive_model_refusal_without_publication() {
    let mut table = CT_Tbl::new();
    for _ in 1..33 {
        let mut outer = CT_Tbl::new();
        let mut row = CT_Row::new();
        let mut cell = CT_Tc::new();
        cell.content.push(CellContent::Table(table));
        row.cells.push(cell);
        outer.rows.push(row);
        table = outer;
    }
    let mut document = Document::new();
    document
        .document
        .body
        .content
        .push(rdocx_oxml::document::BodyContent::Table(Box::new(table)));
    let expected = document
        .validate_model_resource_limits()
        .unwrap_err()
        .to_string();
    let original_children = document.document.body.content.len();
    let original_parts = document.package.parts.clone();
    let actual = document
        .render_template_versioned(&serde_json::json!({}), 1)
        .unwrap_err()
        .to_string();
    assert_eq!(actual, expected);
    assert_eq!(document.document.body.content.len(), original_children);
    assert_eq!(document.package.parts, original_parts);
}

#[test]
fn versioned_template_filters_keep_nested_borrowed_inputs_and_unicode_output() {
    let data = serde_json::json!({
        "title": "root", "greek": "ΟΣ", "turkish": "İ",
        "rows": [
            {"name": "  Straße  ", "tags": [1, false, null, "é"], "items": [{"label": "one"}]},
            {"name": " second ", "tags": ["", null], "items": []},
        ],
    });
    let original = data.clone();
    let mut document = Document::new();
    for text in [
        "{% for row in rows %}",
        r#"{{ row.name | upper | trim | trim }}:{{ row.tags | join("/") }}"#,
        "{% for child in row.items %}",
        "{{ child.label }}",
        "{% endfor %}",
        "{% endfor %}",
        "{{ title }}",
        "{{ greek | lower }}/{{ turkish | lower }}",
    ] {
        document.add_paragraph(text);
    }
    document.render_template_versioned(&data, 1).unwrap();
    let text = document
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text())
        .collect::<Vec<_>>();
    assert_eq!(
        text,
        [
            "STRASSE:1/false//é",
            "one",
            "SECOND:/",
            "root",
            "ος/i\u{307}"
        ]
    );
    assert_eq!(data, original);
    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        text
    );
}

#[test]
fn loop_scope_lookup_retains_actual_input_membership_and_deferred_fallback() {
    let data = serde_json::json!({
        "title": "root", "rows": [{"name": "first", "items": [{"label": "nested"}]}],
    });
    let mut operation =
        rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
            .unwrap();
    let rows = operation
        .template_datum_key(&data, "rows")
        .unwrap()
        .unwrap();
    let row = operation.template_datum_index(rows, 0).unwrap().unwrap();
    let scopes = [Scope {
        name: "row".into(),
        value: Some(row),
        deferred: false,
    }];
    for path in ["row.name", "name"] {
        let expression = parse_expression(path).unwrap();
        let selected = expression_value(&expression, &data, &scopes, &mut operation)
            .unwrap()
            .unwrap();
        assert!(std::ptr::eq(selected, &row["name"]));
        let legacy_path = path.split('.').map(str::to_owned).collect::<Vec<_>>();
        let ResolvedValue::Value(selected) =
            resolve_value(&data, &scopes, &legacy_path, &mut operation).unwrap()
        else {
            panic!("real borrowed scope must not be deferred")
        };
        assert!(std::ptr::eq(selected, &row["name"]));
    }
    let deferred = [Scope {
        name: "item".into(),
        value: None,
        deferred: true,
    }];
    assert!(deferred_expression(
        &parse_expression("item.label").unwrap(),
        &data,
        &deferred
    ));
    assert!(matches!(
        resolve_value(
            &data,
            &deferred,
            &["item".into(), "label".into()],
            &mut operation
        )
        .unwrap(),
        ResolvedValue::Deferred
    ));
    let ResolvedValue::Value(root) =
        resolve_value(&data, &deferred, &["title".into()], &mut operation).unwrap()
    else {
        panic!("known root input remains selected through deferred validation")
    };
    assert!(std::ptr::eq(root, &data["title"]));
}

#[derive(Debug, Clone)]
struct Scope<'data> {
    name: String,
    value: Option<&'data Value>,
    deferred: bool,
}

enum ResolvedValue<'a> {
    Value(&'a Value),
    Deferred,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum EndKind {
    For,
    If,
}

impl EndKind {
    fn name(self) -> &'static str {
        match self {
            Self::For => "endfor",
            Self::If => "endif",
        }
    }
}

#[derive(Debug, Clone)]
enum Control {
    For { name: String, path: Vec<String> },
    If { path: Vec<String> },
    End(EndKind),
}

#[derive(Debug)]
enum Block<T> {
    Item {
        value: T,
    },
    For {
        name: String,
        path: Vec<String>,
        source_range: std::ops::Range<usize>,
        body: Vec<Block<T>>,
    },
    If {
        path: Vec<String>,
        body: Vec<Block<T>>,
    },
}

fn marker_from_sources(sources: &[String]) -> Result<Option<Control>> {
    let with_controls = sources
        .iter()
        .filter(|source| source.contains("{%") || source.contains("%}"))
        .collect::<Vec<_>>();
    if with_controls.is_empty() {
        return Ok(None);
    }
    let nonempty = sources
        .iter()
        .filter(|source| !source.trim().is_empty())
        .collect::<Vec<_>>();
    if nonempty.len() != 1 || with_controls.len() != 1 {
        return Err(template_error(
            "a control marker must occupy its own paragraph or table row",
        ));
    }
    parse_control(nonempty[0]).map(Some)
}

fn parse_control(source: &str) -> Result<Control> {
    let trimmed = source.trim();
    if !trimmed.starts_with("{%") || !trimmed.ends_with("%}") {
        return Err(template_error(
            "a control marker must occupy its own paragraph or table row",
        ));
    }
    let inner = trimmed[2..trimmed.len() - 2].trim();
    if inner.is_empty()
        || inner.contains("{{")
        || inner.contains("{%")
        || inner.contains("}}")
        || inner.contains("%}")
    {
        return Err(template_error("malformed control tag"));
    }
    let words = inner.split_whitespace().collect::<Vec<_>>();
    match words.as_slice() {
        ["for", name, "in", path] => {
            validate_scope_name(name)?;
            Ok(Control::For {
                name: (*name).to_owned(),
                path: parse_path(path)?.into_iter().map(str::to_owned).collect(),
            })
        }
        ["if", path] => Ok(Control::If {
            path: parse_path(path)?.into_iter().map(str::to_owned).collect(),
        }),
        ["endfor"] => Ok(Control::End(EndKind::For)),
        ["endif"] => Ok(Control::End(EndKind::If)),
        _ => Err(template_error("unsupported or malformed control tag")),
    }
}

fn validate_scope_name(name: &str) -> Result<()> {
    if name.is_empty()
        || name.contains('.')
        || name.chars().any(|character| {
            character.is_whitespace() || matches!(character, '{' | '}' | '%' | '\0')
        })
    {
        return Err(template_error("invalid loop variable"));
    }
    Ok(())
}

fn parse_blocks<T: Clone>(
    items: &[T],
    marker: impl Fn(&T) -> Result<Option<Control>>,
) -> Result<Vec<Block<T>>> {
    let (blocks, next) = parse_block_level(items, 0, None, &marker)?;
    if next != items.len() {
        return Err(template_error("unexpected control marker"));
    }
    Ok(blocks)
}

fn parse_block_level<T: Clone, F: Fn(&T) -> Result<Option<Control>>>(
    items: &[T],
    mut index: usize,
    expected: Option<EndKind>,
    marker: &F,
) -> Result<(Vec<Block<T>>, usize)> {
    let mut blocks = Vec::new();
    while index < items.len() {
        match marker(&items[index])? {
            None => {
                blocks.push(Block::Item {
                    value: items[index].clone(),
                });
                index += 1;
            }
            Some(Control::For { name, path }) => {
                let (body, next) = parse_block_level(items, index + 1, Some(EndKind::For), marker)?;
                blocks.push(Block::For {
                    name,
                    path,
                    source_range: index + 1..next - 1,
                    body,
                });
                index = next;
            }
            Some(Control::If { path }) => {
                let (body, next) = parse_block_level(items, index + 1, Some(EndKind::If), marker)?;
                blocks.push(Block::If { path, body });
                index = next;
            }
            Some(Control::End(actual)) => {
                let Some(expected) = expected else {
                    return Err(template_error(&format!(
                        "unexpected {} marker",
                        actual.name()
                    )));
                };
                if actual != expected {
                    return Err(template_error(&format!(
                        "expected {} before {}",
                        expected.name(),
                        actual.name()
                    )));
                }
                return Ok((blocks, index + 1));
            }
        }
    }
    if let Some(expected) = expected {
        return Err(template_error(&format!(
            "missing {} marker",
            expected.name()
        )));
    }
    Ok((blocks, index))
}

fn is_truthy(value: &Value) -> bool {
    match value {
        Value::Null => false,
        Value::Bool(value) => *value,
        Value::Number(value) => value.as_f64().is_some_and(|number| number != 0.0),
        Value::String(value) => !value.is_empty(),
        Value::Array(value) => !value.is_empty(),
        Value::Object(value) => !value.is_empty(),
    }
}

fn resolve_value<'a>(
    root: &'a Value,
    scopes: &[Scope<'a>],
    path: &[String],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<ResolvedValue<'a>> {
    let display_path = path.join(".");
    for scope in scopes.iter().rev() {
        if path
            .first()
            .is_some_and(|component| component == &scope.name)
        {
            if scope.deferred {
                return Ok(ResolvedValue::Deferred);
            }
            return traverse_path(
                scope.value.expect("selected non-deferred scope"),
                &path[1..],
                &display_path,
                operation,
            )
            .map(ResolvedValue::Value);
        }
    }
    let mut saw_deferred = false;
    for scope in scopes.iter().rev() {
        if scope.deferred {
            saw_deferred = true;
            continue;
        }
        if let Some(value) = try_traverse_path(
            scope.value.expect("selected non-deferred scope"),
            path,
            operation,
        )? {
            return Ok(ResolvedValue::Value(value));
        }
    }
    if let Some(value) = try_traverse_path(root, path, operation)? {
        return Ok(ResolvedValue::Value(value));
    }
    if saw_deferred {
        for scope in scopes.iter().rev().filter(|scope| !scope.deferred) {
            if path.first().is_some_and(|component| {
                scope
                    .value
                    .and_then(Value::as_object)
                    .is_some_and(|object| object.contains_key(component))
            }) {
                return traverse_path(
                    scope.value.expect("selected non-deferred scope"),
                    path,
                    &display_path,
                    operation,
                )
                .map(ResolvedValue::Value);
            }
        }
        if path.first().is_some_and(|component| {
            root.as_object()
                .is_some_and(|object| object.contains_key(component))
        }) {
            return traverse_path(root, path, &display_path, operation).map(ResolvedValue::Value);
        }
        return Ok(ResolvedValue::Deferred);
    }
    traverse_path(root, path, &display_path, operation).map(ResolvedValue::Value)
}

fn try_traverse_path<'a>(
    mut value: &'a Value,
    path: &[String],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<Option<&'a Value>> {
    for component in path {
        if !value.is_object() {
            return Ok(None);
        }
        let Some(next) = operation.template_datum_key(value, component)? else {
            return Ok(None);
        };
        value = next;
    }
    Ok(Some(value))
}

fn traverse_path<'a>(
    mut value: &'a Value,
    path: &[String],
    display_path: &str,
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<&'a Value> {
    for component in path {
        let Value::Object(_) = value else {
            return Err(template_error(&format!(
                "path `{display_path}` crosses a non-object value"
            )));
        };
        let Some(next) = operation.template_datum_key(value, component)? else {
            return Err(template_error(&format!("missing path `{display_path}`")));
        };
        value = next;
    }
    Ok(value)
}

fn require_scalar(value: &Value, path: &[String]) -> Result<()> {
    if matches!(value, Value::Array(_) | Value::Object(_)) {
        return Err(template_error(&format!(
            "path `{}` does not resolve to a scalar value",
            path.join("."),
        )));
    }
    Ok(())
}

fn scalar_text<'a>(
    value: &'a Value,
    path: &[String],
    operation: &mut rdocx_oxml::document::ScopedWordOperation<'a>,
) -> Result<String> {
    require_scalar(value, path)?;
    let bytes = operation.template_scalar_input_bytes(value)?;
    operation.reserve_template_text_copy(bytes)?;
    match value {
        Value::String(value) => Ok(value.clone()),
        Value::Number(value) => {
            let mut text = String::with_capacity(bytes);
            std::fmt::write(&mut text, format_args!("{value}"))
                .map_err(|_| template_error("canonical number formatting failed"))?;
            Ok(text)
        }
        Value::Bool(value) => Ok(if *value { "true" } else { "false" }.to_owned()),
        Value::Null => Ok(String::new()),
        Value::Array(_) | Value::Object(_) => {
            unreachable!("scalar kind was checked before admission")
        }
    }
}

fn parse_path(inner: &str) -> Result<Vec<&str>> {
    let path = inner.trim();
    if path.is_empty() {
        return Err(template_error("empty scalar path"));
    }
    let components = path.split('.').collect::<Vec<_>>();
    if components.iter().any(|component| {
        component.is_empty()
            || component.chars().any(|character| {
                character.is_whitespace() || matches!(character, '{' | '}' | '%' | '\0')
            })
    }) {
        return Err(template_error("invalid dotted scalar path"));
    }
    Ok(components)
}

fn template_error(message: &str) -> Error {
    Error::Other(format!("invalid template: {message}"))
}

fn collect_table(table: &CT_Tbl, sources: &mut Vec<String>) {
    for (_, _, control) in &table.content_controls {
        collect_control(control, sources);
    }
    for row in &table.rows {
        collect_row(row, sources);
    }
}

fn collect_control(control: &CT_Sdt, sources: &mut Vec<String>) {
    for content in &control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                sources.extend(placeholder::replaceable_texts(paragraph));
            }
            SdtContent::Table(table) => collect_table(table, sources),
            SdtContent::Row(row) => collect_row(row, sources),
            SdtContent::Cell(cell) => collect_cell(cell, sources),
            SdtContent::ContentControl(nested) => collect_control(nested, sources),
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
}

fn collect_row(row: &CT_Row, sources: &mut Vec<String>) {
    for (_, _, control) in &row.content_controls {
        collect_control(control, sources);
    }
    for cell in &row.cells {
        collect_cell(cell, sources);
    }
}

fn collect_row_marker_sources(row: &CT_Row, sources: &mut Vec<String>) {
    for (_, _, control) in &row.content_controls {
        collect_control_marker_sources(control, sources);
    }
    for cell in &row.cells {
        collect_cell_marker_sources(cell, sources);
    }
}

fn collect_cell_marker_sources(cell: &CT_Tc, sources: &mut Vec<String>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                sources.extend(placeholder::replaceable_texts(paragraph));
            }
            CellContent::Table(_) => {}
            CellContent::ContentControl(control) => {
                collect_control_marker_sources(control, sources);
            }
        }
    }
}

fn collect_control_marker_sources(control: &CT_Sdt, sources: &mut Vec<String>) {
    for content in &control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                sources.extend(placeholder::replaceable_texts(paragraph));
            }
            SdtContent::Table(_) => {}
            SdtContent::Row(row) => collect_row_marker_sources(row, sources),
            SdtContent::Cell(cell) => collect_cell_marker_sources(cell, sources),
            SdtContent::ContentControl(nested) => {
                collect_control_marker_sources(nested, sources);
            }
            SdtContent::Run(run) => sources.push(
                run.content
                    .iter()
                    .filter_map(|content| match content {
                        RunContent::Text(text) => Some(text.text.as_str()),
                        _ => None,
                    })
                    .collect(),
            ),
            SdtContent::RawXml(_) => {}
        }
    }
}

fn collect_cell(cell: &CT_Tc, sources: &mut Vec<String>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                sources.extend(placeholder::replaceable_texts(paragraph));
            }
            CellContent::Table(table) => collect_table(table, sources),
            CellContent::ContentControl(control) => collect_control(control, sources),
        }
    }
}

#[cfg(test)]
pub(crate) fn chart_sources(xml: &[u8]) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut sources = Vec::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(element))
                if matches_local_name(element.name().as_ref(), b"t")
                    || matches_local_name(element.name().as_ref(), b"v") =>
            {
                sources.push(oxml_core::xml_text::read_element_text(
                    &mut reader,
                    element.name(),
                ));
            }
            Ok(_) => {}
            Err(error) => return Err(template_xml_error(error)),
        }
        buffer.clear();
    }

    Ok(sources)
}

fn template_xml_error(error: quick_xml::Error) -> Error {
    Error::Other(format!("invalid template XML: {error}"))
}

#[cfg(test)]
#[test]
fn versioned_template_delimiters_are_quoted_and_escaped_once() {
    let source = r#"A{{ ["x}}|y"] | default("{%literal%}") }}B {{{{name}}}} C {%{% if q %}%}"#;
    let tokens = scalar_tokens(source, true).unwrap();
    let data = serde_json::json!({"x}}|y":"{{inserted}}\u{e000}rdocx-template-0\u{e001}"});
    let mut output = source.to_owned();
    for token in tokens.into_iter().rev() {
        let value = match token.value {
            ScalarValue::Expression(expression) => expression_text(
                &expression,
                &data,
                &[],
                &mut rdocx_oxml::document::ScopedWordOperation::with_template_inputs(
                    vec![],
                    vec![],
                    &data,
                )
                .unwrap(),
            )
            .unwrap(),
            ScalarValue::Literal(text) => text.to_owned(),
        };
        output.replace_range(token.range, &value);
    }
    assert_eq!(
        output,
        "A{{inserted}}\u{e000}rdocx-template-0\u{e001}B {{name}} C {% if q %}"
    );
    for source in [
        "{{ missing",
        "{{ x | default(\"unterminated) }}",
        "{{ x {% }}",
        "bad }}",
    ] {
        assert!(scalar_tokens(source, true).is_err(), "{source}");
    }
}

#[test]
fn scoped_row_consumer_matches_actual_native_direct_and_all_visitors() {
    for body in [
        "<w:tc><w:p><w:r><w:t>{% if flag %}</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc>",
        "<w:tc><w:sdt><w:sdtContent><w:r><w:t>direct</w:t><w:t>only</w:t></w:r><w:p><w:r><w:t>both</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc><w:tc/>",
        "<w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>control</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt><w:tc><w:p/></w:tc>",
    ] {
        let xml = format!(
            r#"<w:tr xmlns:w="{}">{body}</w:tr>"#,
            rdocx_oxml::namespace::W_NS
        );
        let wrapped = format!(
            r#"<w:tbl xmlns:w="{}">{xml}</w:tbl>"#,
            rdocx_oxml::namespace::W_NS
        );
        let table = CT_Tbl::from_xml_fragment_checked(wrapped.as_bytes()).unwrap();
        let row = &table.rows[0];
        let mut old_direct = Vec::new();
        collect_row_marker_sources(row, &mut old_direct);
        let mut old_all = Vec::new();
        collect_row(row, &mut old_all);
        let mut operation = rdocx_oxml::document::ScopedWordOperation::new(xml.as_bytes());
        let source = operation.source(0..xml.len(), &Default::default()).unwrap();
        let (buffers, direct, all, hazard) =
            operation.template_row_replaceable_texts(&source).unwrap();
        let strings = |view: &[usize]| {
            view.iter()
                .map(|index| buffers[*index].as_str())
                .collect::<Vec<_>>()
        };
        assert_eq!(strings(&direct), old_direct);
        assert_eq!(strings(&all), old_all);
        assert_eq!(
            program_marker(&old_direct, true).unwrap().is_some(),
            program_marker(
                &direct
                    .iter()
                    .map(|index| buffers[*index].clone())
                    .collect::<Vec<_>>(),
                true
            )
            .unwrap()
            .is_some()
        );
        assert_eq!(
            hazard,
            !row.content_controls.is_empty()
                || row
                    .extra_xml
                    .iter()
                    .any(|(position, raw)| !CT_Row::raw_is_root_attributes(*position, raw))
        );
    }
}

#[cfg(test)]
#[test]
fn scoped_template_clone_lineage_keeps_physical_body_and_field_slots() {
    use rdocx_oxml::document::BodyContent;
    use rdocx_oxml::text::{CT_R, Field};
    let mut paragraph = CT_P::new();
    let mut run = CT_R::new("same");
    run.content
        .push(RunContent::Field(Field::new("CUSTOM", "same")));
    run.content
        .push(RunContent::Field(Field::new("CUSTOM", "same")));
    paragraph.runs.push(run);
    let mut document = Document::new();
    let mut table = CT_Tbl::new();
    let mut row = CT_Row::new();
    let mut cell = CT_Tc::new();
    cell.content = vec![CellContent::Paragraph(paragraph.clone())];
    row.cells.push(cell);
    table.rows.push(row);
    document.document.body.content = vec![
        BodyContent::Table(Box::new(table)),
        BodyContent::RawXml(Box::new(b"<w:opaque/>".to_vec())),
        BodyContent::Paragraph(Box::new(paragraph)),
    ];
    let (candidate, mut lineage) = TemplateWordLineage::clone_for_staging(&document);
    assert_eq!(
        lineage.preparation[0].as_ref().unwrap().0,
        vec![
            TemplateWordStep::Main,
            TemplateWordStep::BodyItem(0),
            TemplateWordStep::TableRow(0),
            TemplateWordStep::RowCell(0),
            TemplateWordStep::CellItem(0)
        ]
    );
    assert_eq!(
        lineage.preparation[1].as_ref().unwrap().0,
        vec![TemplateWordStep::Main, TemplateWordStep::BodyItem(2)]
    );
    let BodyContent::Paragraph(original) = &document.document.body.content[2] else {
        unreachable!()
    };
    let BodyContent::Paragraph(current) = &candidate.document.body.content[2] else {
        unreachable!()
    };
    let path = current.accepted_run_paths().remove(0);
    let nested = lineage.initial_field(0, &path, 1).unwrap();
    let direct = lineage.initial_field(1, &path, 1).unwrap();
    assert!(
        !std::ptr::eq(nested, direct),
        "equal nested content keeps its own actual frozen owner"
    );
    let RunContent::Field(first) = &original.runs[0].content[1] else {
        unreachable!()
    };
    let RunContent::Field(second) = &original.runs[0].content[2] else {
        unreachable!()
    };
    assert!(std::ptr::eq(
        lineage.initial_field(1, &path, 1).unwrap(),
        first
    ));
    assert!(std::ptr::eq(
        lineage.initial_field(1, &path, 2).unwrap(),
        second
    ));
    assert!(!std::ptr::eq(first, second));
    lineage.installed_field(1, &path, 1);
    assert!(lineage.initial_field(1, &path, 1).is_none());
    assert!(std::ptr::eq(
        lineage.initial_field(1, &path, 2).unwrap(),
        second
    ));
    lineage.finish_preparation();
    assert!(lineage.initial_field(1, &path, 2).is_none());
}

#[cfg(test)]
#[test]
fn scoped_template_initial_preparation_uses_the_existing_operation_refusal() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    for kind in 0..3 {
        let xml = if kind == 1 {
            format!(
                r#"<w:p xmlns:w="{}"><w:sdt><w:sdtContent><w:r><w:t>outside</w:t></w:r><w:fldSimple w:instr="CUSTOM"><w:r><w:t>cache</w:t></w:r></w:fldSimple></w:sdtContent></w:sdt></w:p>"#,
                rdocx_oxml::namespace::W_NS
            )
        } else {
            format!(
                r#"<w:p xmlns:w="{}"><w:r><w:t>outside</w:t></w:r><w:fldSimple w:instr="CUSTOM"><w:r><w:t>cache</w:t></w:r></w:fldSimple></w:p>"#,
                rdocx_oxml::namespace::W_NS
            )
        };
        let paragraph = CT_P::from_xml_fragment(xml.as_bytes()).unwrap();
        let item = if kind == 2 {
            let mut table = CT_Tbl::new();
            let mut row = CT_Row::new();
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            row.cells.push(cell);
            table.rows.push(row);
            BodyContent::Table(Box::new(table))
        } else {
            BodyContent::Paragraph(Box::new(paragraph))
        };
        let mut document = Document::new();
        document.document.body.content = vec![item];
        let data = Value::Null;
        let mut operation = ScopedWordOperation::with_template_word_inputs(
            Vec::new(),
            Vec::new(),
            vec![&document.document.body],
            Vec::new(),
            Vec::new(),
            Vec::new(),
            &data,
        )
        .unwrap();
        let mut fields = Vec::new();
        crate::document::visit_body_paragraphs(
            &document.document.body.content,
            false,
            &mut |paragraph| {
                for path in paragraph.accepted_run_paths() {
                    for content in &paragraph.accepted_run(&path).unwrap().content {
                        if let RunContent::Field(field) = content {
                            fields.push((field, field));
                        }
                    }
                }
            },
        );
        assert_eq!(fields.len(), 1, "one physical Field in fixture {kind}");
        let source = operation
            .source_model_document(&document.document, &[], &fields, &[])
            .unwrap();
        // This actual reader establishes selected current Field coverage for
        // the retained parsed-writer fixture before testing copy exhaustion.
        // Unknown writer coverage alone cannot be used as a debit witness.
        operation.source_document(&source).unwrap();
        let mut refused = false;
        for _ in 0..10000 {
            if operation.repeat_source(&source).is_err() {
                refused = true;
                break;
            }
        }
        assert!(
            refused,
            "the actual known Field copy ledger must be exhausted for owner {kind}"
        );
        let (mut candidate, mut lineage) = TemplateWordLineage::clone_for_staging(&document);
        assert!(
            candidate
                .prepare_staged_template_package(
                    &mut operation,
                    &mut lineage,
                    &mut crate::document::TemplateMainSources::default()
                )
                .is_err(),
            "actual owner {kind}"
        );
        assert!(lineage.prepared);
        let mut ordinary = document.clone_for_staging();
        ordinary.prepare_staged_package().unwrap();
    }
}

#[cfg(test)]
#[test]
fn scoped_template_initial_preparation_preserves_ordinary_model_and_package_results() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    use rdocx_oxml::text::{CT_R, Field};
    let mut paragraph = CT_P::new();
    let mut run = CT_R::new("ordinary sibling");
    run.content
        .push(RunContent::Field(Field::new("CUSTOM", "cache")));
    paragraph.runs.push(run);
    let mut document = Document::new();
    document.document.body.content = vec![BodyContent::Paragraph(Box::new(paragraph))];
    let data = Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&document.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut candidate, mut lineage) = TemplateWordLineage::clone_for_staging(&document);
    let mut ordinary = document.clone_for_staging();
    let expected = ordinary.prepare_staged_package().unwrap();
    let actual = candidate
        .prepare_staged_template_package(
            &mut operation,
            &mut lineage,
            &mut crate::document::TemplateMainSources::default(),
        )
        .unwrap();
    assert_eq!(actual, expected);
    assert_eq!(candidate.document, ordinary.document);
    assert_eq!(candidate.package.parts, ordinary.package.parts);
}

#[cfg(test)]
#[test]
fn scoped_template_initial_writer_carries_actual_main_bytes_and_preserves_legacy_programs() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    use rdocx_oxml::text::{CT_R, Field};
    let mut paragraph = CT_P::new();
    let mut run = CT_R::new("ordinary sibling");
    run.content
        .push(RunContent::Field(Field::new("CUSTOM", "cache")));
    paragraph.runs.push(run);
    let mut document = Document::new();
    document.document.body.content = vec![BodyContent::Paragraph(Box::new(paragraph))];
    let data = Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&document.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut candidate, lineage) = TemplateWordLineage::clone_for_staging(&document);
    let mut ordinary = document.clone_for_staging();
    let expected = ordinary.template_programs().unwrap();
    let mut main_sources = crate::document::TemplateMainSources::default();
    let actual = candidate
        .template_programs_scoped(&mut operation, &lineage, &mut main_sources)
        .unwrap();
    let source = main_sources.source.take();
    let source = source.expect("the actual fresh main writer returns its source");
    assert_eq!(
        source.as_bytes(),
        candidate
            .package
            .get_part(&candidate.doc_part_name)
            .unwrap()
    );
    assert_eq!(candidate.package.parts, ordinary.package.parts);
    assert_eq!(actual.len(), expected.len());
    for (actual, expected) in actual.iter().zip(&expected) {
        assert_eq!(actual.raw, expected.raw);
        assert_eq!(actual.ranges, expected.ranges);
        assert_eq!(actual.items.len(), expected.items.len());
    }
    // The same carrier remains anchored, without renewed ingress input.
    let copied = operation.repeat_source(&source).unwrap();
    assert_eq!(copied.as_bytes(), source.as_bytes());
}

#[cfg(test)]
#[test]
fn scoped_template_initial_writer_uses_the_actual_exhausted_operation() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    use rdocx_oxml::text::{CT_R, Field};
    let mut paragraph = CT_P::new();
    let mut run = CT_R::new("ordinary");
    run.content
        .push(RunContent::Field(Field::new("CUSTOM", "cache")));
    paragraph.runs.push(run);
    let mut document = Document::new();
    document.document.body.content = vec![BodyContent::Paragraph(Box::new(paragraph))];
    let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
        unreachable!()
    };
    let RunContent::Field(field) = &paragraph.runs[0].content[1] else {
        unreachable!()
    };
    let data = Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&document.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let source = operation
        .source_model_document(&document.document, &[], &[(field, field)], &[])
        .unwrap();
    operation.source_document(&source).unwrap();
    assert!((0..10000).any(|_| operation.repeat_source(&source).is_err()));
    let (mut candidate, lineage) = TemplateWordLineage::clone_for_staging(&document);
    let before = candidate.package.parts.clone();
    let mut main_sources = crate::document::TemplateMainSources::default();
    assert!(
        candidate
            .template_programs_scoped(&mut operation, &lineage, &mut main_sources)
            .is_err()
    );
    assert_eq!(
        candidate.package.parts, before,
        "failed main production publishes no part"
    );
    let mut ordinary = document.clone_for_staging();
    ordinary.template_programs().unwrap();
}

#[cfg(test)]
#[test]
fn scoped_template_initial_writer_preserves_opaque_and_inline_control_bytes() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    let xml = format!(
        r#"<w:p xmlns:w="{}"><w:sdt><w:sdtPr><w:tag w:val="kept"/></w:sdtPr><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt><x:opaque xmlns:x="urn:producer" x:keep="yes"/></w:p>"#,
        rdocx_oxml::namespace::W_NS
    );
    let mut document = Document::new();
    document.document.body.content = vec![
        BodyContent::RawXml(Box::new(
            br#"<x:unmodeled xmlns:x="urn:producer">opaque</x:unmodeled>"#.to_vec(),
        )),
        BodyContent::Paragraph(Box::new(CT_P::from_xml_fragment(xml.as_bytes()).unwrap())),
    ];
    let data = Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&document.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut candidate, lineage) = TemplateWordLineage::clone_for_staging(&document);
    let mut ordinary = document.clone_for_staging();
    let expected = ordinary.template_programs().unwrap();
    let mut main_sources = crate::document::TemplateMainSources::default();
    let actual = candidate
        .template_programs_scoped(&mut operation, &lineage, &mut main_sources)
        .unwrap();
    assert_eq!(candidate.package.parts, ordinary.package.parts);
    assert_eq!(actual.len(), expected.len());
    for (actual, expected) in actual.iter().zip(expected) {
        assert_eq!(actual.raw, expected.raw);
    }
}

#[cfg(test)]
#[test]
fn scoped_template_initial_writer_retains_actual_checked_w14_insertion() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    let paragraph = CT_P::from_xml_fragment(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="0000AAAA"><w:r><w:t>plain</w:t></w:r></w:p>"#,
    )
    .unwrap();
    let mut document = Document::new();
    document.document.body.content = vec![BodyContent::Paragraph(Box::new(paragraph))];
    let held_model = document.document.to_xml().unwrap();
    let held_package = document.package.parts.clone();
    let data = Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&document.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut candidate, lineage) = TemplateWordLineage::clone_for_staging(&document);
    let mut ordinary = document.clone_for_staging();
    let expected = ordinary.template_programs().unwrap();
    let mut main_sources = crate::document::TemplateMainSources::default();
    let actual = candidate
        .template_programs_scoped(&mut operation, &lineage, &mut main_sources)
        .unwrap();
    let source = main_sources.source.take();
    let source = source.expect("actual checked w14 insertion returns its producer Source");
    assert_eq!(
        source.as_bytes(),
        candidate
            .package
            .get_part(&candidate.doc_part_name)
            .unwrap()
    );
    assert_eq!(document.document.to_xml().unwrap(), held_model);
    assert_eq!(document.package.parts, held_package);
    assert_eq!(candidate.package.parts, ordinary.package.parts);
    assert_eq!(actual.len(), expected.len());
    for (actual, expected) in actual.iter().zip(&expected) {
        assert_eq!(actual.raw, expected.raw);
        assert_eq!(actual.ranges, expected.ranges);
    }
}

#[cfg(test)]
#[test]
fn scoped_template_prepared_main_reaches_actual_grammar_parser_and_sticky_guard() {
    use rdocx_oxml::document::{BodyContent, ScopedWordOperation};
    use rdocx_oxml::text::{CT_R, Field};
    let mut paragraph = CT_P::new();
    paragraph.runs.push(CT_R::new("Hello {{ name }}"));
    let mut run = CT_R::new("");
    run.content = vec![RunContent::Field(Field::new("CUSTOM", "cache"))];
    paragraph.runs.push(run);
    let mut document = Document::new();
    document.document.body.content = vec![BodyContent::Paragraph(Box::new(paragraph))];
    let definition = document
        .add_numbering_definition(&[crate::document::ListLevel::decimal()])
        .unwrap();
    document.add_numbering_instance(definition, &[]).unwrap();
    let selected = document.add_numbering_instance(definition, &[]).unwrap();
    document
        .paragraph_mut(0)
        .unwrap()
        .set_numbering(selected, 0);
    let data = Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&document.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut candidate, mut lineage) = TemplateWordLineage::clone_for_staging(&document);
    let mut main = crate::document::TemplateMainSources::default();
    candidate
        .template_programs_scoped(&mut operation, &lineage, &mut main)
        .unwrap();
    candidate
        .prepare_staged_template_package(&mut operation, &mut lineage, &mut main)
        .unwrap();
    let programs = candidate
        .template_programs_scoped(&mut operation, &lineage, &mut main)
        .unwrap();
    let source_handle = std::sync::Arc::clone(
        main.source
            .as_ref()
            .expect("prepared writer has its actual Source"),
    );
    let source = source_handle.as_ref();
    assert_eq!(source.as_bytes(), candidate.document.to_xml().unwrap());
    assert_eq!(
        source.as_bytes(),
        candidate
            .package
            .get_part(&candidate.doc_part_name)
            .unwrap(),
        "private staging package and actual Source use the same live IDs"
    );

    let mut scoped = Some((&mut operation, source));
    let parsed = parse_program(&programs[0], false, false, &mut scoped).unwrap();
    assert!(!parsed.scalars.is_empty());
    validate_program_numbering(
        &parsed,
        &parsed.blocks,
        &candidate,
        true,
        &mut Some((&mut operation, source)),
    )
    .unwrap();
    let mut exhausted = false;
    for _ in 0..10000 {
        if operation.repeat_source(source).is_err() {
            exhausted = true;
            break;
        }
    }
    assert!(exhausted);
    assert!(
        validate_program_numbering(
            &parsed,
            &parsed.blocks,
            &candidate,
            true,
            &mut Some((&mut operation, source))
        )
        .is_err()
    );
    validate_program_numbering(&parsed, &parsed.blocks, &candidate, true, &mut None).unwrap();
    assert!(
        candidate
            .template_note_identities(Some(&mut crate::document::TemplateMainWriter {
                operation: &mut operation,
                lineage: &lineage,
                sources: &mut main,
            }))
            .is_err()
    );
    assert!(candidate.template_note_identities(None).unwrap().is_empty());
    assert!(
        parse_program(
            &programs[0],
            false,
            false,
            &mut Some((&mut operation, source))
        )
        .is_err()
    );
    assert!(
        parse_program(&programs[0], false, false, &mut None).is_ok(),
        "ordinary parser remains unchanged"
    );
}

#[cfg(test)]
#[test]
fn scoped_template_original_related_note_reaches_actual_refresh_and_grammar() {
    use rdocx_oxml::document::ScopedWordOperation;
    let mut source = Document::new();
    source.add_paragraph("body");
    source.add_footnote("A&B");
    let bytes = source.to_bytes().unwrap();
    let original = Document::from_bytes(&bytes).unwrap();
    let part = original.footnotes_part_name.clone().unwrap();
    let data = serde_json::Value::Null;
    let mut operation = ScopedWordOperation::with_template_word_inputs(
        original
            .package
            .parts
            .iter()
            .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
            .collect(),
        original
            .footnotes
            .footnotes
            .iter()
            .flat_map(|note| note.paragraphs.iter())
            .collect(),
        vec![&original.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut current, mut lineage) = TemplateWordLineage::clone_for_staging(&original);
    let mut sources = crate::document::TemplateMainSources::default();
    current
        .template_programs_scoped(&mut operation, &lineage, &mut sources)
        .unwrap();
    current
        .prepare_staged_template_package(&mut operation, &mut lineage, &mut sources)
        .unwrap();
    let programs = current
        .template_programs_scoped(&mut operation, &lineage, &mut sources)
        .unwrap();
    let source = sources
        .related_source(&part, current.package.get_part_shared(&part).unwrap())
        .unwrap();
    let emitted = operation
        .source_model_notes_from_read(&current.footnotes, false, source)
        .unwrap();
    assert_eq!(
        emitted.as_bytes(),
        current.footnotes.to_xml_footnotes().unwrap()
    );
    let program = programs
        .iter()
        .find(|program| program.part == part)
        .unwrap();
    let mut scoped = Some((&mut operation, source.as_ref()));
    parse_program(program, false, false, &mut scoped).unwrap();
    operation.reserve_template_text_copy(4).unwrap();
    let text = current
        .footnotes
        .footnotes
        .iter()
        .find(|note| note.id > 0)
        .unwrap()
        .paragraphs[0]
        .accepted_text();
    assert_eq!(text, " A&B");
}

#[cfg(test)]
#[test]
fn scoped_template_imported_note_keeps_actual_published_source_for_next_reader() {
    let mut original = Document::new();
    let id = original.add_footnote("A&B");
    original.add_paragraph("body").add_footnote_ref(id);
    original.to_bytes().unwrap();
    let data = Value::Null;
    let operation = rdocx_oxml::document::ScopedWordOperation::with_template_word_inputs(
        original
            .package
            .parts
            .iter()
            .map(|(part, bytes)| (part.as_str(), bytes.as_slice()))
            .collect(),
        original
            .footnotes
            .footnotes
            .iter()
            .flat_map(|note| note.paragraphs.iter())
            .collect(),
        vec![&original.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut candidate, lineage) = TemplateWordLineage::clone_for_staging(&original);
    let mut evaluation = TemplateEvaluation {
        data: &data,
        operation,
        lineage,
        sources: Default::default(),
        versioned: true,
        occurrence: 0,
        formatter: None,
    };
    let programs = candidate
        .template_programs_scoped(
            &mut evaluation.operation,
            &evaluation.lineage,
            &mut evaluation.sources,
        )
        .unwrap();
    let body = programs
        .iter()
        .find(|program| program.story.kind() == crate::StoryKind::Body)
        .unwrap();
    let parent = body.source.as_ref().unwrap();
    let item = &body.items[0];
    let selected = evaluation
        .operation
        .compose_sources(&[(
            Some(parent.as_ref()),
            item.range.start - body.ranges[0].start..item.range.end - body.ranges[0].start,
            &[],
        )])
        .unwrap();
    let selected = evaluation
        .operation
        .close_source_namespaces(selected, &item.scope)
        .unwrap();
    let mut allocations = Vec::new();
    let imported = candidate
        .template_import_repeated_blocks(
            body,
            crate::document::TemplateFragmentSelection::Source(Box::new(selected)),
            &mut allocations,
            &mut HashSet::new(),
            &mut evaluation,
        )
        .unwrap();
    assert!(
        imported.source.is_some(),
        "the main import returns its actual source"
    );
    assert_eq!(allocations.len(), 1);
    let (_, _, old, part, new) = &allocations[0];
    assert_eq!(*old, id);
    assert_ne!(old, new);
    let identities = candidate
        .template_note_identities(Some(&mut crate::document::TemplateMainWriter {
            operation: &mut evaluation.operation,
            lineage: &evaluation.lineage,
            sources: &mut evaluation.sources,
        }))
        .unwrap();
    assert!(
        identities
            .iter()
            .any(|(_, found, id, _)| found == part && id == new)
    );
    let storage = candidate.package.get_part_shared(part).unwrap();
    let source = evaluation
        .sources
        .related_source(part, storage)
        .expect("newly allocated note keeps its actual publication source");
    let notes = evaluation.operation.source_notes(source).unwrap();
    let (index, note) = notes
        .footnotes
        .iter()
        .enumerate()
        .find(|(_, note)| note.id == *new)
        .unwrap();
    let emitted = evaluation
        .operation
        .source_model_note_from_read(note, false, source, index)
        .unwrap();
    let xml = std::str::from_utf8(emitted.as_bytes()).unwrap();
    let start = xml.find("<w:p").unwrap();
    let end = start + xml[start..].find("</w:p>").unwrap() + 6;
    let paragraph = evaluation
        .operation
        .compose_sources(&[(Some(&emitted), start..end, &[])])
        .unwrap();
    let paragraph = evaluation
        .operation
        .close_source_namespaces(
            paragraph,
            &BTreeMap::from([("w".into(), rdocx_oxml::namespace::W_NS.into())]),
        )
        .unwrap();
    let texts = evaluation
        .operation
        .template_replaceable_texts(&paragraph)
        .unwrap();
    assert_eq!(texts.concat(), " A&B");
    evaluation
        .operation
        .reserve_template_text_copy(texts.iter().map(String::len).sum())
        .unwrap();
}

#[cfg(test)]
#[test]
fn scoped_template_initial_revision_and_ruby_keep_real_visible_inputs() {
    let xml = format!(
        r#"<w:document xmlns:w="{}" xmlns:v="urn:test:kept"><w:body><w:p><w:ins w:id="1" w:author="A"><w:r><w:t>A&amp;B</w:t></w:r></w:ins><w:ruby><w:rubyPr/><w:rt><w:r><w:t>TOKEN</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>BASE</w:t></w:r></w:rubyBase><v:opaque data="kept"/></w:ruby></w:p></w:body></w:document>"#,
        rdocx_oxml::namespace::W_NS
    );
    let mut original = Document::new();
    original.document = crate::document::DocumentModel::unproved(
        rdocx_oxml::document::CT_Document::from_xml(xml.as_bytes()).unwrap(),
    );
    let data = Value::Null;
    let mut operation = rdocx_oxml::document::ScopedWordOperation::with_template_word_inputs(
        Vec::new(),
        Vec::new(),
        vec![&original.document.body],
        Vec::new(),
        Vec::new(),
        Vec::new(),
        &data,
    )
    .unwrap();
    let (mut current, lineage) = TemplateWordLineage::clone_for_staging(&original);
    let mut sources = crate::document::TemplateMainSources::default();
    let programs = current
        .template_programs_scoped(&mut operation, &lineage, &mut sources)
        .unwrap();
    let body = programs
        .iter()
        .find(|program| program.story.kind() == crate::StoryKind::Body)
        .unwrap();
    let item = &body.items[0];
    let parent = body.source.as_ref().unwrap();
    assert!(
        std::str::from_utf8(parent.as_bytes())
            .unwrap()
            .contains("data=\"kept\"")
    );
    let paragraph = operation
        .compose_sources(&[(
            Some(parent.as_ref()),
            item.range.start - body.ranges[0].start..item.range.end - body.ranges[0].start,
            &[],
        )])
        .unwrap();
    let paragraph = operation
        .close_source_namespaces(paragraph, &item.scope)
        .unwrap();
    let texts = operation.template_replaceable_texts(&paragraph).unwrap();
    assert_eq!(texts.concat(), "A&BBASE");
    operation
        .reserve_template_text_copy(texts.iter().map(String::len).sum())
        .unwrap();
}

#[test]
fn template_nonemitting_repeated_scalars_do_not_produce_text() {
    let data = serde_json::json!({"enabled": false, "value": "x".repeat(10_000)});
    for versioned in [false, true] {
        let mut document = Document::new();
        document.add_paragraph("{% if enabled %}");
        document.add_paragraph(&"{{ value }}".repeat(600));
        document.add_paragraph("{% endif %}");
        if versioned {
            document.render_template_versioned(&data, 1).unwrap();
        } else {
            document.render_template(&data).unwrap();
        }
        assert!(document.paragraphs().is_empty());
    }
}

#[cfg(test)]
fn expression_test_operation<'a>(
    data: &'a Value,
    word: &'a [u8],
) -> rdocx_oxml::document::ScopedWordOperation<'a> {
    let mut operation = rdocx_oxml::document::ScopedWordOperation::with_template_inputs(
        vec![("word/document.xml", word)],
        vec![],
        data,
    )
    .unwrap();
    let source = operation
        .source_part("word/document.xml", 0..word.len(), &BTreeMap::new())
        .unwrap();
    operation.template_replaceable_texts(&source).unwrap();
    operation
}

#[test]
fn template_nonemitting_filters_preserve_final_character_errors_and_type_order() {
    let data = serde_json::json!({
        "null": null, "array": ["x"], "bad_members": ["\0", {}],
        "edges": "\u{b}\u{c} Straße \u{b}\u{c}", "middle": "x\u{b}y",
        "sharp": "ß\0", "turkish": "İ\0", "greek": "ΟΣ\0",
        "joined": ["\u{b}Straße", "İ\u{c}"], "n": 1.25e-10,
        "bool": false, "empty": [], "nested": {"name": "value"},
    });
    for source in [
        "null | trim | default(\"fallback\")",
        "missing | upper | default(\"fallback\")",
        "array | upper | join(\",\")",
        "array | join(\",\") | join(\";\")",
        "bad_members | join(\",\")",
        "edges | upper | trim | lower",
        "middle | trim",
        "sharp | upper",
        "turkish | lower",
        "greek | lower | upper | lower",
        "joined | join(\" \\u000b \\u000c \\u000b \\u000c \\u000b\") | trim",
        "joined | join(\"\") | trim | upper | lower",
        "n | upper | trim",
        "bool | lower",
        "empty | join(\"\\u0000\")",
        "missing | default(\" \\u000bok\\u000c \\u000b\") | trim",
        "nested.name[0]",
        "n.name",
    ] {
        let expression = parse_expression(source).unwrap();
        let escaped = source.replace('&', "&amp;").replace('<', "&lt;");
        let word = format!(
            r#"<w:p xmlns:w="{}"><w:r><w:t>{escaped}</w:t></w:r></w:p>"#,
            rdocx_oxml::namespace::W_NS,
        );
        // Word cannot hold literal VT/FF. Only the expression's JSON string
        // spelling enters this physical Word fixture.
        let word = word.replace('\u{b}', "\\u000b").replace('\u{c}', "\\u000c");
        let mut emitted_operation = expression_test_operation(&data, word.as_bytes());
        let emitted =
            expression_text(&expression, &data, &[], &mut emitted_operation).and_then(|text| {
                oxml_core::xml::reject_non_xml_characters("template replacement", &text)?;
                Ok(text)
            });
        let mut validation_operation = expression_test_operation(&data, word.as_bytes());
        let validated = evaluate_expression(
            &expression,
            &data,
            &[],
            &mut validation_operation,
            None,
            false,
        );
        assert_eq!(
            emitted.as_ref().map(|_| ()).map_err(ToString::to_string),
            validated.as_ref().map(|_| ()).map_err(ToString::to_string),
            "{source}",
        );
        if let Ok(text) = emitted {
            assert_eq!(validated.unwrap(), None);
            if source.starts_with("null |") {
                assert_eq!(text, "");
            }
        }
    }
    for (source, position) in [
        ("sharp | upper", 3),
        ("turkish | lower", 3),
        ("greek | lower", 3),
    ] {
        let mut operation =
            rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
                .unwrap();
        let error = evaluate_expression(
            &parse_expression(source).unwrap(),
            &data,
            &[],
            &mut operation,
            None,
            false,
        )
        .unwrap_err()
        .to_string();
        assert!(
            error.contains(&format!("U+0000 at character {position},")),
            "{error}"
        );
    }
}

#[test]
fn template_nonemitting_borrowed_recipes_neither_copy_nor_activate_json_credit() {
    let data = serde_json::json!({
        "value": " x ".repeat(10_000),
        "joined": ["x".repeat(10_000), false, null, 1.25e100],
    });
    let mut operation =
        rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
            .unwrap();
    for source in [
        "value | trim | upper | lower",
        "joined | join(\" / \\u000b \\u000c \\u000b \\u000c \\u000b\") | upper | trim",
        "missing | default(\"long fallback\") | upper | trim",
    ] {
        let expression = parse_expression(source).unwrap();
        for _ in 0..40 {
            // The join's internal separator is deliberately invalid, so its
            // ordinary character error is allowed, but never a work refusal.
            match evaluate_expression(&expression, &data, &[], &mut operation, None, false) {
                Ok(value) => assert_eq!(value, None),
                Err(error) => assert!(
                    error.to_string().contains("XML 1.0 cannot carry"),
                    "{error}"
                ),
            }
        }
    }
    let fallback = format!(
        "missing | default(\"{}\") | upper | trim",
        "x".repeat(30_000)
    );
    let fallback = parse_expression(&fallback).unwrap();
    for _ in 0..40 {
        assert_eq!(
            evaluate_expression(&fallback, &data, &[], &mut operation, None, false).unwrap(),
            None
        );
    }
    assert!(
        operation.reserve_template_text_copy(1).is_err(),
        "validation supplied no JSON copy allowance"
    );
    // Sticky output refusal does not turn pure validation into an emitter.
    assert!(
        evaluate_expression(
            &parse_expression("value | upper").unwrap(),
            &data,
            &[],
            &mut operation,
            None,
            false
        )
        .is_ok()
    );
}

#[test]
fn template_nonemitting_false_program_keeps_errors_and_deferred_paths() {
    for versioned in [false, true] {
        for (source, succeeds) in [
            ("{{ missing }}", false),
            ("{{ row.name }}", true),
            ("{{ known.child }}", false),
        ] {
            let mut document = Document::new();
            for paragraph in ["{% for row in rows %}", source, "{% endfor %}"] {
                document.add_paragraph(paragraph);
            }
            let data = serde_json::json!({"rows": [], "known": 1});
            let before = document.document.body.content.len();
            let result = if versioned {
                document.render_template_versioned(&data, 1)
            } else {
                document.render_template(&data)
            };
            // Unknown unqualified fields are deferred in an empty loop too.
            let succeeds = succeeds || source == "{{ missing }}";
            assert_eq!(
                result.is_ok(),
                succeeds,
                "{versioned}: {source}: {result:?}"
            );
            if !succeeds {
                assert_eq!(document.document.body.content.len(), before);
            }
        }
        let mut document = Document::new();
        for paragraph in ["{% if enabled %}", "{{ missing }}", "{% endif %}"] {
            document.add_paragraph(paragraph);
        }
        let data = serde_json::json!({"enabled": false});
        let result = if versioned {
            document.render_template_versioned(&data, 1)
        } else {
            document.render_template(&data)
        };
        assert!(result.is_err());
    }
}

#[test]
fn template_nonemitting_refactor_preserves_emitting_reservation_error_order() {
    let data = serde_json::json!({"value": "xxxxx", "joined": ["xxxxx", {}]});
    for source in ["value | upper | join(\",\")", "joined | join(\",\")"] {
        let mut operation =
            rdocx_oxml::document::ScopedWordOperation::with_template_inputs(vec![], vec![], &data)
                .unwrap();
        let selected = if source.starts_with("value") {
            operation
                .template_datum_key(&data, "value")
                .unwrap()
                .unwrap()
        } else {
            let array = operation
                .template_datum_key(&data, "joined")
                .unwrap()
                .unwrap();
            operation.template_datum_index(array, 0).unwrap().unwrap()
        };
        let bytes = operation.template_scalar_input_bytes(selected).unwrap();
        let allowance = (bytes + 64) * (rdocx_oxml::namespace::W_NS.len() + 64);
        operation.reserve_template_text_copy(allowance).unwrap();
        let error = expression_text(
            &parse_expression(source).unwrap(),
            &data,
            &[],
            &mut operation,
        )
        .unwrap_err()
        .to_string();
        assert!(
            error.contains("template text copy exceeds selected input work"),
            "{error}"
        );
    }
}

#[test]
fn template_default_copy_is_reserved_against_actual_word_input() {
    let data = serde_json::json!({});
    let word = format!(
        r#"<w:p xmlns:w="{}"><w:r><w:t>missing | default("fallback")</w:t></w:r></w:p>"#,
        rdocx_oxml::namespace::W_NS
    );
    let mut operation = expression_test_operation(&data, word.as_bytes());
    let expression = parse_expression("missing | default(\"fallback\")").unwrap();
    let mut copied = 0;
    loop {
        match expression_text(&expression, &data, &[], &mut operation) {
            Ok(text) => {
                assert_eq!(text, "fallback");
                copied += text.len();
            }
            Err(error) => {
                assert!(
                    error
                        .to_string()
                        .contains("template text copy exceeds selected input work")
                );
                break;
            }
        }
        assert!(copied <= word.len() * (rdocx_oxml::namespace::W_NS.len() + 64));
    }
    assert!(copied > 0);
}
