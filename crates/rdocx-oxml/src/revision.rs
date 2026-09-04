//! Tracked revision metadata and read-only content projections.

use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use crate::content_control::{CT_Sdt, SdtContent};
use crate::document::{BodyContent, CT_Document, CT_SectPr};
use crate::numbering::{parse_scoped_ppr, parse_scoped_rpr, word_prefixes_at};
use crate::properties::{CT_PPr, CT_RPr, is_word_element};
use crate::raw_xml::capture_empty_element;
use crate::table::{CT_Row, CT_Tbl, CT_TblPr, CT_Tc, CellContent};
use crate::text::{CT_P, CT_R, hyperlink_revision_index};

const MAX_REVISION_NESTING_DEPTH: usize = 32;

/// The modeled tracked-change element kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionKind {
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    RunPropertyChange,
    ParagraphPropertyChange,
    TablePropertyChange,
    SectionPropertyChange,
}

/// The typed content projected from a preserved revision subtree.
#[derive(Debug, Clone, PartialEq)]
pub enum RevisionContent {
    Runs(Vec<CT_R>),
    Marker,
    PriorRunProperties(Box<CT_RPr>),
    PriorParagraphProperties(Box<CT_PPr>),
    PriorTableProperties(Box<CT_TblPr>),
    PriorSectionProperties(Box<CT_SectPr>),
}

/// A tracked revision whose captured subtree remains its serialization source.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_camel_case_types)]
pub struct CT_Revision {
    kind: RevisionKind,
    id: i32,
    author: String,
    timestamp: Option<String>,
    raw_xml: Vec<u8>,
    content: RevisionContent,
    content_paragraph: Option<Box<CT_P>>,
    nested_revisions: Vec<(usize, CT_Revision)>,
}

impl CT_Revision {
    pub(crate) fn from_raw(raw_xml: Vec<u8>, word_prefixes: &[String]) -> Option<Self> {
        Self::try_from_raw(raw_xml, word_prefixes).ok()
    }

    pub(crate) fn into_raw_xml(self) -> Vec<u8> {
        self.raw_xml
    }

    fn try_from_raw(raw_xml: Vec<u8>, word_prefixes: &[String]) -> crate::Result<Self> {
        validate_revision_nesting_depth(&raw_xml, word_prefixes)?;
        let mut reader = Reader::from_reader(raw_xml.as_slice());
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let (kind, prefixes) = loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) => {
                    let prefixes = word_prefixes_at(&start, word_prefixes)?;
                    let kind =
                        revision_kind(start.name().as_ref(), &prefixes).ok_or_else(|| {
                            crate::OxmlError::InvalidValue(
                                "unsupported revision element".to_owned(),
                            )
                        })?;
                    let id = required_word_attribute(&start, b"id", &prefixes)?.parse()?;
                    let author = required_word_attribute(&start, b"author", &prefixes)?;
                    let timestamp = optional_word_attribute(&start, b"date", &prefixes)?;
                    break ((kind, id, author, timestamp), prefixes);
                }
                Event::Empty(start) => {
                    let prefixes = word_prefixes_at(&start, word_prefixes)?;
                    let kind =
                        revision_kind(start.name().as_ref(), &prefixes).ok_or_else(|| {
                            crate::OxmlError::InvalidValue(
                                "unsupported revision element".to_owned(),
                            )
                        })?;
                    let id = required_word_attribute(&start, b"id", &prefixes)?.parse()?;
                    let author = required_word_attribute(&start, b"author", &prefixes)?;
                    let timestamp = optional_word_attribute(&start, b"date", &prefixes)?;
                    return Ok(Self {
                        kind,
                        id,
                        author,
                        timestamp,
                        raw_xml,
                        content: RevisionContent::Marker,
                        content_paragraph: None,
                        nested_revisions: Vec::new(),
                    });
                }
                Event::Eof => {
                    return Err(crate::OxmlError::MissingElement(
                        "revision element".to_owned(),
                    ));
                }
                _ => {}
            }
            buffer.clear();
        };
        let (kind, id, author, timestamp) = kind;
        let (content, content_paragraph, nested_revisions) =
            if matches!(kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
                let (content, nested_revisions) = parse_content(&mut reader, kind, &prefixes)?;
                let paragraph = parse_accepted_revision_content(&raw_xml, &prefixes)?;
                (content, Some(Box::new(paragraph)), nested_revisions)
            } else {
                let (content, nested_revisions) = parse_content(&mut reader, kind, &prefixes)?;
                (content, None, nested_revisions)
            };
        Ok(Self {
            kind,
            id,
            author,
            timestamp,
            raw_xml,
            content,
            content_paragraph,
            nested_revisions,
        })
    }

    pub fn kind(&self) -> RevisionKind {
        self.kind
    }

    pub fn id(&self) -> i32 {
        self.id
    }

    pub fn author(&self) -> &str {
        &self.author
    }

    pub fn timestamp(&self) -> Option<&str> {
        self.timestamp.as_deref()
    }

    pub fn content(&self) -> &RevisionContent {
        &self.content
    }

    /// Return the paragraph projection for an insertion, when it has inline wrappers.
    #[doc(hidden)]
    pub fn content_paragraph(&self) -> Option<&CT_P> {
        self.content_paragraph.as_deref()
    }

    /// Return nested revision wrappers at their direct-run boundaries.
    #[doc(hidden)]
    pub fn nested_revisions(&self) -> &[(usize, CT_Revision)] {
        &self.nested_revisions
    }

    pub(crate) fn write_xml<W: std::io::Write>(
        &self,
        writer: &mut quick_xml::Writer<W>,
    ) -> crate::Result<()> {
        writer.get_mut().write_all(&self.raw_xml)?;
        Ok(())
    }

    pub(crate) fn write_xml_with_word_override<W: std::io::Write>(
        &self,
        writer: &mut quick_xml::Writer<W>,
        foreign_word_namespace: Option<&str>,
    ) -> crate::Result<()> {
        crate::text::write_raw_with_word_override(writer, &self.raw_xml, foreign_word_namespace)
    }
}

fn validate_revision_nesting_depth(raw_xml: &[u8], word_prefixes: &[String]) -> crate::Result<()> {
    let mut reader = Reader::from_reader(raw_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes = vec![(word_prefixes.to_vec(), false)];
    let mut revision_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let inherited = &scopes.last().expect("root namespace scope").0;
                let prefixes = word_prefixes_at(&start, inherited)?;
                let is_revision = revision_kind(start.name().as_ref(), &prefixes).is_some();
                if is_revision {
                    revision_depth += 1;
                    if revision_depth > MAX_REVISION_NESTING_DEPTH {
                        return Err(crate::OxmlError::InvalidValue(
                            "revision nesting depth exceeds reader limit".to_owned(),
                        ));
                    }
                }
                scopes.push((prefixes, is_revision));
            }
            Event::Empty(start) => {
                let inherited = &scopes.last().expect("root namespace scope").0;
                let prefixes = word_prefixes_at(&start, inherited)?;
                if revision_kind(start.name().as_ref(), &prefixes).is_some()
                    && revision_depth + 1 > MAX_REVISION_NESTING_DEPTH
                {
                    return Err(crate::OxmlError::InvalidValue(
                        "revision nesting depth exceeds reader limit".to_owned(),
                    ));
                }
            }
            Event::End(_) => {
                let (_, was_revision) = scopes.pop().expect("balanced namespace scope");
                if was_revision {
                    revision_depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(())
}

impl CT_Document {
    /// Return every valid modeled main-document revision in document order.
    pub fn revisions(&self) -> Vec<&CT_Revision> {
        let mut revisions = Vec::new();
        for content in &self.body.content {
            collect_body_content(content, &mut revisions);
        }
        if let Some(section) = &self.body.sect_pr {
            collect_section_properties(section, &mut revisions);
        }
        revisions
    }
}

fn collect_body_content<'a>(content: &'a BodyContent, revisions: &mut Vec<&'a CT_Revision>) {
    match content {
        BodyContent::Paragraph(paragraph) => collect_paragraph(paragraph, revisions),
        BodyContent::Table(table) => collect_table(table, revisions),
        BodyContent::ContentControl(control) => collect_control(control, revisions),
        BodyContent::RawXml(_) => {}
    }
}

fn collect_paragraph<'a>(paragraph: &'a CT_P, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(properties) = &paragraph.properties {
        if let Some(revision) = &properties.numbering_revision {
            revisions.push(revision);
        }
        if let Some(run_properties) = &properties.rpr {
            collect_run_properties(run_properties, revisions);
        }
        if let Some(section) = &properties.sect_pr {
            collect_section_properties(section, revisions);
        }
        if let Some(change) = &properties.change {
            revisions.push(change);
        }
    }

    for boundary in 0..=paragraph.runs.len() {
        let raw_count = paragraph
            .extra_xml
            .iter()
            .filter(|(at, _)| *at == boundary)
            .count();
        for raw_index in 0..=raw_count {
            for (_, _, _, control) in
                paragraph
                    .content_controls
                    .iter()
                    .filter(|(at, raw_before, _, _)| {
                        *at == boundary && (*raw_before).min(raw_count) == raw_index
                    })
            {
                collect_control(control, revisions);
            }
            for (_, _, revision) in paragraph.revisions.iter().filter(|(at, raw_before, _)| {
                *at == boundary
                    && hyperlink_revision_index(*raw_before).is_none()
                    && (*raw_before).min(raw_count) == raw_index
            }) {
                collect_revision(revision, revisions);
            }
            for (hyperlink_index, _hyperlink) in
                paragraph
                    .hyperlinks
                    .iter()
                    .enumerate()
                    .filter(|(_, hyperlink)| {
                        hyperlink.run_start == boundary
                            && hyperlink.run_end == boundary
                            && hyperlink.preserved_raw_before == Some(raw_index)
                    })
            {
                for (_, _, revision) in paragraph.revisions.iter().filter(|(at, slot, _)| {
                    *at == boundary && hyperlink_revision_index(*slot) == Some(hyperlink_index)
                }) {
                    collect_revision(revision, revisions);
                }
            }
        }
        for (_, _, revision) in paragraph.revisions.iter().filter(|(at, slot, _)| {
            *at == boundary
                && hyperlink_revision_index(*slot).is_some_and(|index| {
                    paragraph
                        .hyperlinks
                        .get(index)
                        .is_none_or(|hyperlink| hyperlink.preserved_raw_before.is_none())
                })
        }) {
            collect_revision(revision, revisions);
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            collect_run(run, revisions);
        }
    }
}

fn collect_run<'a>(run: &'a CT_R, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(properties) = &run.properties {
        collect_run_properties(properties, revisions);
    }
}

fn collect_run_properties<'a>(properties: &'a CT_RPr, revisions: &mut Vec<&'a CT_Revision>) {
    revisions.extend(properties.revision_markers.iter());
    if let Some(change) = &properties.change {
        revisions.push(change);
    }
}

fn collect_section_properties<'a>(properties: &'a CT_SectPr, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(change) = &properties.change {
        revisions.push(change);
    }
}

fn collect_table<'a>(table: &'a CT_Tbl, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(change) = table
        .properties
        .as_ref()
        .and_then(|properties| properties.change.as_ref())
    {
        revisions.push(change);
    }
    for boundary in 0..=table.rows.len() {
        collect_controls_at_table_boundary(table, boundary, revisions);
        if let Some(row) = table.rows.get(boundary) {
            collect_row(row, revisions);
        }
    }
}

fn collect_controls_at_table_boundary<'a>(
    table: &'a CT_Tbl,
    boundary: usize,
    revisions: &mut Vec<&'a CT_Revision>,
) {
    let raw_count = table
        .extra_xml
        .iter()
        .filter(|(at, _)| *at == boundary)
        .count();
    for raw_index in 0..=raw_count {
        for (_, _, control) in table.content_controls.iter().filter(|(at, raw_before, _)| {
            *at == boundary && (*raw_before).min(raw_count) == raw_index
        }) {
            collect_control(control, revisions);
        }
    }
}

fn collect_row<'a>(row: &'a CT_Row, revisions: &mut Vec<&'a CT_Revision>) {
    if let Some(properties) = &row.properties {
        revisions.extend(properties.revision_markers.iter());
    }
    for boundary in 0..=row.cells.len() {
        let raw_count = row
            .extra_xml
            .iter()
            .filter(|(at, _)| *at == boundary)
            .count();
        for raw_index in 0..=raw_count {
            for (_, _, control) in row.content_controls.iter().filter(|(at, raw_before, _)| {
                *at == boundary && (*raw_before).min(raw_count) == raw_index
            }) {
                collect_control(control, revisions);
            }
        }
        if let Some(cell) = row.cells.get(boundary) {
            collect_cell(cell, revisions);
        }
    }
}

fn collect_cell<'a>(cell: &'a CT_Tc, revisions: &mut Vec<&'a CT_Revision>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => collect_paragraph(paragraph, revisions),
            CellContent::Table(table) => collect_table(table, revisions),
            CellContent::ContentControl(control) => collect_control(control, revisions),
        }
    }
}

fn collect_control<'a>(control: &'a CT_Sdt, revisions: &mut Vec<&'a CT_Revision>) {
    for (content_index, content) in control.content.iter().enumerate() {
        for (_, revision) in control
            .revisions
            .iter()
            .filter(|(at, _)| *at == content_index)
        {
            collect_revision(revision, revisions);
        }
        match content {
            SdtContent::Paragraph(paragraph) => collect_paragraph(paragraph, revisions),
            SdtContent::Table(table) => collect_table(table, revisions),
            SdtContent::Row(row) => collect_row(row, revisions),
            SdtContent::Cell(cell) => collect_cell(cell, revisions),
            SdtContent::Run(run) => collect_run(run, revisions),
            SdtContent::ContentControl(control) => collect_control(control, revisions),
            SdtContent::RawXml(_) => {}
        }
    }
}

fn collect_revision<'a>(revision: &'a CT_Revision, revisions: &mut Vec<&'a CT_Revision>) {
    revisions.push(revision);
    if let Some(paragraph) = revision.content_paragraph() {
        collect_paragraph(paragraph, revisions);
        return;
    }
    if let RevisionContent::Runs(runs) = revision.content() {
        for boundary in 0..=runs.len() {
            for (_, nested) in revision
                .nested_revisions
                .iter()
                .filter(|(at, _)| *at == boundary)
            {
                collect_revision(nested, revisions);
            }
            if let Some(run) = runs.get(boundary) {
                collect_run(run, revisions);
            }
        }
    } else {
        for (_, nested) in &revision.nested_revisions {
            collect_revision(nested, revisions);
        }
    }
}

fn parse_accepted_revision_content(
    raw_xml: &[u8],
    word_prefixes: &[String],
) -> crate::Result<CT_P> {
    let mut reader = Reader::from_reader(raw_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let content_start = loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) => break reader.buffer_position() as usize,
            Event::Eof => {
                return Err(crate::OxmlError::MissingElement(
                    "insertion element".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    };
    let content_end = raw_xml
        .iter()
        .enumerate()
        .rev()
        .find_map(|(index, byte)| (*byte == b'<').then_some(index))
        .ok_or_else(|| crate::OxmlError::MissingElement("insertion close tag".to_owned()))?;
    let mut paragraph_xml = Vec::from(b"<w:p>".as_slice());
    paragraph_xml.extend_from_slice(&raw_xml[content_start..content_end]);
    paragraph_xml.extend_from_slice(b"</w:p>");
    let mut paragraph_reader = Reader::from_reader(paragraph_xml.as_slice());
    let mut paragraph_buffer = Vec::new();
    match paragraph_reader.read_event_into(&mut paragraph_buffer)? {
        Event::Start(_) => CT_P::from_xml_with_prefixes(&mut paragraph_reader, word_prefixes),
        _ => Err(crate::OxmlError::MissingElement(
            "insertion content".to_owned(),
        )),
    }
}

fn revision_kind(name: &[u8], prefixes: &[String]) -> Option<RevisionKind> {
    [
        (b"ins".as_slice(), RevisionKind::Insertion),
        (b"del".as_slice(), RevisionKind::Deletion),
        (b"moveFrom".as_slice(), RevisionKind::MoveFrom),
        (b"moveTo".as_slice(), RevisionKind::MoveTo),
        (b"rPrChange".as_slice(), RevisionKind::RunPropertyChange),
        (
            b"pPrChange".as_slice(),
            RevisionKind::ParagraphPropertyChange,
        ),
        (b"tblPrChange".as_slice(), RevisionKind::TablePropertyChange),
        (
            b"sectPrChange".as_slice(),
            RevisionKind::SectionPropertyChange,
        ),
    ]
    .into_iter()
    .find_map(|(local, kind)| is_word_element(name, local, prefixes).then_some(kind))
}

fn parse_content(
    reader: &mut Reader<&[u8]>,
    kind: RevisionKind,
    word_prefixes: &[String],
) -> crate::Result<(RevisionContent, Vec<(usize, CT_Revision)>)> {
    let mut runs = Vec::new();
    let mut nested_revisions = Vec::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                let name = start.name();
                if is_word_element(name.as_ref(), b"r", &prefixes) {
                    runs.push(CT_R::from_xml_with_prefixes(reader, &prefixes)?);
                } else if revision_kind(name.as_ref(), &prefixes).is_some() {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                } else if is_word_element(name.as_ref(), b"hyperlink", &prefixes) {
                    parse_run_content_container(
                        reader,
                        b"hyperlink",
                        &prefixes,
                        &mut runs,
                        &mut nested_revisions,
                    )?;
                } else if is_word_element(name.as_ref(), b"rPr", &prefixes)
                    && kind == RevisionKind::RunPropertyChange
                {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    return Ok((
                        RevisionContent::PriorRunProperties(Box::new(parse_scoped_rpr(
                            &raw,
                            word_prefixes,
                        )?)),
                        nested_revisions,
                    ));
                } else if is_word_element(name.as_ref(), b"pPr", &prefixes)
                    && kind == RevisionKind::ParagraphPropertyChange
                {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    return Ok((
                        RevisionContent::PriorParagraphProperties(Box::new(parse_scoped_ppr(
                            &raw,
                            word_prefixes,
                        )?)),
                        nested_revisions,
                    ));
                } else if is_word_element(name.as_ref(), b"tblPr", &prefixes)
                    && kind == RevisionKind::TablePropertyChange
                {
                    let owner_bindings =
                        crate::numbering::local_namespace_overrides(&start, word_prefixes)?;
                    return Ok((
                        RevisionContent::PriorTableProperties(Box::new(
                            CT_TblPr::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &owner_bindings,
                            )?,
                        )),
                        nested_revisions,
                    ));
                } else if is_word_element(name.as_ref(), b"sectPr", &prefixes)
                    && kind == RevisionKind::SectionPropertyChange
                {
                    let owner_bindings =
                        crate::numbering::local_namespace_overrides(&start, word_prefixes)?;
                    return Ok((
                        RevisionContent::PriorSectionProperties(Box::new(
                            CT_SectPr::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &owner_bindings,
                            )?,
                        )),
                        nested_revisions,
                    ));
                } else {
                    reader.read_to_end_into(name, &mut Vec::new())?;
                }
            }
            Event::Empty(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                if revision_kind(start.name().as_ref(), &prefixes).is_some() {
                    let raw = capture_empty_element(&start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                }
            }
            Event::End(_) | Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if runs.is_empty() {
        Ok((RevisionContent::Marker, nested_revisions))
    } else {
        Ok((RevisionContent::Runs(runs), nested_revisions))
    }
}

fn parse_run_content_container(
    reader: &mut Reader<&[u8]>,
    closing_local: &[u8],
    word_prefixes: &[String],
    runs: &mut Vec<CT_R>,
    nested_revisions: &mut Vec<(usize, CT_Revision)>,
) -> crate::Result<()> {
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                let name = start.name();
                if is_word_element(name.as_ref(), b"r", &prefixes) {
                    runs.push(CT_R::from_xml_with_prefixes(reader, &prefixes)?);
                } else if revision_kind(name.as_ref(), &prefixes).is_some() {
                    let raw = crate::raw_xml::capture_element(reader, &start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                } else if is_word_element(name.as_ref(), b"hyperlink", &prefixes) {
                    parse_run_content_container(
                        reader,
                        b"hyperlink",
                        &prefixes,
                        runs,
                        nested_revisions,
                    )?;
                } else {
                    reader.read_to_end_into(name, &mut Vec::new())?;
                }
            }
            Event::Empty(start) => {
                let prefixes = word_prefixes_at(&start, word_prefixes)?;
                if revision_kind(start.name().as_ref(), &prefixes).is_some() {
                    let raw = capture_empty_element(&start)?;
                    if let Some(revision) = CT_Revision::from_raw(raw, &prefixes) {
                        nested_revisions.push((runs.len(), revision));
                    }
                }
            }
            Event::End(end)
                if is_word_element(end.name().as_ref(), closing_local, word_prefixes) =>
            {
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn required_word_attribute(
    start: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> crate::Result<String> {
    optional_word_attribute(start, local, prefixes)?.ok_or_else(|| {
        crate::OxmlError::MissingElement(format!("w:{} attribute", String::from_utf8_lossy(local)))
    })
}

fn optional_word_attribute(
    start: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> crate::Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let Some(separator) = key.iter().position(|byte| *byte == b':') else {
            continue;
        };
        if key.get(separator + 1..) == Some(local)
            && prefixes
                .iter()
                .any(|prefix| prefix.as_bytes() == &key[..separator])
        {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::namespace::W_NS;
    use crate::text::RunContent;

    fn nested_insertions(depth: usize) -> Vec<u8> {
        let mut xml = format!(r#"<w:ins xmlns:w="{W_NS}" w:id="0" w:author="Ada">"#);
        for id in 1..depth {
            xml.push_str(&format!(r#"<w:ins w:id="{id}" w:author="Ada">"#));
        }
        xml.push_str("<w:r><w:t>visible</w:t></w:r>");
        for _ in 0..depth {
            xml.push_str("</w:ins>");
        }
        xml.into_bytes()
    }

    #[test]
    fn revision_nesting_depth_is_bounded_before_recursive_projection() {
        let prefixes = ["w".to_owned()];
        assert!(
            validate_revision_nesting_depth(
                &nested_insertions(MAX_REVISION_NESTING_DEPTH),
                &prefixes,
            )
            .is_ok()
        );
        let too_deep = nested_insertions(MAX_REVISION_NESTING_DEPTH + 1);
        assert!(validate_revision_nesting_depth(&too_deep, &prefixes).is_err());
        assert!(CT_Revision::from_raw(too_deep, &prefixes).is_none());
    }

    #[test]
    fn revision_attributes_are_prefix_tolerant_and_namespace_checked() {
        let raw = format!(
            r#"<word:ins xmlns:word="{W_NS}" word:id="7" word:author="Ada"><word:r><word:t>added</word:t></word:r></word:ins>"#
        )
        .into_bytes();
        let revision = CT_Revision::from_raw(raw, &["word".to_owned()])
            .expect("aliased WordprocessingML revision should be projected");
        assert_eq!(revision.id(), 7);
        assert_eq!(revision.author(), "Ada");
        assert_eq!(revision.timestamp(), None);
        assert_eq!(revision.kind(), RevisionKind::Insertion);

        let escaped = format!(
            r#"<word:ins xmlns:word="{W_NS}" word:id="8" word:author="A &amp; B" word:date="2026-08-17T09:00:00+01:00 &amp; later"/>"#
        )
        .into_bytes();
        let escaped = CT_Revision::from_raw(escaped, &["word".to_owned()])
            .expect("escaped revision metadata should parse");
        assert_eq!(escaped.author(), "A & B");
        assert_eq!(
            escaped.timestamp(),
            Some("2026-08-17T09:00:00+01:00 & later")
        );

        let foreign = br#"<x:ins xmlns:x="urn:not-word" x:id="9" x:author="Eve"/>"#.to_vec();
        assert!(CT_Revision::from_raw(foreign, &["word".to_owned()]).is_none());

        let missing_author = format!(r#"<word:del xmlns:word="{W_NS}" word:id="9"/>"#).into_bytes();
        assert!(CT_Revision::from_raw(missing_author, &["word".to_owned()]).is_none());

        let document_xml = format!(
            r#"<word:document xmlns:word="{W_NS}" xmlns:x="urn:not-word"><word:body><word:p><word:pPr><word:pPrChange word:id="10" word:author="Gia"><word:pPr><word:jc word:val="right"/></word:pPr></word:pPrChange><x:pPrChange x:id="11" x:author="No"/><word:pPrChange word:id="12"/></word:pPr></word:p></word:body></word:document>"#
        );
        let document = CT_Document::from_xml(document_xml.as_bytes()).expect("document parses");
        assert_eq!(document.revisions().len(), 1);
        assert_eq!(document.revisions()[0].id(), 10);
        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        assert!(output.contains(r#"<word:pPrChange word:id="10" word:author="Gia">"#));
        assert!(output.contains(r#"<x:pPrChange x:id="11" x:author="No"/>"#));
        assert!(output.contains(r#"<word:pPrChange word:id="12"/>"#));

        let collision_xml = format!(
            r#"<word:document xmlns:word="{W_NS}" xmlns:x="urn:not-word"><word:body>
<word:tbl><word:tblPr><word:tblPrChange word:id="20" word:author="Tab"><word:tblPr><x:jc x:val="center"/><word:jc x:val="right"/><word:tblBorders><x:top x:val="single" word:sz="invalid"/><word:bottom x:val="single"/></word:tblBorders><word:tblCellMar><x:top x:w="120"/><word:bottom x:w="240"/></word:tblCellMar></word:tblPr></word:tblPrChange></word:tblPr><word:tblGrid/><word:tr><word:tc><word:p/></word:tc></word:tr></word:tbl>
<word:sectPr><word:sectPrChange word:id="21" word:author="Sec"><word:sectPr><x:titlePg/><word:pgSz x:w="12240" x:h="15840"/><word:pgMar x:producer="not-a-number" word:top="720"/></word:sectPr></word:sectPrChange></word:sectPr>
</word:body></word:document>"#
        );
        let collision = CT_Document::from_xml(collision_xml.as_bytes()).expect("document parses");
        let collision_revisions = collision.revisions();
        assert!(matches!(
            collision_revisions[0].content(),
            RevisionContent::PriorTableProperties(properties)
                if properties.jc.is_none()
                    && properties
                        .borders
                        .as_ref()
                        .is_some_and(|borders| borders.top.is_none()
                            && borders.bottom.as_ref().is_some_and(|edge| {
                                edge.val == crate::shared::ST_Border::None
                            }))
                    && properties
                        .cell_margin
                        .as_ref()
                        .is_some_and(|margin| margin.top.is_none() && margin.bottom.is_none())
        ));
        assert!(matches!(
            collision_revisions[1].content(),
            RevisionContent::PriorSectionProperties(properties)
                if properties.title_pg.is_none()
                    && properties.page_width.is_none()
                    && properties.page_height.is_none()
                    && properties.margin_top == Some(crate::units::Twips(720))
        ));

        let paragraph_collision_xml = format!(
            r#"<word:document xmlns:word="{W_NS}" xmlns:x="urn:not-word"><word:body><word:p><word:pPr><word:pPrChange word:id="22" word:author="Para"><word:pPr><word:pBdr><x:top x:val="single" word:sz="invalid"/><word:bottom x:val="single"/></word:pBdr></word:pPr></word:pPrChange></word:pPr></word:p></word:body></word:document>"#
        );
        let paragraph_collision =
            CT_Document::from_xml(paragraph_collision_xml.as_bytes()).expect("document parses");
        assert!(matches!(
            paragraph_collision.revisions()[0].content(),
            RevisionContent::PriorParagraphProperties(properties)
                if properties.borders.as_ref().is_some_and(|borders| {
                    borders.top.is_none()
                        && borders.bottom.as_ref().is_some_and(|edge| {
                            edge.val == crate::shared::ST_Border::None
                        })
                })
        ));

        let control_xml = format!(
            r#"<word:document xmlns:word="{W_NS}"><word:body><word:p><word:sdt><word:sdtContent><word:ins word:id="30" word:author="Sdt"><word:r><word:t>inside</word:t></word:r></word:ins></word:sdtContent></word:sdt></word:p></word:body></word:document>"#
        );
        let control = CT_Document::from_xml(control_xml.as_bytes()).expect("document parses");
        assert_eq!(
            control
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            vec![30]
        );
    }

    #[test]
    fn insertion_content_retains_inline_paragraph_structure() {
        let raw = format!(
            r#"<w:ins xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:ins>"#
        )
        .into_bytes();
        let revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("insertion parses");
        let paragraph = revision
            .content_paragraph()
            .expect("insertion exposes paragraph content");

        assert_eq!(paragraph.text(), "beforelinked");
        assert_eq!(paragraph.hyperlinks.len(), 1);
        assert_eq!(paragraph.hyperlinks[0].rel_id.as_deref(), Some("rId9"));
    }

    #[test]
    fn accepted_content_revisions_retain_inline_content_controls() {
        for kind in ["ins", "moveTo"] {
            let raw = format!(
                r#"<w:{kind} xmlns:w="{W_NS}" w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt></w:{kind}>"#
            )
            .into_bytes();
            let revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("revision parses");
            let paragraph = revision
                .content_paragraph()
                .expect("accepted revision exposes paragraph content");

            assert_eq!(paragraph.text(), "beforecontrol");
            assert_eq!(paragraph.content_controls.len(), 1);
            let crate::content_control::SdtContent::Run(run) =
                &paragraph.content_controls[0].3.content[0]
            else {
                panic!("expected controlled run");
            };
            assert_eq!(run.text(), "control");
        }
    }

    #[test]
    fn insertion_content_retains_nested_revision_projection() {
        let raw = format!(
            r#"<w:ins xmlns:w="{W_NS}" w:id="7" w:author="Ada"><w:r><w:t>before</w:t></w:r><w:del w:id="8" w:author="Ben"><w:r><w:delText>removed</w:delText></w:r></w:del><w:ins w:id="9" w:author="Cy"><w:r><w:t>nested</w:t></w:r></w:ins><w:r><w:t>after</w:t></w:r></w:ins>"#
        )
        .into_bytes();
        let revision = CT_Revision::from_raw(raw, &["w".to_owned()]).expect("insertion parses");

        let RevisionContent::Runs(runs) = revision.content() else {
            panic!("insertion exposes direct runs");
        };
        assert_eq!(
            runs.iter().map(CT_R::text).collect::<String>(),
            "beforeafter"
        );
        assert_eq!(
            revision
                .nested_revisions()
                .iter()
                .map(|(boundary, nested)| (*boundary, nested.kind()))
                .collect::<Vec<_>>(),
            vec![(1, RevisionKind::Deletion), (1, RevisionKind::Insertion),]
        );
        let paragraph = revision
            .content_paragraph()
            .expect("insertion exposes paragraph content");
        assert_eq!(paragraph.text(), "beforeafter");
        assert_eq!(
            paragraph
                .revisions
                .iter()
                .map(|(_, _, nested)| nested.kind())
                .collect::<Vec<_>>(),
            vec![RevisionKind::Deletion, RevisionKind::Insertion]
        );
    }

    #[test]
    fn property_changes_write_in_their_schema_final_slots() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body>
<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>
<w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Ben"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="3" w:author="Cy"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="4" w:author="Dee"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");

        for change in ["pPrChange", "rPrChange", "tblPrChange", "sectPrChange"] {
            assert_eq!(
                output.matches(&format!("<w:{change} ")).count(),
                1,
                "{output}"
            );
        }
        assert!(
            output.find("<w:jc w:val=\"center\"").unwrap() < output.find("<w:pPrChange ").unwrap()
        );
        assert!(output.find("<w:b").unwrap() < output.find("<w:rPrChange ").unwrap());
        assert!(output.find("<w:tblW").unwrap() < output.find("<w:tblPrChange ").unwrap());
        assert!(output.find("<w:pgSz").unwrap() < output.find("<w:sectPrChange ").unwrap());
    }

    #[test]
    fn run_property_raw_positions_cannot_follow_the_schema_final_change() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body><w:p><w:r><w:rPr><x:raw/><w:rPrChange w:id="1" w:author="Ada"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        );
        let mut document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        {
            let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
                panic!("expected paragraph");
            };
            paragraph.runs[0]
                .properties
                .as_mut()
                .expect("run properties")
                .revision_xml_positions[0] = (41, 9);
        }

        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        assert!(output.find("<x:raw").unwrap() < output.find("<w:rPrChange").unwrap());

        {
            let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
                panic!("expected paragraph");
            };
            paragraph.runs[0]
                .properties
                .as_mut()
                .expect("run properties")
                .revision_xml_positions
                .clear();
        }
        let output = String::from_utf8(document.to_xml().expect("fallback writes")).unwrap();
        assert!(output.find("<x:raw").unwrap() < output.find("<w:rPrChange").unwrap());

        let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
            panic!("expected paragraph");
        };
        paragraph.runs[0]
            .properties
            .as_mut()
            .expect("run properties")
            .revision_xml_positions
            .extend([(41, 0), (41, 1)]);
        let output = String::from_utf8(document.to_xml().expect("mismatch writes")).unwrap();
        assert!(output.find("<x:raw").unwrap() < output.find("<w:rPrChange").unwrap());
    }

    #[test]
    fn retained_run_property_children_keep_owner_local_namespace_bindings() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:rPr xmlns:x="urn:producer" xmlns:wa="{W_NS}"><x:raw/><wa:outline><wa:opaque/></wa:outline><wa:rPrChange wa:id="7" wa:author="Ada"><wa:rPr><wa:b/></wa:rPr></wa:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        assert_eq!(document.revisions()[0].id(), 7);

        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        assert_eq!(
            output.matches(r#"<x:raw xmlns:x="urn:producer"/>"#).count(),
            1
        );
        assert_eq!(
            output
                .matches(&format!(r#"<wa:outline xmlns:wa="{W_NS}">"#))
                .count(),
            1
        );
        assert_eq!(
            output
                .matches(&format!(
                    r#"<wa:rPrChange wa:id="7" wa:author="Ada" xmlns:wa="{W_NS}">"#
                ))
                .count(),
            1
        );

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(reopened.revisions()[0].id(), 7);
        let rewritten = String::from_utf8(reopened.to_xml().expect("output rewrites")).unwrap();
        assert_eq!(rewritten.matches(r#"xmlns:x="urn:producer""#).count(), 1);
        assert_eq!(
            rewritten.matches(&format!(r#"xmlns:wa="{W_NS}""#)).count(),
            2
        );
    }

    #[test]
    fn duplicate_property_changes_are_retained_without_replacing_the_first_projection() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:body>
<w:p><w:pPr><w:numPr><w:ins w:id="501" w:author="Ada"/><x:ins x:mark="num"/><w:ins w:id="502" w:author="Ada"/></w:numPr><w:pPrChange w:id="101" w:author="Ada"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange><x:pPrChange x:mark="paragraph"/><w:pPrChange w:id="102" w:author="Ada"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="201" w:author="Ada"><w:rPr><w:b/></w:rPr></w:rPrChange><x:rPrChange x:mark="run"/><w:rPrChange w:id="202" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblPrChange w:id="301" w:author="Ada"><w:tblPr><w:jc w:val="left"/></w:tblPr></w:tblPrChange><x:tblPrChange x:mark="table"/><w:tblPrChange w:id="302" w:author="Ada"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr><w:sectPrChange w:id="401" w:author="Ada"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange><x:sectPrChange x:mark="section"/><w:sectPrChange w:id="402" w:author="Ada"><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:sectPrChange></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let mut ids = document
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        ids.sort_unstable();
        assert_eq!(ids, [102, 202, 302, 402, 502]);

        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        for id in [101, 102, 201, 202, 301, 302, 401, 402, 501, 502] {
            assert_eq!(output.matches(&format!(r#"w:id="{id}""#)).count(), 1);
        }
        for owner in ["pPr", "rPr", "tblPr", "sectPr"] {
            let first = output
                .find(&format!(
                    r#"w:{owner}Change w:id="{}""#,
                    match owner {
                        "pPr" => 101,
                        "rPr" => 201,
                        "tblPr" => 301,
                        _ => 401,
                    }
                ))
                .unwrap();
            let typed = output
                .find(&format!(
                    r#"w:{owner}Change w:id="{}""#,
                    match owner {
                        "pPr" => 102,
                        "rPr" => 202,
                        "tblPr" => 302,
                        _ => 402,
                    }
                ))
                .unwrap();
            assert!(
                first < typed,
                "duplicate must precede typed schema-final change"
            );
            let foreign = output
                .find(&format!(r#"<x:{owner}Change x:mark="#))
                .unwrap();
            assert!(first < foreign && foreign < typed);
        }
        let numbering_positions = [
            output.find(r#"w:id="501""#).unwrap(),
            output.find(r#"<x:ins x:mark="num"/>"#).unwrap(),
            output.find(r#"w:id="502""#).unwrap(),
        ];
        assert!(numbering_positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        let mut reopened_ids = reopened
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        reopened_ids.sort_unstable();
        assert_eq!(reopened_ids, [102, 202, 302, 402, 502]);
    }

    #[test]
    fn every_property_revision_keeps_owner_local_aliases() {
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body>
<w:p><w:pPr xmlns:pa="{W_NS}" xmlns:px="urn:paragraph"><w:numPr xmlns:na="{W_NS}" xmlns:nx="urn:numbering"><na:ins na:id="501" na:author="Ada"/><nx:ins nx:mark="raw"/><na:ins na:id="502" na:author="Ada"/></w:numPr><pa:pPrChange pa:id="101" pa:author="Ada"><pa:pPr/></pa:pPrChange><px:pPrChange px:mark="raw"/><pa:pPrChange pa:id="102" pa:author="Ada"><pa:pPr/></pa:pPrChange></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
<w:tbl><w:tblPr xmlns:ta="{W_NS}" xmlns:tx="urn:table"><ta:tblPrChange ta:id="301" ta:author="Ada"><ta:tblPr/></ta:tblPrChange><tx:tblPrChange tx:mark="raw"/><ta:tblPrChange ta:id="302" ta:author="Ada"><ta:tblPr/></ta:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr xmlns:sa="{W_NS}" xmlns:sx="urn:section"><sa:sectPrChange sa:id="401" sa:author="Ada"><sa:sectPr/></sa:sectPrChange><sx:sectPrChange sx:mark="raw"/><sa:sectPrChange sa:id="402" sa:author="Ada"><sa:sectPr/></sa:sectPrChange></w:sectPr>
</w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let output = String::from_utf8(document.to_xml().expect("document writes")).unwrap();
        for (prefix, local) in [
            ("pa", "pPrChange"),
            ("na", "ins"),
            ("ta", "tblPrChange"),
            ("sa", "sectPrChange"),
        ] {
            assert_eq!(output.matches(&format!("<{prefix}:{local} ")).count(), 2);
            assert_eq!(
                output
                    .matches(&format!(r#"xmlns:{prefix}="{W_NS}""#))
                    .count(),
                2
            );
        }
        for (prefix, namespace, local) in [
            ("px", "urn:paragraph", "pPrChange"),
            ("nx", "urn:numbering", "ins"),
            ("tx", "urn:table", "tblPrChange"),
            ("sx", "urn:section", "sectPrChange"),
        ] {
            assert_eq!(output.matches(&format!("<{prefix}:{local} ")).count(), 1);
            assert_eq!(
                output
                    .matches(&format!(r#"xmlns:{prefix}="{namespace}""#))
                    .count(),
                1
            );
        }
        for (first, foreign, second) in [
            (101, "<px:pPrChange", 102),
            (501, "<nx:ins", 502),
            (301, "<tx:tblPrChange", 302),
            (401, "<sx:sectPrChange", 402),
        ] {
            let positions = [
                output.find(&format!(r#"id="{first}""#)).unwrap(),
                output.find(foreign).unwrap(),
                output.find(&format!(r#"id="{second}""#)).unwrap(),
            ];
            assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
        }

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        let mut ids = reopened
            .revisions()
            .iter()
            .map(|revision| revision.id())
            .collect::<Vec<_>>();
        ids.sort_unstable();
        assert_eq!(ids, [102, 302, 402, 502]);
        let rewritten = String::from_utf8(reopened.to_xml().expect("output rewrites")).unwrap();
        for prefix in ["pa", "na", "ta", "sa"] {
            assert_eq!(
                rewritten
                    .matches(&format!(r#"xmlns:{prefix}="{W_NS}""#))
                    .count(),
                2
            );
        }
        for (prefix, namespace) in [
            ("px", "urn:paragraph"),
            ("nx", "urn:numbering"),
            ("tx", "urn:table"),
            ("sx", "urn:section"),
        ] {
            assert_eq!(
                rewritten
                    .matches(&format!(r#"xmlns:{prefix}="{namespace}""#))
                    .count(),
                1
            );
        }
    }

    #[test]
    fn revision_elements_round_trip_unchanged_and_report_metadata() {
        let revisions = [
            r#"<w:ins w:id="1" w:author="Ada" w:date="2026-08-01T10:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>"#,
            r#"<w:del w:id="2" w:author="Ben"><w:r><w:delText xml:space="preserve"> gone </w:delText></w:r></w:del>"#,
            r#"<w:moveFrom w:id="3" w:author="Cy"><w:r><w:t>from</w:t></w:r></w:moveFrom>"#,
            r#"<w:moveTo w:id="4" w:author="Dee"><w:r><w:t>to</w:t></w:r></w:moveTo>"#,
        ];
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p>{}<w:r><w:rPr><w:ins w:id="5" w:author="Eve"/><w:rPrChange w:id="6" w:author="Fox"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>end</w:t></w:r><x:ins xmlns:x="urn:not-word" x:id="99" x:author="No"/></w:p><w:tbl><w:tblPr><w:tblPrChange w:id="7" w:author="Gia"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid/><w:tr><w:trPr><w:del w:id="8" w:author="Hal"/></w:trPr><w:tc><w:p><w:pPr><w:pPrChange w:id="9" w:author="Ivy"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:sectPrChange w:id="10" w:author="Jay"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>"#,
            revisions.concat()
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");
        let reported = document.revisions();
        assert_eq!(
            reported
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            (1..=10).collect::<Vec<_>>()
        );
        assert_eq!(reported[0].author(), "Ada");
        assert_eq!(reported[0].timestamp(), Some("2026-08-01T10:00:00Z"));
        assert_eq!(reported[1].timestamp(), None);
        assert_eq!(reported[1].kind(), RevisionKind::Deletion);
        let RevisionContent::Runs(runs) = reported[1].content() else {
            panic!("deletion should project runs");
        };
        assert!(matches!(
            runs[0].content.as_slice(),
            [RunContent::DeletedText(text)] if text.text == " gone " && text.preserve_space
        ));
        assert!(matches!(
            reported[5].content(),
            RevisionContent::PriorRunProperties(properties) if properties.italic == Some(true)
        ));
        assert!(matches!(
            reported[6].content(),
            RevisionContent::PriorTableProperties(properties) if properties.jc.is_some()
        ));
        assert!(matches!(
            reported[8].content(),
            RevisionContent::PriorParagraphProperties(properties) if properties.jc.is_some()
        ));
        assert!(matches!(
            reported[9].content(),
            RevisionContent::PriorSectionProperties(properties) if properties.title_pg == Some(true)
        ));

        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        for raw in revisions {
            assert!(
                output.contains(raw),
                "missing exact subtree {raw} in {output}"
            );
        }
        for raw in [
            r#"<w:rPrChange w:id="6" w:author="Fox"><w:rPr><w:i/></w:rPr></w:rPrChange>"#,
            r#"<w:tblPrChange w:id="7" w:author="Gia"><w:tblPr><w:jc w:val="center"/></w:tblPr></w:tblPrChange>"#,
            r#"<w:pPrChange w:id="9" w:author="Ivy"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>"#,
            r#"<w:sectPrChange w:id="10" w:author="Jay"><w:sectPr><w:titlePg/></w:sectPr></w:sectPrChange>"#,
        ] {
            assert!(
                output.contains(raw),
                "missing exact subtree {raw} in {output}"
            );
        }
        assert!(output.contains(r#"<x:ins xmlns:x="urn:not-word" x:id="99" x:author="No"/>"#));
        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(reopened.revisions().len(), 10);
    }

    #[test]
    fn hyperlink_and_nested_content_revisions_round_trip_and_report_in_order() {
        let nested = r#"<w:ins w:id="11" w:author="Ada"><w:del w:id="12" w:author="Ben"><w:r><w:delText>nested</w:delText></w:r></w:del></w:ins>"#;
        let hyperlink_before = r#"<w:hyperlink><w:ins w:id="21" w:author="Cy"><w:r><w:t>before</w:t></w:r></w:ins></w:hyperlink>"#;
        let direct_after =
            r#"<w:del w:id="22" w:author="Dee"><w:r><w:delText>after</w:delText></w:r></w:del>"#;
        let direct_before =
            r#"<w:ins w:id="23" w:author="Eve"><w:r><w:t>before</w:t></w:r></w:ins>"#;
        let hyperlink_after = r#"<w:hyperlink><w:del w:id="24" w:author="Fox"><w:r><w:delText>after</w:delText></w:r></w:del></w:hyperlink>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rId5"><w:r><w:t xml:space="preserve">before </w:t></w:r>{nested}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:hyperlink></w:p><w:p>{hyperlink_before}{direct_after}{direct_before}{hyperlink_after}</w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        assert_eq!(
            document
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            [11, 12, 21, 22, 23, 24]
        );
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        assert!(output.contains(nested), "missing exact subtree in {output}");
        for raw in [
            hyperlink_before,
            direct_after,
            direct_before,
            hyperlink_after,
        ] {
            assert!(
                output.contains(raw),
                "missing exact subtree {raw} in {output}"
            );
        }
        let positions = [21, 22, 23, 24].map(|id| {
            output
                .find(&format!(r#"w:id="{id}""#))
                .expect("revision id remains present")
        });
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(
            reopened
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            [11, 12, 21, 22, 23, 24]
        );
    }

    #[test]
    fn modeled_hyperlinks_preserve_unreported_raw_children_and_foreign_owners() {
        let malformed = r#"<w:ins w:id="bad"><w:r><w:t>raw revision</w:t></w:r></w:ins>"#;
        let foreign_child = r#"<x:opaque x:flag="1"/>"#;
        let foreign_hyperlink = r#"<x:hyperlink xmlns:x="urn:not-word"><w:r><w:t>foreign owner</w:t></w:r></x:hyperlink>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rId5" xmlns:x="urn:opaque"><w:r><w:t>before</w:t></w:r>{malformed}<w:ins w:id="31" w:author="Ada"><w:r><w:t>reported</w:t></w:r></w:ins>{foreign_child}<w:r><w:t>after</w:t></w:r></w:hyperlink>{foreign_hyperlink}</w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        assert_eq!(
            document
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            [31]
        );
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        assert!(
            output.contains(malformed),
            "missing malformed subtree in {output}"
        );
        assert!(
            output.contains(foreign_child),
            "missing foreign child in {output}"
        );
        assert!(
            output.contains(foreign_hyperlink),
            "foreign owner was promoted or lost in {output}"
        );
        let positions = [
            output.find("<w:t>before</w:t>").unwrap(),
            output.find(malformed).unwrap(),
            output.find(r#"w:id="31""#).unwrap(),
            output.find(foreign_child).unwrap(),
            output.find("<w:t>after</w:t>").unwrap(),
        ];
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(reopened.revisions().len(), 1);
        assert_eq!(reopened.revisions()[0].id(), 31);
    }

    #[test]
    fn aliased_hyperlink_runs_survive_a_locally_shadowed_word_prefix() {
        let foreign_run_properties = r#"<w:rPr/>"#;
        let foreign_text = r#"<w:t>foreign</w:t>"#;
        let foreign_drawing = r#"<w:drawing><w:opaque/></w:drawing>"#;
        let foreign_after = r#"<w:after/>"#;
        let lexical_literal = r#"<f:literal xmlns:f="urn:foreign" f:value="w: >">w:</f:literal>"#;
        let nested_word =
            r#"<f:box xmlns:f="urn:foreign" xmlns:w2="urn:two" f:value=">"><w:nested/></f:box>"#;
        let raw_at_boundary = r#"<w:boundary/>"#;
        let xml = format!(
            r#"<w:document xmlns:w="{W_NS}" xmlns:wx="{W_NS}"><w:body><w:p><wx:hyperlink xmlns:w="urn:foreign"><wx:r>{foreign_run_properties}<wx:rPr><wx:rStyle><wx:opaque/></wx:rStyle><w:pre/><wx:b wx:val="0"/><w:betweenDuplicate/><wx:b/><w:b/><w:between/><wx:i/><w:beforeChange/><wx:rPrChange wx:id="62" wx:author="Ada"><wx:rPr><wx:b/></wx:rPr></wx:rPrChange><w:afterChange/></wx:rPr>{foreign_text}<wx:t>before</wx:t>{foreign_drawing}<wx:t>middle</wx:t>{foreign_after}{lexical_literal}{nested_word}</wx:r><wx:ins wx:id="61" wx:author="Ada"><wx:r><wx:t>reported</wx:t></wx:r></wx:ins>{raw_at_boundary}<wx:r><wx:t>after</wx:t></wx:r></wx:hyperlink></w:p></w:body></w:document>"#
        );
        let document = CT_Document::from_xml(xml.as_bytes()).expect("document parses");

        assert_eq!(
            document.body.paragraphs().next().unwrap().text(),
            "beforemiddleafter"
        );
        assert_eq!(
            document
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            vec![62, 61]
        );
        let output =
            String::from_utf8(document.to_xml().expect("document writes")).expect("UTF-8 output");
        assert!(output.contains(r#"xmlns:w="urn:foreign""#));
        let foreign_run_properties = r#"<w:rPr xmlns:w="urn:foreign"/>"#;
        let foreign_property = r#"<w:b xmlns:w="urn:foreign"/>"#;
        let foreign_before_property = r#"<w:pre xmlns:w="urn:foreign"/>"#;
        let foreign_between_duplicate = r#"<w:betweenDuplicate xmlns:w="urn:foreign"/>"#;
        let foreign_between_property = r#"<w:between xmlns:w="urn:foreign"/>"#;
        let foreign_before_change = r#"<w:beforeChange xmlns:w="urn:foreign"/>"#;
        let foreign_after_change = r#"<w:afterChange xmlns:w="urn:foreign"/>"#;
        let foreign_text = r#"<w:t xmlns:w="urn:foreign">foreign</w:t>"#;
        let foreign_drawing = r#"<w:drawing xmlns:w="urn:foreign"><w:opaque/></w:drawing>"#;
        let foreign_after = r#"<w:after xmlns:w="urn:foreign"/>"#;
        for raw in [
            foreign_run_properties,
            foreign_before_property,
            foreign_between_duplicate,
            foreign_property,
            foreign_between_property,
            foreign_before_change,
            foreign_after_change,
            foreign_text,
            foreign_drawing,
            foreign_after,
        ] {
            assert_eq!(output.matches(raw).count(), 1, "raw child {raw}");
        }
        assert_eq!(output.matches(lexical_literal).count(), 1);
        assert!(!output.contains(r#"<f:literal xmlns:w="urn:foreign""#));
        let nested_word = r#"<f:box xmlns:f="urn:foreign" xmlns:w2="urn:two" f:value=">" xmlns:w="urn:foreign"><w:nested/></f:box>"#;
        assert_eq!(output.matches(nested_word).count(), 1);
        assert!(output.contains(raw_at_boundary));
        assert!(output.contains(&format!(r#"<w:r xmlns:w="{W_NS}">"#)));
        let positions = [
            output.find(foreign_run_properties).unwrap(),
            output.find("<wx:rStyle>").unwrap(),
            output.find(foreign_before_property).unwrap(),
            output.find(foreign_between_duplicate).unwrap(),
            output.find("<w:b/>").unwrap(),
            output.find(foreign_property).unwrap(),
            output.find(foreign_between_property).unwrap(),
            output.find("<w:i/>").unwrap(),
            output.find(foreign_before_change).unwrap(),
            output.find(foreign_after_change).unwrap(),
            output.find(r#"wx:id="62""#).unwrap(),
            output.find(foreign_text).unwrap(),
            output.find("<w:t>before</w:t>").unwrap(),
            output.find(foreign_drawing).unwrap(),
            output.find("<w:t>middle</w:t>").unwrap(),
            output.find(foreign_after).unwrap(),
            output.find(r#"wx:id="61""#).unwrap(),
            output.find(raw_at_boundary).unwrap(),
            output.find("<w:t>after</w:t>").unwrap(),
        ];
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let reopened = CT_Document::from_xml(output.as_bytes()).expect("output reparses");
        assert_eq!(
            reopened.body.paragraphs().next().unwrap().text(),
            "beforemiddleafter"
        );
        assert_eq!(
            reopened
                .revisions()
                .iter()
                .map(|revision| revision.id())
                .collect::<Vec<_>>(),
            vec![62, 61]
        );
    }
}
