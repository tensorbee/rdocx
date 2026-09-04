//! Pure evaluation of Word fields against an explicit document context.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::path::Path;

use oxml_core::custom_properties::CustomPropertyValue;
use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Body, CT_Document, CT_SectPr};
use rdocx_oxml::footnotes::{CT_Footnotes, NoteType};
use rdocx_oxml::header_footer::CT_HdrFtr;
use rdocx_oxml::namespace::{R_NS, W_NS, matches_local_name};
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::revision::{CT_Revision, RevisionContent, RevisionKind};
use rdocx_oxml::shared::ST_SectionType;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{
    CT_P, CT_R, Field, FieldArgument, FieldInstruction, RunContent, hyperlink_revision_index,
};

use crate::{Document, Error, Result, style};

/// A deterministic civil date and time supplied to field evaluation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FieldDateTime {
    pub year: i32,
    pub month: u8,
    pub day: u8,
    pub hour: u8,
    pub minute: u8,
    pub second: u8,
}

/// Explicit values that are not stored in the document package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FieldEvaluationContext {
    pub now: Option<FieldDateTime>,
    pub file_name: Option<String>,
    pub file_path: Option<String>,
    pub merge_fields: BTreeMap<String, String>,
    pub included_text: BTreeMap<String, String>,
    /// One-based source record number for mail-merge control fields.
    pub merge_record_number: Option<u32>,
    /// One-based output sequence number for mail-merge control fields.
    pub merge_sequence_number: Option<u32>,
}

/// The result of evaluating one field in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldEvaluation {
    pub field_index: usize,
    pub instruction: String,
    pub cached_result: String,
    pub outcome: FieldOutcome,
}

/// A pure field-evaluation decision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FieldOutcome {
    Resolved(String),
    DeferredPagination,
    TableOfContents(TocField),
    TableOfContentsEntry(TcField),
    MailMergeControl(MailMergeControl),
    Barcode(BarcodeField),
    KeepStored { diagnostic: String },
}

/// A validated table-of-contents rebuild request.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TocField {
    pub heading_levels: Option<(u8, u8)>,
    pub custom_styles: Vec<(String, u8)>,
    pub entries: TocEntrySelection,
    pub sequence_identifier: Option<String>,
    pub bookmark: Option<String>,
    pub hyperlink: bool,
    pub use_outline_levels: bool,
    pub omit_page_number_levels: Option<(u8, u8)>,
    pub page_number_separator: Option<String>,
    pub entry_page_separator: Option<String>,
}

/// Counts produced by one atomic table-of-contents rebuild.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TocRebuildReport {
    pub entry_count: usize,
    pub bookmark_count: usize,
    pub diagnostic_count: usize,
}

/// Which TC entries contribute to a table of contents.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TocEntrySelection {
    None,
    All,
    Identifier(String),
}

/// A validated table-of-contents entry request.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TcField {
    pub entry: String,
    pub level: u8,
    pub table_identifier: Option<String>,
    pub omit_page_number: bool,
}

/// A validated mail-merge control decision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MailMergeControl {
    NextRecord { record_number: u32 },
    NextRecordIf { condition: bool, record_number: u32 },
    SkipRecordIf { condition: bool, record_number: u32 },
    RecordNumber(u32),
    SequenceNumber(u32),
}

/// A validated generated-barcode request.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarcodeField {
    pub value: String,
    pub kind: BarcodeKind,
    pub height: Option<u32>,
    pub scale: Option<u16>,
    pub error_correction: Option<u8>,
    pub point_of_sale_style: Option<BarcodePointOfSaleStyle>,
    pub case_style: Option<BarcodeCaseStyle>,
    pub fix_check_digit: bool,
    pub rotation: Option<u8>,
    pub foreground_color: Option<u32>,
    pub background_color: Option<u32>,
    pub display_text: bool,
    pub add_start_stop: bool,
}

/// A barcode type accepted by Word's barcode field grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodeKind {
    Upca,
    Upce,
    Jan13,
    Jan8,
    Ean13,
    Ean8,
    Case,
    Itf14,
    Nw7,
    Code39,
    Code128,
    JpPost,
    Qr,
}

/// A point-of-sale style accepted by `\p`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodePointOfSaleStyle {
    Standard,
    SupplementalTwoDigit,
    SupplementalFiveDigit,
    Case,
}

/// An ITF14 case style accepted by `\c`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodeCaseStyle {
    Standard,
    Extended,
    Add,
}

impl Document {
    /// Evaluate every typed field without mutating stored results.
    pub fn evaluate_fields(
        &self,
        context: &FieldEvaluationContext,
    ) -> Result<Vec<FieldEvaluation>> {
        self.evaluate_fields_with_policy(context, false)
    }

    fn evaluate_fields_with_policy(
        &self,
        context: &FieldEvaluationContext,
        missing_merge_fields_as_empty: bool,
    ) -> Result<Vec<FieldEvaluation>> {
        let mut evaluator = if missing_merge_fields_as_empty {
            Evaluator::for_mail_merge(self, context)
        } else {
            Evaluator::new(self, context)
        };

        let mut main = Vec::new();
        collect_body_paragraphs(&self.document.body, &mut main);
        evaluator.evaluate_story("main", &main);

        for (part_name, xml) in referenced_header_footer_parts(self, true) {
            if let Ok(part) = CT_HdrFtr::from_xml(&xml) {
                let paragraphs = part.paragraphs.iter().collect::<Vec<_>>();
                evaluator.evaluate_story(&format!("header:{part_name}"), &paragraphs);
            }
        }
        for (part_name, xml) in referenced_header_footer_parts(self, false) {
            if let Ok(part) = CT_HdrFtr::from_xml(&xml) {
                let paragraphs = part.paragraphs.iter().collect::<Vec<_>>();
                evaluator.evaluate_story(&format!("footer:{part_name}"), &paragraphs);
            }
        }

        let footnotes = normal_note_paragraphs(&self.footnotes);
        evaluator.evaluate_story("footnotes", &footnotes);

        for (_, xml) in relationship_parts(self, rel_types::ENDNOTES) {
            if let Ok(part) = CT_Footnotes::from_xml(&xml) {
                let endnotes = normal_note_paragraphs(&part);
                evaluator.evaluate_story("endnotes", &endnotes);
            }
        }

        Ok(evaluator.results)
    }

    /// Evaluate and materialize every typed field cache in document order.
    pub fn update_fields(&mut self, context: &FieldEvaluationContext) -> Result<usize> {
        self.update_fields_with_policy(context, false)
    }

    /// Rebuild every supported table of contents already present in the document.
    ///
    /// The operation stages bookmarks, cached entries, and deterministic page
    /// targets on an independent document. Any malformed or ambiguous source
    /// leaves the receiver unchanged.
    pub fn rebuild_toc(&mut self) -> Result<TocRebuildReport> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let document_xml = candidate
            .package
            .get_part(&candidate.doc_part_name)
            .ok_or(Error::NoDocumentPart)?
            .to_vec();
        let toc_spans = scan_dynamic_toc_spans(&document_xml)?;
        let simple_diagnostic_count = count_simple_toc_fields(&document_xml)?;
        if toc_spans.is_empty() {
            return Ok(TocRebuildReport {
                diagnostic_count: simple_diagnostic_count,
                ..Default::default()
            });
        }

        let (toc_fields, mut diagnostic_count) =
            parse_dynamic_toc_fields(&candidate, &document_xml, &toc_spans)?;
        diagnostic_count += simple_diagnostic_count;
        if toc_fields.iter().all(Option::is_none) {
            return Ok(TocRebuildReport {
                diagnostic_count,
                ..Default::default()
            });
        }
        let bookmark_state = inspect_toc_bookmarks(&candidate.document.body, &document_xml)?;
        let sources = discover_toc_sources(&candidate, &toc_spans, &toc_fields, &bookmark_state)?;
        let rebuilt_toc_spans = toc_spans
            .iter()
            .zip(&toc_fields)
            .filter_map(|(span, field)| field.is_some().then_some(span.clone()))
            .collect::<Vec<_>>();
        let bookmark_repairs =
            toc_crossing_bookmark_repairs(&bookmark_state, &toc_spans, &toc_fields);
        let mut allocator = TocBookmarkAllocator::new(bookmark_state);
        let mut bookmark_by_paragraph = BTreeMap::<usize, TocBookmark>::new();
        for source in sources.iter().flatten() {
            if !source.needs_bookmark {
                continue;
            }
            if bookmark_by_paragraph.contains_key(&source.paragraph_index) {
                continue;
            }
            let is_partial_boundary_source = rebuilt_toc_spans.iter().any(|span| {
                source.paragraph_index == span.begin_paragraph
                    || source.paragraph_index == span.end_paragraph
            });
            if !is_partial_boundary_source
                && let Some(existing) = allocator.whole_paragraph_name(source.paragraph_index)
            {
                bookmark_by_paragraph.insert(
                    source.paragraph_index,
                    TocBookmark {
                        id: existing.0,
                        name: existing.1,
                        insert: false,
                    },
                );
            } else {
                let allocated = allocator.allocate()?;
                bookmark_by_paragraph.insert(
                    source.paragraph_index,
                    TocBookmark {
                        id: allocated.0,
                        name: allocated.1,
                        insert: true,
                    },
                );
            }
        }
        let bookmark_count = bookmark_by_paragraph
            .values()
            .filter(|bookmark| bookmark.insert)
            .count();
        let provisional_xml = insert_toc_bookmarks_xml(
            &document_xml,
            &bookmark_by_paragraph,
            &bookmark_repairs,
            &toc_spans,
            &rebuilt_toc_spans,
        )?;
        CT_Document::from_xml(&provisional_xml)?;
        candidate
            .package
            .set_part(&candidate.doc_part_name, provisional_xml.clone());
        candidate = reopen_staged_document(candidate)?;
        let provisional_spans = scan_dynamic_toc_spans(&provisional_xml)?;
        if provisional_spans.len() != toc_fields.len() {
            return Err(Error::Other(
                "table of contents ownership changed while staging bookmarks".to_owned(),
            ));
        }
        let end_bookmark_starts = provisional_spans
            .iter()
            .enumerate()
            .map(|(toc_index, span)| {
                end_boundary_bookmark_starts(
                    toc_index,
                    span,
                    &provisional_spans,
                    &toc_fields,
                    &bookmark_by_paragraph,
                    &bookmark_repairs,
                )
            })
            .collect::<Vec<_>>();
        let mut placeholders = Vec::new();
        let mut edits = Vec::new();
        let mut entry_count = 0usize;
        for (toc_index, ((span, toc), toc_sources)) in provisional_spans
            .iter()
            .zip(&toc_fields)
            .zip(&sources)
            .enumerate()
        {
            let Some(toc) = toc else {
                continue;
            };
            let generated = render_toc_entries(
                toc_index,
                toc,
                toc_sources,
                &bookmark_by_paragraph,
                &provisional_xml,
                &mut placeholders,
            )?;
            entry_count += toc_sources.len();
            edits.push(FieldSourceEdit {
                start: span.result_start,
                end: span.result_end,
                replacement: dynamic_toc_replacement(
                    &provisional_xml,
                    span,
                    &generated,
                    &end_bookmark_starts[toc_index],
                )?,
            });
        }
        let mut staged_xml = provisional_xml;
        for edit in edits.into_iter().rev() {
            staged_xml.splice(edit.start..edit.end, edit.replacement);
        }
        CT_Document::from_xml(&staged_xml)?;
        candidate
            .package
            .set_part(&candidate.doc_part_name, staged_xml);
        let mut provisional = reopen_staged_document(candidate)?;

        let page_values = deterministic_toc_page_values(&provisional)?;
        let current_xml = provisional
            .package
            .get_part(&provisional.doc_part_name)
            .ok_or(Error::NoDocumentPart)?;
        let mut final_xml = current_xml.to_vec();
        let final_spans = scan_dynamic_toc_spans(&final_xml)?;
        if final_spans.len() != toc_fields.len() {
            return Err(Error::Other(
                "table of contents ownership changed during page substitution".to_owned(),
            ));
        }
        let mut page_edits = Vec::with_capacity(placeholders.len());
        for placeholder in &placeholders {
            let target = page_values.get(&placeholder.bookmark).ok_or_else(|| {
                Error::Other(format!(
                    "table of contents page target {} was not resolved",
                    placeholder.bookmark
                ))
            })?;
            let span = &final_spans[placeholder.toc_index];
            let owned = &final_xml[span.result_start..span.result_end];
            let matches = byte_match_offsets(owned, placeholder.token.as_bytes());
            if matches.len() != 1 {
                return Err(Error::Other(format!(
                    "table of contents page placeholder {} was not uniquely owned",
                    placeholder.token
                )));
            }
            let start = span.result_start + matches[0];
            page_edits.push(FieldSourceEdit {
                start,
                end: start + placeholder.token.len(),
                replacement: target.as_bytes().to_vec(),
            });
        }
        page_edits.sort_by_key(|edit| edit.start);
        for edit in page_edits.into_iter().rev() {
            final_xml.splice(edit.start..edit.end, edit.replacement);
        }
        let final_spans = scan_dynamic_toc_spans(&final_xml)?;
        final_xml =
            relocate_end_boundary_bookmark_starts(final_xml, &final_spans, &end_bookmark_starts)?;
        CT_Document::from_xml(&final_xml)?;
        provisional
            .package
            .set_part(&provisional.doc_part_name, final_xml);
        let mut completed = reopen_staged_document(provisional)?;
        completed.invalidate_layout();
        self.commit_staged_mutation(completed);

        Ok(TocRebuildReport {
            entry_count,
            bookmark_count,
            diagnostic_count,
        })
    }

    fn update_fields_with_policy(
        &mut self,
        context: &FieldEvaluationContext,
        missing_merge_fields_as_empty: bool,
    ) -> Result<usize> {
        let evaluations =
            self.evaluate_fields_with_policy(context, missing_merge_fields_as_empty)?;
        let updates = evaluations
            .iter()
            .map(|evaluation| match &evaluation.outcome {
                FieldOutcome::Resolved(value) => CachedFieldUpdate {
                    cached_result: value.clone(),
                    dirty: false,
                },
                FieldOutcome::DeferredPagination
                | FieldOutcome::TableOfContents(_)
                | FieldOutcome::TableOfContentsEntry(_)
                | FieldOutcome::MailMergeControl(_)
                | FieldOutcome::Barcode(_)
                | FieldOutcome::KeepStored { .. } => CachedFieldUpdate {
                    cached_result: evaluation.cached_result.clone(),
                    dirty: true,
                },
            })
            .collect::<Vec<_>>();
        if updates
            .iter()
            .any(|update| !update.cached_result.chars().all(valid_xml_character))
        {
            return Err(Error::Other(
                "field result contains a character forbidden by XML 1.0".to_owned(),
            ));
        }
        if updates.is_empty() {
            return Ok(0);
        }

        let mut document = self.document.clone();
        let mut footnotes = self.footnotes.clone();
        let mut staged_parts = Vec::new();
        let mut update_index = 0usize;

        apply_updates_to_body(&mut document.body, &updates, &mut update_index);

        for (part_name, xml) in referenced_header_footer_parts(self, true) {
            if let Ok(mut part) = CT_HdrFtr::from_xml(&xml) {
                let part_start = update_index;
                apply_updates_to_paragraphs(&mut part.paragraphs, &updates, &mut update_index);
                if update_index > part_start {
                    let paragraphs = part.paragraphs.iter().collect::<Vec<_>>();
                    let updated = patch_story_field_sources(
                        &xml,
                        &paragraphs,
                        PackageStoryKind::HeaderFooter,
                    )?;
                    CT_HdrFtr::from_xml(&updated)?;
                    staged_parts.push((part_name, updated));
                }
            }
        }
        for (part_name, xml) in referenced_header_footer_parts(self, false) {
            if let Ok(mut part) = CT_HdrFtr::from_xml(&xml) {
                let part_start = update_index;
                apply_updates_to_paragraphs(&mut part.paragraphs, &updates, &mut update_index);
                if update_index > part_start {
                    let paragraphs = part.paragraphs.iter().collect::<Vec<_>>();
                    let updated = patch_story_field_sources(
                        &xml,
                        &paragraphs,
                        PackageStoryKind::HeaderFooter,
                    )?;
                    CT_HdrFtr::from_xml(&updated)?;
                    staged_parts.push((part_name, updated));
                }
            }
        }

        let footnotes_start = update_index;
        apply_updates_to_notes(&mut footnotes, &updates, &mut update_index);
        let mut footnotes_dirty = self.footnotes_dirty;
        if update_index > footnotes_start && !self.footnotes_dirty {
            if let Some((part_name, xml)) = relationship_parts(self, rel_types::FOOTNOTES)
                .into_iter()
                .next()
            {
                let paragraphs = normal_note_paragraphs(&footnotes);
                let updated =
                    patch_story_field_sources(&xml, &paragraphs, PackageStoryKind::Footnotes)?;
                CT_Footnotes::from_xml(&updated)?;
                staged_parts.push((part_name, updated));
            } else {
                footnotes_dirty = true;
            }
        }

        for (part_name, xml) in relationship_parts(self, rel_types::ENDNOTES) {
            if let Ok(mut part) = CT_Footnotes::from_xml(&xml) {
                let part_start = update_index;
                apply_updates_to_notes(&mut part, &updates, &mut update_index);
                if update_index > part_start {
                    let paragraphs = normal_note_paragraphs(&part);
                    let updated =
                        patch_story_field_sources(&xml, &paragraphs, PackageStoryKind::Endnotes)?;
                    CT_Footnotes::from_xml(&updated)?;
                    staged_parts.push((part_name, updated));
                }
            }
        }

        if update_index != updates.len() {
            return Err(Error::Other(format!(
                "field update traversal consumed {update_index} of {} staged evaluations",
                updates.len()
            )));
        }

        let document_xml = document.to_xml()?;
        rdocx_oxml::document::CT_Document::from_xml(&document_xml)?;
        if footnotes_dirty && !footnotes.footnotes.is_empty() {
            let footnotes_xml = footnotes.to_xml_footnotes()?;
            CT_Footnotes::from_xml(&footnotes_xml)?;
        }

        self.document = document;
        self.footnotes = footnotes;
        self.footnotes_dirty = footnotes_dirty;
        for (part_name, xml) in staged_parts {
            self.package.set_part(&part_name, xml);
        }
        self.invalidate_layout();
        Ok(updates.len())
    }

    /// Materialize one independent document for each flat mail-merge record.
    pub fn mail_merge(&self, records: &[BTreeMap<String, String>]) -> Result<Vec<Document>> {
        if records.is_empty() {
            return Err(Error::Other(
                "mail merge requires at least one record".to_owned(),
            ));
        }

        let mut outputs = Vec::with_capacity(records.len());
        for (record_index, record) in records.iter().enumerate() {
            let record_number = u32::try_from(record_index)
                .ok()
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| Error::Other("mail merge record count exceeds u32".to_owned()))?;
            let context = FieldEvaluationContext {
                merge_fields: record.clone(),
                merge_record_number: Some(record_number),
                merge_sequence_number: Some(record_number),
                ..Default::default()
            };
            let mut candidate = self.clone_for_staging();
            candidate.update_fields_with_policy(&context, true)?;
            let bytes = candidate.to_bytes()?;
            outputs.push(Document::from_bytes(&bytes)?);
        }
        Ok(outputs)
    }

    /// Materialize one document whose record bodies form next-page sections.
    pub fn mail_merge_sections(&self, records: &[BTreeMap<String, String>]) -> Result<Document> {
        reject_varying_non_body_merge_fields(self, records)?;
        let mut candidates = self.mail_merge(records)?;

        let mut identity_state = BodyIdentityState::from_documents(&candidates)?;
        for candidate in candidates.iter_mut().skip(1) {
            remap_body_identities(candidate, &mut identity_state)?;
        }

        let bodies = candidates
            .iter()
            .map(|candidate| candidate.document.body.clone())
            .collect::<Vec<_>>();
        let mut combined = candidates.remove(0);
        combined.document.body.content.clear();
        combined.document.body.sect_pr = None;

        let final_index = bodies.len() - 1;
        for (index, mut body) in bodies.into_iter().enumerate() {
            combined.document.body.content.append(&mut body.content);
            if index == final_index {
                combined.document.body.sect_pr = body.sect_pr;
            } else {
                let mut section = body.sect_pr.unwrap_or_else(empty_section_properties);
                section.section_type = Some(ST_SectionType::NextPage);
                let mut paragraph = CT_P::new();
                paragraph.properties = Some(CT_PPr {
                    sect_pr: Some(section),
                    ..Default::default()
                });
                combined
                    .document
                    .body
                    .content
                    .push(BodyContent::Paragraph(paragraph));
            }
        }
        combined.invalidate_layout();

        let bytes = combined.to_bytes()?;
        Document::from_bytes(&bytes)
    }

    /// Update typed field caches, then save the package to a file path.
    pub fn save_with_field_updates<P: AsRef<Path>>(
        &mut self,
        path: P,
        context: &FieldEvaluationContext,
    ) -> Result<()> {
        self.update_fields(context)?;
        self.save(path)
    }

    /// Update typed field caches, then save the package to bytes.
    pub fn to_bytes_with_field_updates(
        &mut self,
        context: &FieldEvaluationContext,
    ) -> Result<Vec<u8>> {
        self.update_fields(context)?;
        self.to_bytes()
    }
}

#[derive(Debug, Clone)]
struct DynamicTocSpan {
    instruction: String,
    field_start: usize,
    field_end: usize,
    begin_paragraph: usize,
    end_paragraph: usize,
    begin_run_start: usize,
    instruction_paragraph_start: usize,
    result_start: usize,
    result_end: usize,
    result_start_position: TocRunPosition,
    result_end_position: TocRunPosition,
    end_run_end: usize,
    start_paragraph_name: String,
    separator_wrapper_names: Vec<String>,
    instruction_runs: Vec<DynamicInstructionRun>,
    end_paragraph_start: usize,
    end_paragraph_content_start: usize,
    end_wrapper_prefixes: Vec<(usize, usize)>,
}

#[derive(Debug)]
struct DynamicFieldScan {
    instruction: String,
    field_start: usize,
    begin_paragraph: usize,
    begin_paragraph_start: usize,
    begin_run_start: usize,
    separator_paragraph: Option<usize>,
    separator_run_start: Option<usize>,
    result_start: Option<usize>,
    result_start_position: Option<TocRunPosition>,
    start_paragraph_name: Option<String>,
    separator_wrapper_names: Vec<String>,
    instruction_runs: Vec<DynamicInstructionRun>,
}

#[derive(Debug, Clone)]
struct DynamicInstructionRun {
    start: usize,
    end: usize,
    inherited_namespaces: BTreeMap<String, String>,
}

#[derive(Debug)]
struct DynamicXmlElement {
    local_name: Vec<u8>,
    qualified_name: String,
    typed_block_owner: Option<TypedBlockOwner>,
    is_word: bool,
    is_typed_paragraph: bool,
    is_typed_inline_owner: bool,
    revision_depth: usize,
    sdt_content_seen: bool,
    namespace_bindings: BTreeMap<String, String>,
    inherited_namespaces: BTreeMap<String, String>,
    run_position: Option<TocRunPosition>,
    hyperlink_plan: Option<DynamicHyperlinkPlan>,
    start: usize,
    start_tag_end: usize,
    paragraph: Option<usize>,
}

#[derive(Debug)]
struct DynamicHyperlinkPlan {
    revision_orders: Vec<TocRawOrder>,
    next_revision: usize,
    preserved_raw: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TypedBlockOwner {
    Document,
    Body,
    Table,
    Row,
    Cell,
    ContentControl(BlockControlOwner),
    Content(BlockControlOwner),
    Paragraph,
}

const MAX_DYNAMIC_REVISION_NESTING_DEPTH: usize = 32;

fn dynamic_namespace_bindings(
    element: &BytesStart<'_>,
    elements: &[DynamicXmlElement],
) -> Result<(BTreeMap<String, String>, BTreeMap<String, String>)> {
    let inherited = elements
        .last()
        .map_or_else(BTreeMap::new, |parent| parent.namespace_bindings.clone());
    let mut bindings = inherited.clone();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Other(format!("invalid XML namespace attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            Some(String::new())
        } else {
            key.strip_prefix(b"xmlns:")
                .map(|prefix| String::from_utf8_lossy(prefix).into_owned())
        };
        if let Some(prefix) = prefix {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .map_err(|error| Error::Other(format!("invalid XML namespace attribute: {error}")))?
                .into_owned();
            if value.is_empty() {
                bindings.remove(&prefix);
            } else {
                bindings.insert(prefix, value);
            }
        }
    }
    Ok((bindings, inherited))
}

#[derive(Clone, Copy)]
enum DynamicContentControlOwner {
    Block(BlockControlOwner),
    Inline,
}

fn dynamic_element_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut reader = quick_xml::Reader::from_reader(&xml[start..]);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("invalid table of contents XML element: {error}"))
        })? {
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => {
                return Ok(start + reader.buffer_position() as usize);
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Other("table of contents XML element has an unmatched end".to_owned())
                })?;
                if depth == 0 {
                    return Ok(start + reader.buffer_position() as usize);
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "table of contents XML element is unclosed".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn xml_fragment_with_namespaces(
    raw: &[u8],
    bindings: &BTreeMap<String, String>,
    description: &str,
) -> Result<Vec<u8>> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (insertion, local_namespaces) = match reader
        .read_event_into(&mut buffer)
        .map_err(|error| Error::Other(format!("invalid {description}: {error}")))?
    {
        Event::Start(start) | Event::Empty(start) => {
            let mut local_namespaces = HashSet::new();
            for attribute in start.attributes() {
                let attribute = attribute
                    .map_err(|error| Error::Other(format!("invalid {description}: {error}")))?;
                let key = attribute.key.as_ref();
                if key == b"xmlns" {
                    local_namespaces.insert(String::new());
                } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                    local_namespaces.insert(String::from_utf8_lossy(prefix).into_owned());
                }
            }
            let tag_end = reader.buffer_position() as usize;
            let insertion = if tag_end >= 2 && raw[tag_end - 2] == b'/' {
                tag_end - 2
            } else {
                tag_end - 1
            };
            (insertion, local_namespaces)
        }
        _ => return Err(Error::Other(format!("{description} has no start tag"))),
    };
    let mut output = Vec::with_capacity(raw.len() + bindings.len() * 32);
    output.extend_from_slice(&raw[..insertion]);
    for (prefix, namespace) in bindings {
        if prefix == "xml" || local_namespaces.contains(prefix) {
            continue;
        }
        if prefix.is_empty() {
            output.extend_from_slice(b" xmlns=\"");
        } else {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
        }
        output.extend_from_slice(xml_escape_attribute(namespace).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(&raw[insertion..]);
    Ok(output)
}

fn dynamic_content_control_is_typed(
    xml: &[u8],
    start: usize,
    bindings: &BTreeMap<String, String>,
    owner: DynamicContentControlOwner,
) -> Result<bool> {
    let end = dynamic_element_end(xml, start)?;
    let control = xml_fragment_with_namespaces(
        &xml[start..end],
        bindings,
        "table of contents content control",
    )?;
    let mut candidate = Vec::with_capacity(control.len() + 256);
    candidate.extend_from_slice(format!(r#"<w:document xmlns:w="{W_NS}"><w:body>"#).as_bytes());
    let (prefix, suffix): (&[u8], &[u8]) = match owner {
        DynamicContentControlOwner::Block(BlockControlOwner::Body) => (b"", b""),
        DynamicContentControlOwner::Block(BlockControlOwner::Table) => {
            (b"<w:tbl>", b"<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>")
        }
        DynamicContentControlOwner::Block(BlockControlOwner::Row) => {
            (b"<w:tbl><w:tr>", b"<w:tc><w:p/></w:tc></w:tr></w:tbl>")
        }
        DynamicContentControlOwner::Block(BlockControlOwner::Cell) => {
            (b"<w:tbl><w:tr><w:tc>", b"<w:p/></w:tc></w:tr></w:tbl>")
        }
        DynamicContentControlOwner::Inline => (b"<w:p>", b"</w:p>"),
    };
    candidate.extend_from_slice(prefix);
    candidate.extend_from_slice(&control);
    candidate.extend_from_slice(suffix);
    candidate.extend_from_slice(b"</w:body></w:document>");
    let Ok(document) = CT_Document::from_xml(&candidate) else {
        return Ok(false);
    };
    let accepted = match owner {
        DynamicContentControlOwner::Block(BlockControlOwner::Body) => matches!(
            document.body.content.first(),
            Some(BodyContent::ContentControl(_))
        ),
        DynamicContentControlOwner::Block(BlockControlOwner::Table) => document
            .body
            .tables()
            .next()
            .is_some_and(|table| table.content_controls.len() == 1),
        DynamicContentControlOwner::Block(BlockControlOwner::Row) => document
            .body
            .tables()
            .next()
            .and_then(|table| table.rows.first())
            .is_some_and(|row| row.content_controls.len() == 1),
        DynamicContentControlOwner::Block(BlockControlOwner::Cell) => document
            .body
            .tables()
            .next()
            .and_then(|table| table.rows.first())
            .and_then(|row| row.cells.first())
            .is_some_and(|cell| {
                matches!(cell.content.first(), Some(CellContent::ContentControl(_)))
            }),
        DynamicContentControlOwner::Inline => document
            .body
            .paragraphs()
            .next()
            .is_some_and(|paragraph| paragraph.content_controls.len() == 1),
    };
    Ok(accepted)
}

fn validate_dynamic_content_control(
    xml: &[u8],
    start: usize,
    word: bool,
    local: &[u8],
    bindings: &BTreeMap<String, String>,
    typed_block_owner: &mut Option<TypedBlockOwner>,
    is_typed_inline_owner: &mut bool,
) -> Result<()> {
    if !word || local != b"sdt" {
        return Ok(());
    }
    let owner = match *typed_block_owner {
        Some(TypedBlockOwner::ContentControl(owner)) => {
            Some(DynamicContentControlOwner::Block(owner))
        }
        _ if *is_typed_inline_owner => Some(DynamicContentControlOwner::Inline),
        _ => None,
    };
    if let Some(owner) = owner
        && !dynamic_content_control_is_typed(xml, start, bindings, owner)?
    {
        *typed_block_owner = None;
        *is_typed_inline_owner = false;
    }
    Ok(())
}

fn is_dynamic_revision_element(local: &[u8]) -> bool {
    matches!(
        local,
        b"ins"
            | b"del"
            | b"moveFrom"
            | b"moveTo"
            | b"rPrChange"
            | b"pPrChange"
            | b"tblPrChange"
            | b"sectPrChange"
    )
}

fn invalidate_overdeep_revision_owner(
    word: bool,
    local: &[u8],
    elements: &mut [DynamicXmlElement],
) -> Option<usize> {
    if !word
        || !is_dynamic_revision_element(local)
        || elements
            .iter()
            .filter(|element| element.is_word && is_dynamic_revision_element(&element.local_name))
            .count()
            < MAX_DYNAMIC_REVISION_NESTING_DEPTH
    {
        return None;
    }
    let owner = elements.iter().position(|element| {
        element.is_typed_inline_owner && matches!(element.local_name.as_slice(), b"ins" | b"moveTo")
    })?;
    let start = elements[owner].start;
    for element in &mut elements[owner..] {
        element.is_typed_inline_owner = false;
    }
    Some(start)
}

fn mark_typed_sdt_content(
    elements: &mut [DynamicXmlElement],
    local: &[u8],
    typed_block_owner: Option<TypedBlockOwner>,
    is_typed_inline_owner: bool,
) {
    if local == b"sdtContent"
        && (matches!(typed_block_owner, Some(TypedBlockOwner::Content(_))) || is_typed_inline_owner)
        && let Some(parent) = elements.last_mut()
    {
        parent.sdt_content_seen = true;
    }
}

fn scan_dynamic_toc_spans(xml: &[u8]) -> Result<Vec<DynamicTocSpan>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut fields = Vec::<DynamicFieldScan>::new();
    let mut spans = Vec::<DynamicTocSpan>::new();
    let mut paragraph_count = 0usize;
    let mut paragraph_run_boundaries = Vec::<usize>::new();
    let mut paragraph_nested_run_orders = Vec::<HashMap<(usize, TocRawOrder), usize>>::new();
    let mut paragraph_raw_before = Vec::<usize>::new();
    let mut instruction_depth = None;

    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    fields.retain(|field| field.field_start < start);
                    spans.retain(|span| span.field_start < start);
                    instruction_depth = None;
                }
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    let index = paragraph_count;
                    paragraph_count += 1;
                    paragraph_run_boundaries.push(0);
                    paragraph_nested_run_orders.push(HashMap::new());
                    paragraph_raw_before.push(0);
                    Some(index)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let mut is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let revision_depth = if is_typed_inline_owner {
                    elements.last().map_or(0, |parent| parent.revision_depth)
                        + usize::from(matches!(local.as_slice(), b"ins" | b"moveTo"))
                } else {
                    0
                };
                let modeled_simple_field = if word && local == b"fldSimple" {
                    resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        !Field::new(&instruction, "").instruction.name.is_empty()
                    })
                } else {
                    false
                };
                let run_position = dynamic_toc_run_position(
                    word,
                    &local,
                    is_typed_inline_owner,
                    modeled_simple_field,
                    &mut elements,
                    paragraph,
                    &mut paragraph_run_boundaries,
                    &mut paragraph_nested_run_orders,
                    &mut paragraph_raw_before,
                );
                let direct_paragraph_child = direct_typed_paragraph_parent(&elements, paragraph);
                let hyperlink_plan = if word && local == b"hyperlink" && direct_paragraph_child {
                    Some(dynamic_hyperlink_plan(
                        xml,
                        before,
                        &namespace_bindings,
                        paragraph
                            .and_then(|index| paragraph_raw_before.get(index).copied())
                            .ok_or_else(|| {
                                Error::Other(
                                    "table of contents hyperlink has no paragraph position"
                                        .to_owned(),
                                )
                            })?,
                    )?)
                } else {
                    None
                };
                advance_direct_paragraph_raw_child(
                    word,
                    &local,
                    is_typed_inline_owner,
                    modeled_simple_field,
                    direct_paragraph_child,
                    paragraph,
                    hyperlink_plan.as_ref(),
                    &mut paragraph_raw_before,
                );
                if word && local == b"fldChar" && direct_word_run_parent(&elements, paragraph) {
                    update_dynamic_field_stack(
                        &element,
                        reader.resolver(),
                        before,
                        paragraph,
                        &elements,
                        &mut fields,
                        &mut spans,
                    )?;
                }
                if word
                    && local == b"instrText"
                    && !fields.is_empty()
                    && direct_word_run_parent(&elements, paragraph)
                {
                    instruction_depth = Some(elements.len());
                }
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner,
                    revision_depth,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position,
                    hyperlink_plan,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    fields.retain(|field| field.field_start < start);
                    spans.retain(|span| span.field_start < start);
                    instruction_depth = None;
                }
                let paragraph = elements.last().and_then(|element| element.paragraph);
                let typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                let direct_paragraph_child = direct_typed_paragraph_parent(&elements, paragraph);
                advance_direct_paragraph_raw_child(
                    word,
                    &local,
                    false,
                    false,
                    direct_paragraph_child,
                    paragraph,
                    None,
                    &mut paragraph_raw_before,
                );
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                if word
                    && matches_local_name(element.name().as_ref(), b"fldChar")
                    && direct_word_run_parent(&elements, paragraph)
                {
                    update_dynamic_field_stack(
                        &element,
                        reader.resolver(),
                        before,
                        paragraph,
                        &elements,
                        &mut fields,
                        &mut spans,
                    )?;
                }
            }
            Event::Text(text) if instruction_depth.is_some() => {
                if let Some(field) = fields.last_mut()
                    && field.separator_paragraph.is_none()
                {
                    let decoded = text.decode().map_err(|error| {
                        Error::Other(format!("invalid table of contents instruction: {error}"))
                    })?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::Other(format!("invalid table of contents instruction: {error}"))
                    })?;
                    field.instruction.push_str(&unescaped);
                }
            }
            Event::CData(text) if instruction_depth.is_some() => {
                if let Some(field) = fields.last_mut()
                    && field.separator_paragraph.is_none()
                {
                    field.instruction.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid table of contents instruction: {error}"))
                    })?);
                }
            }
            Event::Comment(_) | Event::PI(_) => {
                if let Some(paragraph) = elements.last().and_then(|element| {
                    element
                        .is_typed_paragraph
                        .then_some(element.paragraph)
                        .flatten()
                }) && let Some(raw_before) = paragraph_raw_before.get_mut(paragraph)
                {
                    *raw_before += 1;
                }
            }
            Event::End(element) => {
                let Some(closed) = elements.pop() else {
                    return Err(Error::Other(
                        "table of contents XML has an unmatched end element".to_owned(),
                    ));
                };
                if closed
                    .hyperlink_plan
                    .as_ref()
                    .is_some_and(|plan| plan.preserved_raw)
                    && let Some(paragraph) = closed.paragraph
                    && let Some(raw_before) = paragraph_raw_before.get_mut(paragraph)
                {
                    *raw_before += 1;
                }
                if word && matches_local_name(element.name().as_ref(), b"instrText") {
                    instruction_depth = None;
                }
                if word && matches_local_name(element.name().as_ref(), b"r") {
                    for field in &mut fields {
                        if closed.is_typed_inline_owner && field.result_start.is_none() {
                            field.instruction_runs.push(DynamicInstructionRun {
                                start: closed.start,
                                end: after,
                                inherited_namespaces: closed.inherited_namespaces.clone(),
                            });
                        }
                        if field.separator_run_start == Some(closed.start)
                            && field.result_start.is_none()
                        {
                            field.result_start = Some(after);
                        }
                    }
                    if let Some(span) = spans
                        .iter_mut()
                        .rev()
                        .find(|span| span.result_end == closed.start && span.end_run_end == 0)
                    {
                        span.end_run_end = after;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !elements.is_empty() {
        return Err(Error::Other(
            "table of contents XML has an unclosed element".to_owned(),
        ));
    }
    if fields
        .iter()
        .any(|field| Field::new(&field.instruction, "").instruction.name == "TOC")
    {
        return Err(Error::Other(
            "table of contents field is missing its end marker".to_owned(),
        ));
    }
    spans.sort_by_key(|span| span.field_start);
    if spans
        .windows(2)
        .any(|pair| pair[1].field_start < pair[0].field_end)
    {
        return Err(Error::Other(
            "nested table of contents fields have ambiguous ownership".to_owned(),
        ));
    }
    let paragraphs = toc_paragraph_insertions(xml)?;
    for span in &mut spans {
        span.end_paragraph_content_start = paragraphs
            .get(span.end_paragraph)
            .map(|paragraph| paragraph.content_start)
            .ok_or_else(|| {
                Error::Other(
                    "table of contents end paragraph was not found in package XML".to_owned(),
                )
            })?;
    }
    Ok(spans)
}

fn count_simple_toc_fields(xml: &[u8]) -> Result<usize> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut simple_toc_starts = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    simple_toc_starts.retain(|field_start| *field_start < start);
                }
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    Some(0)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let mut is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let revision_depth = if is_typed_inline_owner {
                    elements.last().map_or(0, |parent| parent.revision_depth)
                        + usize::from(matches!(local.as_slice(), b"ins" | b"moveTo"))
                } else {
                    0
                };
                if word
                    && local == b"fldSimple"
                    && paragraph.is_some()
                    && accepted_simple_toc_parent(&elements, paragraph)
                    && resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        Field::new(&instruction, "").instruction.name == "TOC"
                    })
                {
                    simple_toc_starts.push(before);
                }
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner,
                    revision_depth,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                if let Some(start) = invalidate_overdeep_revision_owner(word, &local, &mut elements)
                {
                    simple_toc_starts.retain(|field_start| *field_start < start);
                }
                let paragraph = elements.last().and_then(|element| element.paragraph);
                let typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                if word
                    && matches_local_name(element.name().as_ref(), b"fldSimple")
                    && elements
                        .last()
                        .and_then(|element| element.paragraph)
                        .is_some()
                    && accepted_simple_toc_parent(
                        &elements,
                        elements.last().and_then(|element| element.paragraph),
                    )
                    && resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"instr",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, instruction)| {
                        Field::new(&instruction, "").instruction.name == "TOC"
                    })
                {
                    simple_toc_starts.push(before);
                }
            }
            Event::End(_) => {
                elements.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(simple_toc_starts.len())
}

fn classify_typed_block_owner(
    word: bool,
    local: &[u8],
    elements: &[DynamicXmlElement],
) -> Option<TypedBlockOwner> {
    if !word {
        return None;
    }
    let parent = elements
        .last()
        .and_then(|element| element.typed_block_owner);
    let content_owner = match parent {
        Some(TypedBlockOwner::Body | TypedBlockOwner::Content(BlockControlOwner::Body)) => {
            Some(BlockControlOwner::Body)
        }
        Some(TypedBlockOwner::Table | TypedBlockOwner::Content(BlockControlOwner::Table)) => {
            Some(BlockControlOwner::Table)
        }
        Some(TypedBlockOwner::Row | TypedBlockOwner::Content(BlockControlOwner::Row)) => {
            Some(BlockControlOwner::Row)
        }
        Some(TypedBlockOwner::Cell | TypedBlockOwner::Content(BlockControlOwner::Cell)) => {
            Some(BlockControlOwner::Cell)
        }
        _ => None,
    };
    match local {
        b"document" if elements.is_empty() => Some(TypedBlockOwner::Document),
        b"body" if parent == Some(TypedBlockOwner::Document) => Some(TypedBlockOwner::Body),
        b"p" if matches!(
            content_owner,
            Some(BlockControlOwner::Body | BlockControlOwner::Cell)
        ) =>
        {
            Some(TypedBlockOwner::Paragraph)
        }
        b"tbl"
            if matches!(
                content_owner,
                Some(BlockControlOwner::Body | BlockControlOwner::Cell)
            ) =>
        {
            Some(TypedBlockOwner::Table)
        }
        b"tr" if content_owner == Some(BlockControlOwner::Table) => Some(TypedBlockOwner::Row),
        b"tc" if content_owner == Some(BlockControlOwner::Row) => Some(TypedBlockOwner::Cell),
        b"sdt" if content_owner.is_some() => {
            Some(TypedBlockOwner::ContentControl(content_owner.unwrap()))
        }
        b"sdtContent"
            if matches!(parent, Some(TypedBlockOwner::ContentControl(_)))
                && elements
                    .last()
                    .is_some_and(|control| !control.sdt_content_seen) =>
        {
            let Some(TypedBlockOwner::ContentControl(owner)) = parent else {
                unreachable!();
            };
            Some(TypedBlockOwner::Content(owner))
        }
        _ => None,
    }
}

fn direct_word_run_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    paragraph.is_some()
        && elements.last().is_some_and(|parent| {
            parent.is_typed_inline_owner
                && parent.local_name == b"r"
                && parent.paragraph == paragraph
        })
}

fn dynamic_toc_run_position(
    word: bool,
    local: &[u8],
    is_typed_inline_owner: bool,
    modeled_simple_field: bool,
    elements: &mut [DynamicXmlElement],
    paragraph: Option<usize>,
    next_boundaries: &mut [usize],
    nested_run_orders: &mut [HashMap<(usize, TocRawOrder), usize>],
    paragraph_raw_before: &mut [usize],
) -> Option<TocRunPosition> {
    let paragraph = paragraph?;
    let next = next_boundaries.get_mut(paragraph)?;
    let inherited = elements.last().and_then(|parent| parent.run_position);
    if word && local == b"r" && is_typed_inline_owner {
        let nested = elements.iter().rev().take_while(|element| {
            !element.is_typed_paragraph || element.paragraph != Some(paragraph)
        });
        if nested.into_iter().any(|element| {
            matches!(
                element.local_name.as_slice(),
                b"sdt" | b"sdtContent" | b"ins" | b"moveTo"
            )
        }) {
            let outer = inherited.unwrap_or(TocRunPosition {
                run_boundary: *next,
                raw_order: TocRawOrder::Raw(0),
                nested_order: 0,
            });
            return nested_leaf_position(paragraph, outer, nested_run_orders);
        }
        let boundary = *next;
        let position = nested_leaf_position(
            paragraph,
            TocRunPosition {
                run_boundary: boundary,
                raw_order: TocRawOrder::AfterRaw,
                nested_order: 0,
            },
            nested_run_orders,
        );
        *next += 1;
        *paragraph_raw_before.get_mut(paragraph)? = 0;
        return position;
    }
    if word
        && local == b"fldSimple"
        && modeled_simple_field
        && accepted_simple_field_parent(elements, paragraph)
    {
        if elements.last().is_some_and(|parent| {
            parent.is_typed_inline_owner
                && matches!(parent.local_name.as_slice(), b"ins" | b"moveTo")
        }) {
            let outer = inherited.unwrap_or(TocRunPosition {
                run_boundary: *next,
                raw_order: TocRawOrder::Raw(0),
                nested_order: 0,
            });
            return nested_leaf_position(paragraph, outer, nested_run_orders);
        }
        let boundary = *next;
        let position = nested_leaf_position(
            paragraph,
            TocRunPosition {
                run_boundary: boundary,
                raw_order: TocRawOrder::AfterRaw,
                nested_order: 0,
            },
            nested_run_orders,
        );
        *next += 1;
        *paragraph_raw_before.get_mut(paragraph)? = 0;
        return position;
    }
    if is_typed_inline_owner && local == b"hyperlink" {
        return inherited;
    }
    if is_typed_inline_owner && matches!(local, b"sdt" | b"sdtContent" | b"ins" | b"moveTo") {
        if let Some(inherited) = inherited {
            return Some(inherited);
        }
        let raw_order = if elements
            .last()
            .is_some_and(|parent| parent.is_typed_inline_owner && parent.local_name == b"hyperlink")
        {
            let parent = elements.last_mut()?;
            let plan = parent.hyperlink_plan.as_mut()?;
            let order = *plan.revision_orders.get(plan.next_revision)?;
            plan.next_revision += 1;
            order
        } else {
            TocRawOrder::Raw(*paragraph_raw_before.get(paragraph)?)
        };
        return Some(TocRunPosition {
            run_boundary: *next,
            raw_order,
            nested_order: 0,
        });
    }
    None
}

fn advance_direct_paragraph_raw_child(
    word: bool,
    local: &[u8],
    is_typed_inline_owner: bool,
    modeled_simple_field: bool,
    direct_paragraph_child: bool,
    paragraph: Option<usize>,
    hyperlink_plan: Option<&DynamicHyperlinkPlan>,
    paragraph_raw_before: &mut [usize],
) {
    if !direct_paragraph_child {
        return;
    }
    let modeled_without_raw = word
        && (local == b"pPr"
            || local == b"r" && is_typed_inline_owner
            || local == b"fldSimple" && modeled_simple_field
            || local == b"sdt" && is_typed_inline_owner
            || local == b"hyperlink");
    if modeled_without_raw
        || hyperlink_plan.is_some_and(|plan| plan.preserved_raw)
        || local == b"hyperlink"
    {
        return;
    }
    if let Some(raw_before) = paragraph.and_then(|index| paragraph_raw_before.get_mut(index)) {
        *raw_before += 1;
    }
}

fn dynamic_hyperlink_plan(
    xml: &[u8],
    start: usize,
    bindings: &BTreeMap<String, String>,
    raw_before: usize,
) -> Result<DynamicHyperlinkPlan> {
    let raw = dynamic_element_slice(xml, start)?;
    let mut scoped = Vec::new();
    scoped.extend_from_slice(b"<rdocx-scope");
    for (prefix, namespace) in bindings {
        if prefix == "xml" {
            continue;
        }
        if prefix.is_empty() {
            scoped.extend_from_slice(b" xmlns=\"");
        } else {
            scoped.extend_from_slice(b" xmlns:");
            scoped.extend_from_slice(prefix.as_bytes());
            scoped.extend_from_slice(b"=\"");
        }
        scoped.extend_from_slice(xml_escape_attribute(namespace).as_bytes());
        scoped.push(b'"');
    }
    scoped.push(b'>');
    scoped.extend_from_slice(raw);
    scoped.extend_from_slice(b"</rdocx-scope>");

    let mut reader = NsReader::from_reader(scoped.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut direct_runs = 0usize;
    let mut revision_boundaries = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid hyperlink XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        match event {
            Event::Start(element) => {
                if depth == 2 && word && local_name(element.name().as_ref()) == b"r" {
                    direct_runs += 1;
                } else if depth == 2
                    && word
                    && matches!(local_name(element.name().as_ref()), b"ins" | b"moveTo")
                {
                    let valid_id = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    .is_some_and(|(_, id)| id.parse::<i32>().is_ok());
                    let has_author = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"author",
                        AttributeNamespace::Word,
                    )?
                    .is_some();
                    if valid_id && has_author {
                        revision_boundaries.push(direct_runs);
                    }
                }
                depth += 1;
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "table of contents hyperlink is not balanced".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
    let preserved_raw = direct_runs == 0;
    let revision_orders = revision_boundaries
        .into_iter()
        .map(|boundary| {
            if preserved_raw {
                TocRawOrder::Raw(raw_before)
            } else if boundary == direct_runs {
                TocRawOrder::BeforeRaw
            } else {
                TocRawOrder::AfterRaw
            }
        })
        .collect();
    Ok(DynamicHyperlinkPlan {
        revision_orders,
        next_revision: 0,
        preserved_raw,
    })
}

fn dynamic_element_slice(xml: &[u8], start: usize) -> Result<&[u8]> {
    let tail = xml.get(start..).ok_or_else(|| {
        Error::Other("table of contents element offset is outside the document".to_owned())
    })?;
    let mut reader = quick_xml::Reader::from_reader(tail);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?
        {
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => {
                return Ok(&tail[..reader.buffer_position() as usize]);
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok(&tail[..reader.buffer_position() as usize]);
                }
            }
            Event::Eof => {
                return Err(Error::Other(
                    "table of contents element is not balanced".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn nested_leaf_position(
    paragraph: usize,
    outer: TocRunPosition,
    nested_run_orders: &mut [HashMap<(usize, TocRawOrder), usize>],
) -> Option<TocRunPosition> {
    let next = nested_run_orders
        .get_mut(paragraph)?
        .entry((outer.run_boundary, outer.raw_order))
        .or_default();
    let nested_order = *next;
    *next += 1;
    Some(TocRunPosition {
        nested_order,
        ..outer
    })
}

fn accepted_simple_field_parent(elements: &[DynamicXmlElement], paragraph: usize) -> bool {
    direct_typed_paragraph_parent(elements, Some(paragraph))
        || elements.last().is_some_and(|parent| {
            parent.is_typed_inline_owner
                && matches!(parent.local_name.as_slice(), b"ins" | b"moveTo")
                && parent.paragraph == Some(paragraph)
        })
}

fn typed_inline_owner(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    word: bool,
    local: &[u8],
    elements: &[DynamicXmlElement],
    paragraph: Option<usize>,
) -> Result<bool> {
    if !word || paragraph.is_none() {
        return Ok(false);
    }
    let Some(parent) = elements.last() else {
        return Ok(false);
    };
    let parent_is_paragraph = parent.is_typed_paragraph && parent.paragraph == paragraph;
    let parent_is_inline = parent.is_typed_inline_owner && parent.paragraph == paragraph;
    let parent_local = parent.local_name.as_slice();
    let valid = match local {
        b"r" => {
            parent_is_paragraph
                || (parent_is_inline
                    && matches!(
                        parent_local,
                        b"hyperlink" | b"sdtContent" | b"ins" | b"moveTo"
                    ))
        }
        b"hyperlink" => {
            parent_is_paragraph || (parent_is_inline && matches!(parent_local, b"ins" | b"moveTo"))
        }
        b"sdt" => {
            parent_is_paragraph
                || (parent_is_inline && matches!(parent_local, b"sdtContent" | b"ins" | b"moveTo"))
        }
        b"sdtContent" => parent_is_inline && parent_local == b"sdt" && !parent.sdt_content_seen,
        b"ins" | b"moveTo" => {
            let valid_parent = parent_is_paragraph
                || (parent_is_inline
                    && matches!(
                        parent_local,
                        b"hyperlink" | b"sdtContent" | b"ins" | b"moveTo"
                    ));
            let valid_id =
                resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
                    .is_some_and(|(_, id)| id.parse::<i32>().is_ok());
            let has_author =
                resolved_element_attribute(element, resolver, b"author", AttributeNamespace::Word)?
                    .is_some();
            valid_parent
                && parent.revision_depth < MAX_DYNAMIC_REVISION_NESTING_DEPTH
                && valid_id
                && has_author
        }
        _ => false,
    };
    Ok(valid)
}

fn direct_typed_paragraph_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    let Some(paragraph) = paragraph else {
        return false;
    };
    elements
        .last()
        .is_some_and(|parent| parent.is_typed_paragraph && parent.paragraph == Some(paragraph))
}

fn accepted_bookmark_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    direct_typed_paragraph_parent(elements, paragraph)
        || paragraph.is_some()
            && elements.last().is_some_and(|parent| {
                parent.is_typed_inline_owner
                    && matches!(
                        parent.local_name.as_slice(),
                        b"hyperlink" | b"sdtContent" | b"ins" | b"moveTo"
                    )
                    && parent.paragraph == paragraph
            })
}

fn accepted_simple_toc_parent(elements: &[DynamicXmlElement], paragraph: Option<usize>) -> bool {
    direct_typed_paragraph_parent(elements, paragraph)
        || paragraph.is_some()
            && elements.last().is_some_and(|parent| {
                parent.is_typed_inline_owner
                    && matches!(parent.local_name.as_slice(), b"ins" | b"moveTo")
                    && parent.paragraph == paragraph
            })
}

fn update_dynamic_field_stack(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    event_start: usize,
    paragraph: Option<usize>,
    elements: &[DynamicXmlElement],
    fields: &mut Vec<DynamicFieldScan>,
    spans: &mut Vec<DynamicTocSpan>,
) -> Result<()> {
    let Some(paragraph) = paragraph else {
        return Ok(());
    };
    let kind =
        resolved_element_attribute(element, resolver, b"fldCharType", AttributeNamespace::Word)?
            .map(|(_, value)| value);
    match kind.as_deref() {
        Some("begin") => fields.push(DynamicFieldScan {
            instruction: String::new(),
            field_start: event_start,
            begin_paragraph: paragraph,
            begin_paragraph_start: elements
                .iter()
                .rev()
                .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph))
                .map(|element| element.start)
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents begin marker is not contained by a paragraph".to_owned(),
                    )
                })?,
            begin_run_start: elements
                .iter()
                .rev()
                .find(|element| {
                    element.is_word
                        && element.local_name == b"r"
                        && element.paragraph == Some(paragraph)
                })
                .map(|element| element.start)
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents begin marker is not contained by a run".to_owned(),
                    )
                })?,
            separator_paragraph: None,
            separator_run_start: None,
            result_start: None,
            result_start_position: None,
            start_paragraph_name: None,
            separator_wrapper_names: Vec::new(),
            instruction_runs: Vec::new(),
        }),
        Some("separate") => {
            let Some(field) = fields.last_mut() else {
                return Ok(());
            };
            if field.separator_paragraph.is_some() {
                return Err(Error::Other(
                    "table of contents field has more than one separator".to_owned(),
                ));
            }
            let run = elements
                .iter()
                .rev()
                .find(|element| {
                    element.is_word
                        && element.local_name == b"r"
                        && element.paragraph == Some(paragraph)
                })
                .ok_or_else(|| {
                    Error::Other("table of contents separator is not contained by a run".to_owned())
                })?;
            let para = elements
                .iter()
                .rev()
                .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph))
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents separator is not contained by a paragraph".to_owned(),
                    )
                })?;
            field.separator_paragraph = Some(paragraph);
            field.separator_run_start = Some(run.start);
            field.result_start_position = run.run_position.map(|position| {
                if position.raw_order == TocRawOrder::AfterRaw {
                    TocRunPosition {
                        run_boundary: position.run_boundary + 1,
                        raw_order: TocRawOrder::BeforeRaw,
                        nested_order: 0,
                    }
                } else {
                    TocRunPosition {
                        nested_order: position.nested_order + 1,
                        ..position
                    }
                }
            });
            field.start_paragraph_name = Some(para.qualified_name.clone());
            let paragraph_position = elements
                .iter()
                .position(|element| std::ptr::eq(element, para))
                .expect("paragraph was borrowed from the element stack");
            let run_position = elements
                .iter()
                .rposition(|element| std::ptr::eq(element, run))
                .expect("run was borrowed from the element stack");
            field.separator_wrapper_names = elements[paragraph_position + 1..run_position]
                .iter()
                .filter(|element| element.is_typed_inline_owner)
                .map(|element| element.qualified_name.clone())
                .collect();
        }
        Some("end") => {
            let Some(field) = fields.pop() else {
                return Ok(());
            };
            if Field::new(&field.instruction, "").instruction.name != "TOC" {
                return Ok(());
            }
            let result_start = field.result_start.ok_or_else(|| {
                Error::Other("table of contents field is missing its separator".to_owned())
            })?;
            let separator_paragraph = field.separator_paragraph.ok_or_else(|| {
                Error::Other("table of contents field is missing its separator".to_owned())
            })?;
            if separator_paragraph == paragraph {
                return Err(Error::Other(
                    "table of contents result must span paragraph boundaries".to_owned(),
                ));
            }
            if separator_paragraph != field.begin_paragraph {
                return Err(Error::Other(
                    "table of contents instruction crosses a paragraph boundary".to_owned(),
                ));
            }
            let end_run = elements
                .iter()
                .rev()
                .find(|element| {
                    element.is_word
                        && element.local_name == b"r"
                        && element.paragraph == Some(paragraph)
                })
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents end marker is not contained by a run".to_owned(),
                    )
                })?;
            let end_para = elements
                .iter()
                .rev()
                .find(|element| element.is_typed_paragraph && element.paragraph == Some(paragraph))
                .ok_or_else(|| {
                    Error::Other(
                        "table of contents end marker is not contained by a paragraph".to_owned(),
                    )
                })?;
            if event_start < result_start || end_run.start < result_start {
                return Err(Error::Other(
                    "table of contents result boundaries are reversed".to_owned(),
                ));
            }
            let paragraph_position = elements
                .iter()
                .position(|element| std::ptr::eq(element, end_para))
                .expect("end paragraph was borrowed from the element stack");
            let run_position = elements
                .iter()
                .rposition(|element| std::ptr::eq(element, end_run))
                .expect("end run was borrowed from the element stack");
            let wrapper_chain = &elements[paragraph_position + 1..run_position];
            let mut end_wrapper_prefixes = Vec::new();
            let mut wrapper_index = 0usize;
            while wrapper_index < wrapper_chain.len() {
                let wrapper = &wrapper_chain[wrapper_index];
                if !wrapper.is_typed_inline_owner {
                    wrapper_index += 1;
                    continue;
                }
                if wrapper.local_name == b"sdt"
                    && let Some(content) = wrapper_chain.get(wrapper_index + 1)
                    && content.is_typed_inline_owner
                    && content.local_name == b"sdtContent"
                {
                    end_wrapper_prefixes.push((wrapper.start, content.start_tag_end));
                    wrapper_index += 2;
                } else {
                    end_wrapper_prefixes.push((wrapper.start, wrapper.start_tag_end));
                    wrapper_index += 1;
                }
            }
            spans.push(DynamicTocSpan {
                instruction: field.instruction,
                field_start: field.field_start,
                field_end: event_start,
                begin_paragraph: field.begin_paragraph,
                end_paragraph: paragraph,
                begin_run_start: field.begin_run_start,
                instruction_paragraph_start: field.begin_paragraph_start,
                result_start,
                result_end: end_run.start,
                result_start_position: field.result_start_position.ok_or_else(|| {
                    Error::Other(
                        "table of contents separator run position was not retained".to_owned(),
                    )
                })?,
                result_end_position: end_run.run_position.ok_or_else(|| {
                    Error::Other("table of contents end run position was not retained".to_owned())
                })?,
                end_run_end: 0,
                start_paragraph_name: field.start_paragraph_name.ok_or_else(|| {
                    Error::Other("table of contents field is missing its separator".to_owned())
                })?,
                separator_wrapper_names: field.separator_wrapper_names,
                instruction_runs: field.instruction_runs,
                end_paragraph_start: end_para.start,
                end_paragraph_content_start: 0,
                end_wrapper_prefixes,
            });
        }
        _ => {}
    }
    Ok(())
}

fn parse_dynamic_toc_fields(
    document: &Document,
    xml: &[u8],
    spans: &[DynamicTocSpan],
) -> Result<(Vec<Option<TocField>>, usize)> {
    let mut paragraphs = Vec::new();
    collect_body_paragraphs(&document.document.body, &mut paragraphs);
    let context = FieldEvaluationContext::default();
    let mut output = Vec::with_capacity(spans.len());
    let mut diagnostic_count = 0usize;
    for span in spans {
        let field = parse_dynamic_toc_field(xml, span)?;
        let mut evaluator = Evaluator::new(document, &context);
        match evaluator.evaluate_field(&field, "main", &paragraphs, span.begin_paragraph) {
            FieldOutcome::TableOfContents(toc) => output.push(Some(toc)),
            FieldOutcome::KeepStored { .. } => {
                diagnostic_count += 1;
                output.push(None);
            }
            _ => {
                return Err(Error::Other(
                    "table of contents instruction was not recognized".to_owned(),
                ));
            }
        }
    }
    Ok((output, diagnostic_count))
}

fn parse_dynamic_toc_field(xml: &[u8], span: &DynamicTocSpan) -> Result<Field> {
    let prefix = span
        .start_paragraph_name
        .split_once(':')
        .map_or("w", |(prefix, _)| prefix);
    let mut source = xml[span.instruction_paragraph_start..span.result_start].to_vec();
    source.extend_from_slice(
        format!("<{prefix}:r><{prefix}:fldChar {prefix}:fldCharType=\"end\"/></{prefix}:r>")
            .as_bytes(),
    );
    append_toc_wrapper_closures(&mut source, &span.separator_wrapper_names);
    source.extend_from_slice(b"</");
    source.extend_from_slice(span.start_paragraph_name.as_bytes());
    source.push(b'>');
    let parse_paragraph = |source: &[u8]| -> Result<CT_P> {
        let mut reader = quick_xml::Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).map_err(|error| {
                Error::Other(format!("invalid table of contents field: {error}"))
            })? {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    return Ok(CT_P::from_xml(&mut reader)?);
                }
                Event::Eof => {
                    return Err(Error::Other(
                        "table of contents instruction paragraph was not found".to_owned(),
                    ));
                }
                _ => {}
            }
            buffer.clear();
        }
    };
    parse_paragraph(&source)?;

    let mut projected = format!("<w:p xmlns:w=\"{W_NS}\">").into_bytes();
    for run in &span.instruction_runs {
        append_instruction_run_with_namespaces(&mut projected, &xml[run.start..run.end], run)?;
    }
    projected.extend_from_slice(b"<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
    let paragraph = parse_paragraph(&projected)?;
    accepted_toc_runs(&paragraph)
        .into_iter()
        .flat_map(|run| &run.run.content)
        .find_map(|content| match content {
            RunContent::Field(field) if field.effective_instruction().name == "TOC" => {
                Some(field.clone())
            }
            _ => None,
        })
        .ok_or_else(|| {
            Error::Other(format!(
                "table of contents instruction could not be parsed: {}",
                span.instruction.trim()
            ))
        })
}

fn append_instruction_run_with_namespaces(
    output: &mut Vec<u8>,
    raw: &[u8],
    run: &DynamicInstructionRun,
) -> Result<()> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (insertion, local_namespaces) =
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!(
                "invalid table of contents instruction run: {error}"
            ))
        })? {
            Event::Start(start) | Event::Empty(start) => {
                let mut local_namespaces = HashSet::new();
                for attribute in start.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!(
                            "invalid table of contents instruction run: {error}"
                        ))
                    })?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" {
                        local_namespaces.insert(String::new());
                    } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                        local_namespaces.insert(String::from_utf8_lossy(prefix).into_owned());
                    }
                }
                let tag_end = reader.buffer_position() as usize;
                let insertion = if tag_end >= 2 && raw[tag_end - 2] == b'/' {
                    tag_end - 2
                } else {
                    tag_end - 1
                };
                (insertion, local_namespaces)
            }
            _ => {
                return Err(Error::Other(
                    "table of contents instruction run has no start tag".to_owned(),
                ));
            }
        };
    output.extend_from_slice(&raw[..insertion]);
    for (prefix, namespace) in &run.inherited_namespaces {
        if prefix == "xml" || local_namespaces.contains(prefix) {
            continue;
        }
        if prefix.is_empty() {
            output.extend_from_slice(b" xmlns=\"");
        } else {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
        }
        output.extend_from_slice(xml_escape_attribute(namespace).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(&raw[insertion..]);
    Ok(())
}

#[derive(Debug, Clone)]
struct TocBookmark {
    id: i32,
    name: String,
    insert: bool,
}

#[derive(Debug)]
struct TocBookmarkState {
    max_id: i32,
    names: HashSet<String>,
    named_ranges: HashMap<String, (TocDocumentPosition, TocDocumentPosition)>,
    whole_paragraphs: HashMap<usize, (i32, String)>,
    ranges: Vec<TocBookmarkRange>,
}

#[derive(Debug, Clone)]
struct TocBookmarkRange {
    id: i32,
    name: String,
    start_offset: usize,
    end_offset: usize,
}

#[derive(Debug)]
struct TocBookmarkRepair {
    id: i32,
    name: String,
    toc_index: usize,
    replace_start: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
enum TocRawOrder {
    BeforeRaw,
    Raw(usize),
    AfterRaw,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct TocRunPosition {
    run_boundary: usize,
    raw_order: TocRawOrder,
    nested_order: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct TocDocumentPosition {
    paragraph: usize,
    accepted_run_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct TocOwnedPosition {
    paragraph: usize,
    run: TocRunPosition,
}

fn inspect_toc_bookmarks(body: &CT_Body, xml: &[u8]) -> Result<TocBookmarkState> {
    let whole_bookmarks = whole_paragraph_bookmark_ids(xml)?;
    let marker_offsets = toc_bookmark_marker_offsets(xml)?;
    let mut paragraphs = Vec::new();
    collect_body_paragraphs(body, &mut paragraphs);
    let mut starts = HashMap::<i32, (String, usize, usize)>::new();
    let mut start_order = Vec::new();
    let mut ends = HashMap::<i32, (usize, usize)>::new();
    let mut names = HashSet::new();
    let mut max_id = 0i32;
    for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
        for marker in &paragraph.bookmark_markers {
            let Some(id) = marker.id() else {
                return Err(Error::Other(
                    "table of contents source has a bookmark without an ID".to_owned(),
                ));
            };
            max_id = max_id.max(id);
            if marker.is_start() {
                let name = marker.name().ok_or_else(|| {
                    Error::Other(
                        "table of contents source has a bookmark without a name".to_owned(),
                    )
                })?;
                if !names.insert(name.to_owned()) || starts.contains_key(&id) {
                    return Err(Error::Other(
                        "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                    ));
                }
                starts.insert(
                    id,
                    (
                        name.to_owned(),
                        paragraph_index,
                        marker.projected_run_index(),
                    ),
                );
                start_order.push(id);
            } else if ends
                .insert(id, (paragraph_index, marker.projected_run_index()))
                .is_some()
            {
                return Err(Error::Other(
                    "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                ));
            }
        }
    }
    if starts.len() != ends.len() || starts.keys().any(|id| !ends.contains_key(id)) {
        return Err(Error::Other(
            "table of contents source has an unmatched bookmark".to_owned(),
        ));
    }
    let mut named_ranges = HashMap::new();
    let mut whole_paragraphs = HashMap::new();
    let mut ranges = Vec::new();
    for id in start_order {
        let (name, start_paragraph, start_run) = starts
            .remove(&id)
            .expect("bookmark start order contains each validated start once");
        let (end_paragraph, end_run) = ends[&id];
        let start_offset = marker_offsets.get(&(id, true)).copied().ok_or_else(|| {
            Error::Other("table of contents bookmark start offset was not retained".to_owned())
        })?;
        let end_offset = marker_offsets.get(&(id, false)).copied().ok_or_else(|| {
            Error::Other("table of contents bookmark end offset was not retained".to_owned())
        })?;
        if start_offset > end_offset {
            return Err(Error::Other(
                "table of contents source has a reversed bookmark".to_owned(),
            ));
        }
        named_ranges.insert(
            name.clone(),
            (
                TocDocumentPosition {
                    paragraph: start_paragraph,
                    accepted_run_index: start_run,
                },
                TocDocumentPosition {
                    paragraph: end_paragraph,
                    accepted_run_index: end_run,
                },
            ),
        );
        ranges.push(TocBookmarkRange {
            id,
            name: name.clone(),
            start_offset,
            end_offset,
        });
        if start_paragraph == end_paragraph && whole_bookmarks.contains(&id) {
            whole_paragraphs
                .entry(start_paragraph)
                .or_insert((id, name));
        }
    }
    Ok(TocBookmarkState {
        max_id,
        names,
        named_ranges,
        whole_paragraphs,
        ranges,
    })
}

fn toc_crossing_bookmark_repairs(
    state: &TocBookmarkState,
    spans: &[DynamicTocSpan],
    fields: &[Option<TocField>],
) -> Vec<TocBookmarkRepair> {
    state
        .ranges
        .iter()
        .filter_map(|range| {
            let containing_span = |offset: usize| {
                spans
                    .iter()
                    .zip(fields)
                    .enumerate()
                    .find(|(_, (span, field))| {
                        field.is_some() && offset >= span.result_start && offset < span.result_end
                    })
                    .map(|(index, _)| index)
            };
            let start_span = containing_span(range.start_offset);
            let end_span = containing_span(range.end_offset);
            match (start_span, end_span) {
                (Some(toc_index), None) => Some(TocBookmarkRepair {
                    id: range.id,
                    name: range.name.clone(),
                    toc_index,
                    replace_start: true,
                }),
                (None, Some(toc_index)) => Some(TocBookmarkRepair {
                    id: range.id,
                    name: range.name.clone(),
                    toc_index,
                    replace_start: false,
                }),
                _ => None,
            }
        })
        .collect()
}

#[derive(Clone, Copy)]
enum TocParagraphToken {
    Content,
    BookmarkStart(i32),
    BookmarkEnd(i32),
}

fn toc_bookmark_marker_offsets(xml: &[u8]) -> Result<HashMap<(i32, bool), usize>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut offsets = HashMap::new();
    let mut paragraph_count = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                invalidate_overdeep_revision_owner(word, &local, &mut elements);
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    let paragraph = paragraph_count;
                    paragraph_count += 1;
                    Some(paragraph)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                let mut is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                if word
                    && matches!(local.as_slice(), b"bookmarkStart" | b"bookmarkEnd")
                    && accepted_bookmark_parent(&elements, paragraph)
                    && let Some((_, id)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    && let Ok(id) = id.parse::<i32>()
                    && offsets
                        .insert((id, local == b"bookmarkStart"), before)
                        .is_some()
                {
                    return Err(Error::Other(
                        "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                    ));
                }
                let revision_depth = if is_typed_inline_owner {
                    elements.last().map_or(0, |parent| {
                        parent.revision_depth
                            + usize::from(matches!(local.as_slice(), b"ins" | b"moveTo"))
                    })
                } else {
                    0
                };
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner,
                    revision_depth,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                invalidate_overdeep_revision_owner(word, &local, &mut elements);
                let paragraph = elements.last().and_then(|element| element.paragraph);
                let typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let is_typed_inline_owner = typed_inline_owner(
                    &element,
                    reader.resolver(),
                    word,
                    &local,
                    &elements,
                    paragraph,
                )?;
                if word
                    && matches!(local.as_slice(), b"bookmarkStart" | b"bookmarkEnd")
                    && accepted_bookmark_parent(&elements, paragraph)
                    && let Some((_, id)) = resolved_element_attribute(
                        &element,
                        reader.resolver(),
                        b"id",
                        AttributeNamespace::Word,
                    )?
                    && let Ok(id) = id.parse::<i32>()
                    && offsets
                        .insert((id, local == b"bookmarkStart"), before)
                        .is_some()
                {
                    return Err(Error::Other(
                        "table of contents source has ambiguous duplicate bookmarks".to_owned(),
                    ));
                }
                mark_typed_sdt_content(
                    &mut elements,
                    &local,
                    typed_block_owner,
                    is_typed_inline_owner,
                );
            }
            Event::End(_) => {
                elements.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(offsets)
}

fn whole_paragraph_bookmark_ids(xml: &[u8]) -> Result<HashSet<i32>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut paragraphs = Vec::<Vec<TocParagraphToken>>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let mut is_typed_inline_owner = false;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    paragraphs.push(Vec::new());
                    Some(paragraphs.len() - 1)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                if let Some(index) = paragraph
                    && elements
                        .last()
                        .is_some_and(|parent| parent.is_typed_paragraph)
                    && local != b"pPr"
                {
                    paragraphs[index].push(paragraph_token(
                        &element,
                        reader.resolver(),
                        word,
                        &local,
                    )?);
                }
                mark_typed_sdt_content(&mut elements, &local, typed_block_owner, false);
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner: false,
                    revision_depth: 0,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                let typed_block_owner = classify_typed_block_owner(word, local, &elements);
                mark_typed_sdt_content(&mut elements, local, typed_block_owner, false);
                if let Some(parent) = elements.last()
                    && parent.is_typed_paragraph
                    && local != b"pPr"
                    && let Some(index) = parent.paragraph
                {
                    paragraphs[index].push(paragraph_token(
                        &element,
                        reader.resolver(),
                        word,
                        local,
                    )?);
                }
            }
            Event::End(_) => {
                elements.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut whole = HashSet::new();
    for tokens in paragraphs {
        let content = tokens
            .iter()
            .enumerate()
            .filter_map(|(index, token)| {
                matches!(token, TocParagraphToken::Content).then_some(index)
            })
            .collect::<Vec<_>>();
        let mut starts = HashMap::new();
        let mut ends = HashMap::new();
        for (index, token) in tokens.iter().enumerate() {
            match token {
                TocParagraphToken::BookmarkStart(id) => {
                    starts.insert(*id, index);
                }
                TocParagraphToken::BookmarkEnd(id) => {
                    ends.insert(*id, index);
                }
                TocParagraphToken::Content => {}
            }
        }
        for (id, start) in starts {
            let Some(end) = ends.get(&id).copied() else {
                continue;
            };
            if content
                .iter()
                .all(|position| *position > start && *position < end)
            {
                whole.insert(id);
            }
        }
    }
    Ok(whole)
}

fn paragraph_token(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    word: bool,
    local: &[u8],
) -> Result<TocParagraphToken> {
    if word && matches!(local, b"bookmarkStart" | b"bookmarkEnd") {
        let id = resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
            .and_then(|(_, id)| id.parse::<i32>().ok());
        if let Some(id) = id {
            return Ok(if local == b"bookmarkStart" {
                TocParagraphToken::BookmarkStart(id)
            } else {
                TocParagraphToken::BookmarkEnd(id)
            });
        }
    }
    Ok(TocParagraphToken::Content)
}

struct TocBookmarkAllocator {
    next_id: Option<i32>,
    next_name: Option<u32>,
    names: HashSet<String>,
    whole_paragraphs: HashMap<usize, (i32, String)>,
}

impl TocBookmarkAllocator {
    fn new(state: TocBookmarkState) -> Self {
        Self {
            next_id: state.max_id.checked_add(1),
            next_name: Some(1),
            names: state.names,
            whole_paragraphs: state.whole_paragraphs,
        }
    }

    fn whole_paragraph_name(&self, paragraph: usize) -> Option<(i32, String)> {
        self.whole_paragraphs.get(&paragraph).cloned()
    }

    fn allocate(&mut self) -> Result<(i32, String)> {
        let id = self.next_id.ok_or_else(|| {
            Error::Other("table of contents exhausted the bookmark ID range".to_owned())
        })?;
        loop {
            let suffix = self.next_name.ok_or_else(|| {
                Error::Other("table of contents exhausted the bookmark name range".to_owned())
            })?;
            let name = format!("_Toc{suffix}");
            self.next_name = suffix.checked_add(1);
            if self.names.insert(name.clone()) {
                self.next_id = id.checked_add(1);
                return Ok((id, name));
            }
        }
    }
}

#[derive(Debug, Clone)]
struct TocSource {
    paragraph_index: usize,
    level: u8,
    title: String,
    omit_page_number: bool,
    needs_bookmark: bool,
    sequence_prefix: Option<String>,
}

fn discover_toc_sources(
    document: &Document,
    spans: &[DynamicTocSpan],
    fields: &[Option<TocField>],
    bookmarks: &TocBookmarkState,
) -> Result<Vec<Vec<TocSource>>> {
    let mut paragraphs = Vec::new();
    collect_body_paragraphs(&document.document.body, &mut paragraphs);
    let context = FieldEvaluationContext::default();
    let mut all_sources = Vec::with_capacity(fields.len());
    for toc in fields {
        let Some(toc) = toc else {
            all_sources.push(Vec::new());
            continue;
        };
        let bookmark_range = toc
            .bookmark
            .as_ref()
            .map(|name| {
                bookmarks.named_ranges.get(name).copied().ok_or_else(|| {
                    Error::Other(format!(
                        "table of contents source bookmark {name} was not found"
                    ))
                })
            })
            .transpose()?;
        let mut sources = Vec::new();
        let mut evaluator = Evaluator::new(document, &context);
        let mut sequence_value = None;
        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            let paragraph_fully_owned = spans.iter().any(|span| {
                paragraph_index > span.begin_paragraph && paragraph_index < span.end_paragraph
            });
            if paragraph_fully_owned {
                continue;
            }
            let accepted_runs = accepted_toc_runs(paragraph);
            for (accepted_run_index, run) in accepted_runs.iter().enumerate() {
                let position = TocDocumentPosition {
                    paragraph: paragraph_index,
                    accepted_run_index,
                };
                let in_bookmark =
                    bookmark_range.is_none_or(|(start, end)| position >= start && position < end);
                for content in &run.run.content {
                    let RunContent::Field(field) = content else {
                        continue;
                    };
                    let owned_position = TocOwnedPosition {
                        paragraph: paragraph_index,
                        run: TocRunPosition {
                            run_boundary: run.run_boundary,
                            raw_order: run.raw_order,
                            nested_order: run.nested_order,
                        },
                    };
                    if toc_source_position_is_owned(spans, owned_position) {
                        continue;
                    }
                    let outcome =
                        evaluator.evaluate_field(field, "main", &paragraphs, paragraph_index);
                    if field.instruction.name == "SEQ"
                        && let Some(identifier) = field.instruction.arguments.first()
                        && field_argument_text(identifier).is_some_and(|name| {
                            toc.sequence_identifier
                                .as_deref()
                                .is_some_and(|selected| selected.eq_ignore_ascii_case(name))
                        })
                        && let FieldOutcome::Resolved(ref value) = outcome
                    {
                        sequence_value = Some(value.clone());
                    }
                    if !in_bookmark {
                        continue;
                    }
                    let FieldOutcome::TableOfContentsEntry(tc) = outcome else {
                        continue;
                    };
                    if toc_accepts_tc(toc, &tc) {
                        let omit_page_number = tc.omit_page_number
                            || toc
                                .omit_page_number_levels
                                .is_some_and(|(start, end)| (start..=end).contains(&tc.level));
                        sources.push(TocSource {
                            paragraph_index,
                            level: tc.level,
                            title: tc.entry,
                            omit_page_number,
                            needs_bookmark: toc.hyperlink || !omit_page_number,
                            sequence_prefix: sequence_value.clone(),
                        });
                    }
                }
            }
            let paragraph_in_bookmark = bookmark_range.is_none_or(|(start, end)| {
                accepted_runs
                    .iter()
                    .enumerate()
                    .any(|(accepted_run_index, run)| {
                        let position = TocDocumentPosition {
                            paragraph: paragraph_index,
                            accepted_run_index,
                        };
                        let owned_position = TocOwnedPosition {
                            paragraph: paragraph_index,
                            run: TocRunPosition {
                                run_boundary: run.run_boundary,
                                raw_order: run.raw_order,
                                nested_order: run.nested_order,
                            },
                        };
                        !toc_source_position_is_owned(spans, owned_position)
                            && position >= start
                            && position < end
                    })
            });
            if !paragraph_in_bookmark {
                continue;
            }
            let Some(level) = toc_paragraph_level(document, toc, paragraph) else {
                continue;
            };
            let title = accepted_runs
                .iter()
                .filter(|run| {
                    !toc_source_position_is_owned(
                        spans,
                        TocOwnedPosition {
                            paragraph: paragraph_index,
                            run: TocRunPosition {
                                run_boundary: run.run_boundary,
                                raw_order: run.raw_order,
                                nested_order: run.nested_order,
                            },
                        },
                    )
                })
                .map(|run| run.run.text())
                .collect::<String>();
            if title.is_empty() {
                continue;
            }
            let omit_page_number = toc
                .omit_page_number_levels
                .is_some_and(|(start, end)| (start..=end).contains(&level));
            sources.push(TocSource {
                paragraph_index,
                level,
                title,
                omit_page_number,
                needs_bookmark: toc.hyperlink || !omit_page_number,
                sequence_prefix: sequence_value.clone(),
            });
        }
        sources.sort_by_key(|source| source.paragraph_index);
        all_sources.push(sources);
    }
    Ok(all_sources)
}

fn toc_source_position_is_owned(spans: &[DynamicTocSpan], position: TocOwnedPosition) -> bool {
    spans.iter().any(|span| {
        if position.paragraph > span.begin_paragraph && position.paragraph < span.end_paragraph {
            return true;
        }
        if position.paragraph == span.begin_paragraph {
            return position.run >= span.result_start_position;
        }
        if position.paragraph == span.end_paragraph {
            return position.run < span.result_end_position;
        }
        false
    })
}

#[derive(Clone, Copy)]
struct AcceptedTocRun<'a> {
    run: &'a CT_R,
    run_boundary: usize,
    raw_order: TocRawOrder,
    nested_order: usize,
}

fn accepted_toc_runs(paragraph: &CT_P) -> Vec<AcceptedTocRun<'_>> {
    let mut runs = Vec::new();
    append_accepted_paragraph_runs(paragraph, None, &mut runs);
    let mut nested_run_orders = HashMap::<(usize, TocRawOrder), usize>::new();
    for run in &mut runs {
        let next = nested_run_orders
            .entry((run.run_boundary, run.raw_order))
            .or_default();
        run.nested_order = *next;
        *next += 1;
    }
    runs
}

fn append_accepted_paragraph_runs<'a>(
    paragraph: &'a CT_P,
    inherited_position: Option<(usize, TocRawOrder)>,
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=paragraph.runs.len() {
        let mut owners = paragraph
            .content_controls
            .iter()
            .filter(|(at, _, _, _)| *at == boundary)
            .map(|(_, raw_before, _, control)| {
                (
                    TocRawOrder::Raw(*raw_before),
                    0u8,
                    AcceptedParagraphOwner::Control(control),
                )
            })
            .chain(
                paragraph
                    .revisions
                    .iter()
                    .filter(|(at, _, _)| *at == boundary)
                    .map(|(_, slot, revision)| {
                        (
                            accepted_revision_raw_order(paragraph, boundary, *slot),
                            1u8,
                            AcceptedParagraphOwner::Revision(revision),
                        )
                    }),
            )
            .collect::<Vec<_>>();
        owners.sort_by_key(|(raw_order, kind, _)| (*raw_order, *kind));
        for (raw_order, _, owner) in owners {
            let position = inherited_position.unwrap_or((boundary, raw_order));
            match owner {
                AcceptedParagraphOwner::Control(control) => {
                    append_accepted_control_runs(control, position, runs);
                }
                AcceptedParagraphOwner::Revision(revision) => {
                    append_accepted_revision_runs(revision, position, runs);
                }
            }
        }
        if let Some(run) = paragraph.runs.get(boundary) {
            let (run_boundary, raw_order) =
                inherited_position.unwrap_or((boundary, TocRawOrder::AfterRaw));
            runs.push(AcceptedTocRun {
                run,
                run_boundary,
                raw_order,
                nested_order: 0,
            });
        }
    }
}

fn accepted_revision_raw_order(paragraph: &CT_P, boundary: usize, slot: usize) -> TocRawOrder {
    let Some(index) = hyperlink_revision_index(slot) else {
        return TocRawOrder::Raw(slot);
    };
    if let Some(raw_before) = paragraph
        .hyperlinks
        .get(index)
        .and_then(|hyperlink| hyperlink.preserved_raw_before)
    {
        TocRawOrder::Raw(raw_before)
    } else if paragraph
        .hyperlinks
        .get(index)
        .is_some_and(|hyperlink| boundary == hyperlink.run_end)
    {
        TocRawOrder::BeforeRaw
    } else {
        TocRawOrder::AfterRaw
    }
}

enum AcceptedParagraphOwner<'a> {
    Control(&'a CT_Sdt),
    Revision(&'a CT_Revision),
}

fn append_accepted_control_runs<'a>(
    control: &'a CT_Sdt,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=control.content.len() {
        for (_, revision) in control.revisions().iter().filter(|(at, _)| *at == boundary) {
            append_accepted_revision_runs(revision, position, runs);
        }
        if let Some(content) = control.content.get(boundary) {
            match content {
                SdtContent::Run(run) => runs.push(AcceptedTocRun {
                    run,
                    run_boundary: position.0,
                    raw_order: position.1,
                    nested_order: 0,
                }),
                SdtContent::ContentControl(control) => {
                    append_accepted_control_runs(control, position, runs)
                }
                SdtContent::Paragraph(paragraph) => {
                    append_accepted_paragraph_runs(paragraph, Some(position), runs)
                }
                SdtContent::Table(table) => append_accepted_table_runs(table, position, runs),
                SdtContent::Row(row) => append_accepted_row_runs(row, position, runs),
                SdtContent::Cell(cell) => append_accepted_cell_runs(cell, position, runs),
                SdtContent::RawXml(_) => {}
            }
        }
    }
}

fn append_accepted_revision_runs<'a>(
    revision: &'a CT_Revision,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    match revision.kind() {
        RevisionKind::Insertion | RevisionKind::MoveTo => {
            if let Some(paragraph) = revision.content_paragraph() {
                append_accepted_paragraph_runs(paragraph, Some(position), runs);
                return;
            }
            let RevisionContent::Runs(direct) = revision.content() else {
                return;
            };
            for boundary in 0..=direct.len() {
                for (_, nested) in revision
                    .nested_revisions()
                    .iter()
                    .filter(|(at, _)| *at == boundary)
                {
                    append_accepted_revision_runs(nested, position, runs);
                }
                if let Some(run) = direct.get(boundary) {
                    runs.push(AcceptedTocRun {
                        run,
                        run_boundary: position.0,
                        raw_order: position.1,
                        nested_order: 0,
                    });
                }
            }
        }
        RevisionKind::Deletion
        | RevisionKind::MoveFrom
        | RevisionKind::RunPropertyChange
        | RevisionKind::ParagraphPropertyChange
        | RevisionKind::TablePropertyChange
        | RevisionKind::SectionPropertyChange => {}
    }
}

fn append_accepted_table_runs<'a>(
    table: &'a CT_Tbl,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_accepted_control_runs(control, position, runs);
        }
        if let Some(row) = table.rows.get(boundary) {
            append_accepted_row_runs(row, position, runs);
        }
    }
}

fn append_accepted_row_runs<'a>(
    row: &'a CT_Row,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(at, _, _)| *at == boundary)
        {
            append_accepted_control_runs(control, position, runs);
        }
        if let Some(cell) = row.cells.get(boundary) {
            append_accepted_cell_runs(cell, position, runs);
        }
    }
}

fn append_accepted_cell_runs<'a>(
    cell: &'a CT_Tc,
    position: (usize, TocRawOrder),
    runs: &mut Vec<AcceptedTocRun<'a>>,
) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                append_accepted_paragraph_runs(paragraph, Some(position), runs)
            }
            CellContent::Table(table) => append_accepted_table_runs(table, position, runs),
            CellContent::ContentControl(control) => {
                append_accepted_control_runs(control, position, runs)
            }
        }
    }
}

fn field_argument_text(argument: &FieldArgument) -> Option<&str> {
    match argument {
        FieldArgument::Text(value) => Some(value),
        FieldArgument::Nested(_) => None,
    }
}

fn toc_accepts_tc(toc: &TocField, tc: &TcField) -> bool {
    match &toc.entries {
        TocEntrySelection::None => false,
        TocEntrySelection::All => true,
        TocEntrySelection::Identifier(identifier) => tc
            .table_identifier
            .as_deref()
            .is_some_and(|candidate| candidate.eq_ignore_ascii_case(identifier)),
    }
}

fn toc_paragraph_level(document: &Document, toc: &TocField, paragraph: &CT_P) -> Option<u8> {
    let properties = paragraph.properties.as_ref()?;
    if let Some(style_id) = properties.style_id.as_deref() {
        for (name, level) in &toc.custom_styles {
            let matches_id = style_id.eq_ignore_ascii_case(name);
            let matches_name = document
                .styles
                .get_by_id(style_id)
                .and_then(|style| style.name.as_deref())
                .is_some_and(|style_name| style_name.eq_ignore_ascii_case(name));
            if matches_id || matches_name {
                return Some(*level);
            }
        }
        if let Some((start, end)) = toc.heading_levels
            && let Some(level) = heading_style_level(document, style_id)
            && (start..=end).contains(&level)
        {
            return Some(level);
        }
    }
    if toc.use_outline_levels {
        let level = properties
            .outline_lvl
            .and_then(|level| level.checked_add(1))
            .and_then(|level| u8::try_from(level).ok())?;
        if (1..=9).contains(&level) {
            return Some(level);
        }
    }
    None
}

fn heading_style_level(document: &Document, style_id: &str) -> Option<u8> {
    let candidates = std::iter::once(style_id).chain(
        document
            .styles
            .get_by_id(style_id)
            .and_then(|style| style.name.as_deref()),
    );
    for candidate in candidates {
        let normalized = candidate
            .chars()
            .filter(|character| !character.is_whitespace())
            .collect::<String>()
            .to_ascii_lowercase();
        if let Some(suffix) = normalized.strip_prefix("heading")
            && let Ok(level) = suffix.parse::<u8>()
            && (1..=9).contains(&level)
        {
            return Some(level);
        }
    }
    None
}

#[derive(Debug)]
struct TocParagraphInsertion {
    content_start: usize,
    content_end: usize,
}

fn insert_toc_bookmarks_xml(
    xml: &[u8],
    bookmarks: &BTreeMap<usize, TocBookmark>,
    repairs: &[TocBookmarkRepair],
    all_toc_spans: &[DynamicTocSpan],
    toc_spans: &[DynamicTocSpan],
) -> Result<Vec<u8>> {
    let paragraphs = toc_paragraph_insertions(xml)?;
    if bookmarks.keys().any(|index| *index >= paragraphs.len()) {
        return Err(Error::Other(
            "table of contents source paragraph was not found in package XML".to_owned(),
        ));
    }
    let mut insertions = BTreeMap::<usize, Vec<u8>>::new();
    for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
        let bookmark = bookmarks
            .get(&paragraph_index)
            .filter(|bookmark| bookmark.insert);
        if bookmark.is_none() {
            continue;
        }
        let end_boundary = toc_spans
            .iter()
            .any(|span| span.end_paragraph == paragraph_index);
        let fragment_start = toc_spans
            .iter()
            .filter(|span| span.end_paragraph == paragraph_index)
            .map(|span| span.end_run_end)
            .max()
            .unwrap_or(paragraph.content_start);
        let content_end = toc_spans
            .iter()
            .filter(|span| span.begin_paragraph == paragraph_index)
            .map(|span| span.begin_run_start)
            .min()
            .unwrap_or(paragraph.content_end);
        if fragment_start > content_end {
            return Err(Error::Other(
                "table of contents source bookmark has no unowned paragraph range".to_owned(),
            ));
        }
        if !end_boundary && let Some(bookmark) = bookmark {
            let start_insertion = insertions.entry(fragment_start).or_default();
            start_insertion.extend_from_slice(
                format!(
                    "<w:bookmarkStart w:id=\"{}\" w:name=\"{}\"/>",
                    bookmark.id,
                    xml_escape_attribute(&bookmark.name)
                )
                .as_bytes(),
            );
        }

        let end_insertion = insertions.entry(content_end).or_default();
        if let Some(bookmark) = bookmark {
            end_insertion
                .extend_from_slice(format!("<w:bookmarkEnd w:id=\"{}\"/>", bookmark.id).as_bytes());
        }
    }
    let mut repair_ends = BTreeMap::<usize, Vec<i32>>::new();
    for repair in repairs.iter().filter(|repair| !repair.replace_start) {
        let span = all_toc_spans.get(repair.toc_index).ok_or_else(|| {
            Error::Other("table of contents bookmark repair owner was not found".to_owned())
        })?;
        repair_ends
            .entry(span.begin_run_start)
            .or_default()
            .push(repair.id);
    }
    for (position, ids) in repair_ends {
        let insertion = insertions.entry(position).or_default();
        for id in ids.into_iter().rev() {
            insertion.extend_from_slice(format!("<w:bookmarkEnd w:id=\"{id}\"/>").as_bytes());
        }
    }
    let mut output = xml.to_vec();
    for (position, replacement) in insertions.into_iter().rev() {
        output.splice(position..position, replacement);
    }
    Ok(output)
}

fn toc_paragraph_insertions(xml: &[u8]) -> Result<Vec<TocParagraphInsertion>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut elements = Vec::<DynamicXmlElement>::new();
    let mut paragraphs = Vec::<TocParagraphInsertion>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid table of contents XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = local_name(element.name().as_ref()).to_vec();
                let (namespace_bindings, inherited_namespaces) =
                    dynamic_namespace_bindings(&element, &elements)?;
                let mut typed_block_owner = classify_typed_block_owner(word, &local, &elements);
                let mut is_typed_inline_owner = false;
                validate_dynamic_content_control(
                    xml,
                    before,
                    word,
                    &local,
                    &namespace_bindings,
                    &mut typed_block_owner,
                    &mut is_typed_inline_owner,
                )?;
                let is_typed_paragraph = typed_block_owner == Some(TypedBlockOwner::Paragraph);
                let paragraph = if is_typed_paragraph {
                    let index = paragraphs.len();
                    paragraphs.push(TocParagraphInsertion {
                        content_start: after,
                        content_end: after,
                    });
                    Some(index)
                } else {
                    elements.last().and_then(|element| element.paragraph)
                };
                mark_typed_sdt_content(&mut elements, &local, typed_block_owner, false);
                elements.push(DynamicXmlElement {
                    local_name: local,
                    qualified_name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    typed_block_owner,
                    is_word: word,
                    is_typed_paragraph,
                    is_typed_inline_owner: false,
                    revision_depth: 0,
                    sdt_content_seen: false,
                    namespace_bindings,
                    inherited_namespaces,
                    run_position: None,
                    hyperlink_plan: None,
                    start: before,
                    start_tag_end: after,
                    paragraph,
                });
            }
            Event::Empty(element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                let typed_block_owner = classify_typed_block_owner(word, local, &elements);
                mark_typed_sdt_content(&mut elements, local, typed_block_owner, false);
                if word
                    && local == b"pPr"
                    && let Some(parent) = elements.last()
                    && parent.is_typed_paragraph
                    && let Some(paragraph) = parent.paragraph
                {
                    paragraphs[paragraph].content_start = after;
                }
            }
            Event::End(element) => {
                let Some(closed) = elements.pop() else {
                    return Err(Error::Other(
                        "table of contents XML has an unmatched end element".to_owned(),
                    ));
                };
                if word
                    && matches_local_name(element.name().as_ref(), b"pPr")
                    && let Some(paragraph) = closed.paragraph
                    && elements
                        .last()
                        .is_some_and(|parent| parent.is_word && parent.local_name == b"p")
                {
                    paragraphs[paragraph].content_start = after;
                }
                if word
                    && matches_local_name(element.name().as_ref(), b"p")
                    && let Some(paragraph) = closed.paragraph
                {
                    paragraphs[paragraph].content_end = before;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(paragraphs)
}

#[derive(Debug)]
struct TocPagePlaceholder {
    toc_index: usize,
    token: String,
    bookmark: String,
}

fn render_toc_entries(
    toc_index: usize,
    toc: &TocField,
    sources: &[TocSource],
    bookmarks: &BTreeMap<usize, TocBookmark>,
    source_xml: &[u8],
    placeholders: &mut Vec<TocPagePlaceholder>,
) -> Result<Vec<u8>> {
    let mut output = String::new();
    for source in sources {
        let bookmark = bookmarks.get(&source.paragraph_index);
        if source.needs_bookmark && bookmark.is_none() {
            return Err(Error::Other(
                "table of contents source bookmark was not allocated".to_owned(),
            ));
        }
        output.push_str("<w:p><w:pPr><w:pStyle w:val=\"TOC");
        output.push_str(&source.level.to_string());
        output.push_str("\"/>");
        if !source.omit_page_number && toc.page_number_separator.is_none() {
            output.push_str(
                "<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/></w:tabs>",
            );
        }
        output.push_str("</w:pPr>");
        if toc.hyperlink {
            output.push_str("<w:hyperlink w:anchor=\"");
            output.push_str(&xml_escape_attribute(
                &bookmark.expect("required above").name,
            ));
            output.push_str("\">");
        }
        output.push_str("<w:r><w:t>");
        output.push_str(&xml_escape_text(&source.title));
        output.push_str("</w:t></w:r>");
        if toc.hyperlink {
            output.push_str("</w:hyperlink>");
        }
        if !source.omit_page_number {
            if let Some(separator) = toc.page_number_separator.as_deref() {
                output.push_str("<w:r><w:t xml:space=\"preserve\">");
                output.push_str(&xml_escape_text(separator));
                output.push_str("</w:t></w:r>");
            } else {
                output.push_str("<w:r><w:tab/></w:r>");
            }
            if let Some(sequence) = source.sequence_prefix.as_deref() {
                output.push_str("<w:r><w:t>");
                output.push_str(&xml_escape_text(sequence));
                output.push_str("</w:t></w:r><w:r><w:t>");
                output.push_str(&xml_escape_text(
                    toc.entry_page_separator.as_deref().unwrap_or("-"),
                ));
                output.push_str("</w:t></w:r>");
            }
            let bookmark = &bookmark.expect("required above").name;
            let token = collision_safe_toc_page_token(source_xml, output.as_bytes(), placeholders);
            output.push_str("<w:fldSimple w:instr=\" PAGEREF ");
            output.push_str(&xml_escape_attribute(bookmark));
            output.push_str(" \\h \" w:dirty=\"0\"><w:r><w:t>");
            output.push_str(&token);
            output.push_str("</w:t></w:r></w:fldSimple>");
            placeholders.push(TocPagePlaceholder {
                toc_index,
                token,
                bookmark: bookmark.clone(),
            });
        }
        output.push_str("</w:p>");
    }
    Ok(output.into_bytes())
}

fn collision_safe_toc_page_token(
    source_xml: &[u8],
    generated: &[u8],
    placeholders: &[TocPagePlaceholder],
) -> String {
    let index = placeholders.len();
    for nonce in 0u64.. {
        let token = format!("__RDOCX_TOC_PAGE_{index}_{nonce}__");
        if find_bytes(source_xml, token.as_bytes()).is_none()
            && find_bytes(generated, token.as_bytes()).is_none()
            && placeholders
                .iter()
                .all(|placeholder| placeholder.token != token)
        {
            return token;
        }
    }
    unreachable!("u64 placeholder nonce space is finite but non-empty")
}

fn xml_escape_text(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn xml_escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn end_boundary_bookmark_starts(
    toc_index: usize,
    span: &DynamicTocSpan,
    spans: &[DynamicTocSpan],
    fields: &[Option<TocField>],
    bookmarks: &BTreeMap<usize, TocBookmark>,
    repairs: &[TocBookmarkRepair],
) -> Vec<u8> {
    let mut output = Vec::new();
    for repair in repairs
        .iter()
        .filter(|repair| repair.toc_index == toc_index && repair.replace_start)
    {
        output.extend_from_slice(
            format!(
                "<w:bookmarkStart w:id=\"{}\" w:name=\"{}\"/>",
                repair.id,
                xml_escape_attribute(&repair.name)
            )
            .as_bytes(),
        );
    }
    let owns_last_end_boundary = fields.get(toc_index).is_some_and(Option::is_some)
        && spans
            .iter()
            .zip(fields)
            .filter(|(candidate, field)| {
                field.is_some() && candidate.end_paragraph == span.end_paragraph
            })
            .all(|(candidate, _)| candidate.end_run_end <= span.end_run_end);
    if owns_last_end_boundary
        && let Some(bookmark) = bookmarks
            .get(&span.end_paragraph)
            .filter(|bookmark| bookmark.insert)
    {
        output.extend_from_slice(
            format!(
                "<w:bookmarkStart w:id=\"{}\" w:name=\"{}\"/>",
                bookmark.id,
                xml_escape_attribute(&bookmark.name)
            )
            .as_bytes(),
        );
    }
    output
}

fn dynamic_toc_replacement(
    xml: &[u8],
    span: &DynamicTocSpan,
    generated: &[u8],
    end_boundary_bookmark_starts: &[u8],
) -> Result<Vec<u8>> {
    if span.result_start > span.end_paragraph_start
        || span.end_paragraph_start > span.end_paragraph_content_start
        || span.end_paragraph_content_start > span.result_end
        || span.result_end > xml.len()
    {
        return Err(Error::Other(
            "table of contents result paragraph boundaries are invalid".to_owned(),
        ));
    }
    let mut prior_wrapper_end = span.end_paragraph_content_start;
    for &(start, end) in &span.end_wrapper_prefixes {
        if start < prior_wrapper_end || start >= end || end > span.result_end {
            return Err(Error::Other(
                "table of contents end-marker wrappers are invalid".to_owned(),
            ));
        }
        prior_wrapper_end = end;
    }
    let mut replacement = Vec::new();
    append_toc_wrapper_closures(&mut replacement, &span.separator_wrapper_names);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(span.start_paragraph_name.as_bytes());
    replacement.push(b'>');
    replacement.extend_from_slice(generated);
    replacement.extend_from_slice(&xml[span.end_paragraph_start..span.end_paragraph_content_start]);
    replacement.extend_from_slice(end_boundary_bookmark_starts);
    for &(start, end) in &span.end_wrapper_prefixes {
        replacement.extend_from_slice(&xml[start..end]);
    }
    Ok(replacement)
}

fn relocate_end_boundary_bookmark_starts(
    mut xml: Vec<u8>,
    spans: &[DynamicTocSpan],
    marker_starts: &[Vec<u8>],
) -> Result<Vec<u8>> {
    let mut edits = Vec::new();
    for (toc_index, span) in spans.iter().enumerate() {
        let markers = marker_starts.get(toc_index).ok_or_else(|| {
            Error::Other("table of contents end bookmark owner was not retained".to_owned())
        })?;
        if markers.is_empty() {
            continue;
        }
        let search_end = span
            .end_wrapper_prefixes
            .first()
            .map(|(start, _)| *start)
            .unwrap_or(span.result_end);
        if span.end_paragraph_content_start > search_end || search_end > span.end_run_end {
            return Err(Error::Other(
                "table of contents end bookmark boundaries are invalid".to_owned(),
            ));
        }
        let matches =
            byte_match_offsets(&xml[span.end_paragraph_content_start..search_end], markers);
        if matches.len() != 1 {
            return Err(Error::Other(
                "table of contents end bookmark start was not uniquely staged".to_owned(),
            ));
        }
        let start = span.end_paragraph_content_start + matches[0];
        let end = start + markers.len();
        if end > span.end_run_end {
            return Err(Error::Other(
                "table of contents end bookmark start crosses its field boundary".to_owned(),
            ));
        }
        let mut replacement = xml[end..span.end_run_end].to_vec();
        replacement.extend_from_slice(markers);
        edits.push(FieldSourceEdit {
            start,
            end: span.end_run_end,
            replacement,
        });
    }
    edits.sort_by_key(|edit| edit.start);
    if edits.windows(2).any(|pair| pair[0].end > pair[1].start) {
        return Err(Error::Other(
            "table of contents end bookmark repairs overlap".to_owned(),
        ));
    }
    for edit in edits.into_iter().rev() {
        xml.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(xml)
}

fn append_toc_wrapper_closures(output: &mut Vec<u8>, wrappers: &[String]) {
    for name in wrappers.iter().rev() {
        output.extend_from_slice(b"</");
        output.extend_from_slice(name.as_bytes());
        output.push(b'>');
    }
}

fn reopen_staged_document(document: Document) -> Result<Document> {
    let mut output = std::io::Cursor::new(Vec::new());
    document.package.write_to(&mut output)?;
    Document::from_bytes(output.get_ref())
}

fn deterministic_toc_page_values(document: &Document) -> Result<HashMap<String, String>> {
    let layout = document.layout_deterministic()?;
    let mut output = HashMap::new();
    let mut invalid_target = None;
    for page in &layout.layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Text(run) = element
                && let Some(oxml_layout::FieldKind::TargetPage(target)) = run.field_kind
                && let Some(name) = layout.page_reference_name(target)
            {
                let valid_page = run
                    .text
                    .parse::<usize>()
                    .is_ok_and(|value| (1..=layout.layout.pages.len()).contains(&value));
                if !valid_page
                    || output
                        .get(name)
                        .is_some_and(|existing| existing != &run.text)
                {
                    invalid_target.get_or_insert_with(|| name.to_owned());
                } else {
                    output
                        .entry(name.to_owned())
                        .or_insert_with(|| run.text.clone());
                }
            }
        });
    }
    if let Some(target) = invalid_target {
        return Err(Error::Other(format!(
            "table of contents page target {target} was not resolved"
        )));
    }
    Ok(output)
}

fn byte_match_offsets(input: &[u8], needle: &[u8]) -> Vec<usize> {
    if needle.is_empty() {
        return Vec::new();
    }
    let mut output = Vec::new();
    let mut start = 0usize;
    while let Some(relative) = find_bytes(&input[start..], needle) {
        let found = start + relative;
        output.push(found);
        start = found + needle.len();
    }
    output
}

fn non_body_story_parts(document: &Document) -> Result<Vec<(String, Vec<u8>)>> {
    let mut parts = merge_referenced_header_footer_parts(document)?
        .into_iter()
        .chain(
            relationship_parts(document, rel_types::FOOTNOTES)
                .into_iter()
                .map(|(name, xml)| (format!("footnotes:{name}"), xml)),
        )
        .chain(
            relationship_parts(document, rel_types::ENDNOTES)
                .into_iter()
                .map(|(name, xml)| (format!("endnotes:{name}"), xml)),
        )
        .collect::<Vec<_>>();
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    Ok(parts)
}

fn merge_referenced_header_footer_parts(document: &Document) -> Result<Vec<(String, Vec<u8>)>> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Ok(Vec::new());
    };
    let xml = document.document.to_xml()?;
    let mut reader = NsReader::from_reader(xml.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut references = Vec::<(String, bool)>::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid main document XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if namespace_is_word(&reader.resolver().resolve_element(element.name()).0) =>
            {
                let name = element.name();
                let local = local_name(name.as_ref());
                let is_header = if local == b"headerReference" {
                    true
                } else if local == b"footerReference" {
                    false
                } else {
                    buffer.clear();
                    continue;
                };
                if let Some((_, rel_id)) = resolved_element_attribute(
                    &element,
                    reader.resolver(),
                    b"id",
                    AttributeNamespace::Relationship,
                )? {
                    references.push((rel_id, is_header));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut seen = HashSet::new();
    let mut parts = Vec::new();
    for (rel_id, is_header) in references {
        let Some(relationship) = relationships.get_by_id(&rel_id) else {
            continue;
        };
        let relationship_type = if is_header {
            rel_types::HEADER
        } else {
            rel_types::FOOTER
        };
        if relationship.rel_type != relationship_type
            || relationship.target_mode.as_deref() == Some("External")
        {
            continue;
        }
        let part_name =
            OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
        if seen.insert(part_name.clone())
            && let Some(xml) = document.package.get_part(&part_name)
        {
            let story = if is_header { "header" } else { "footer" };
            parts.push((format!("{story}:{part_name}"), xml.to_vec()));
        }
    }
    Ok(parts)
}

fn reject_varying_non_body_merge_fields(
    document: &Document,
    records: &[BTreeMap<String, String>],
) -> Result<()> {
    if records.is_empty() {
        return Ok(());
    }
    let mut names = HashSet::new();
    for (_, xml) in non_body_story_parts(document)? {
        collect_raw_merge_field_names(&xml, &mut names)?;
    }
    for name in names {
        let expected = records[0].get(&name).map(String::as_str).unwrap_or("");
        if records
            .iter()
            .skip(1)
            .any(|record| record.get(&name).map(String::as_str).unwrap_or("") != expected)
        {
            return Err(Error::Other(
                "sectioned mail merge cannot vary fields in headers, footers, footnotes, or endnotes"
                    .to_owned(),
            ));
        }
    }
    Ok(())
}

struct ComplexInstruction {
    text: String,
    collecting: bool,
}

fn collect_raw_merge_field_names(xml: &[u8], names: &mut HashSet<String>) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut complex = Vec::<ComplexInstruction>::new();
    let mut in_instruction_text = false;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package story XML: {error}")))?;
        let word = match &event {
            Event::Start(element) | Event::Empty(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            Event::End(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            _ => false,
        };
        match event {
            Event::Start(element) if word => {
                if collect_raw_field_element(&element, reader.resolver(), names, &mut complex)? {
                    in_instruction_text = true;
                }
            }
            Event::Empty(element) if word => {
                collect_raw_field_element(&element, reader.resolver(), names, &mut complex)?;
            }
            Event::Text(text) if in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    let decoded = text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::Other(format!("invalid field instruction entity: {error}"))
                    })?;
                    instruction.text.push_str(&unescaped);
                }
            }
            Event::CData(text) if in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    instruction.text.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?);
                }
            }
            Event::End(element)
                if word && matches_local_name(element.name().as_ref(), b"instrText") =>
            {
                in_instruction_text = false;
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_raw_field_element(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    names: &mut HashSet<String>,
    complex: &mut Vec<ComplexInstruction>,
) -> Result<bool> {
    let name = element.name();
    let local = local_name(name.as_ref());
    if local == b"fldSimple" {
        if let Some((_, instruction)) =
            resolved_element_attribute(element, resolver, b"instr", AttributeNamespace::Word)?
        {
            collect_merge_field_name(&instruction, names);
        }
    } else if local == b"fldChar" {
        match resolved_element_attribute(
            element,
            resolver,
            b"fldCharType",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value)
        .as_deref()
        {
            Some("begin") => complex.push(ComplexInstruction {
                text: String::new(),
                collecting: true,
            }),
            Some("separate") => {
                if let Some(instruction) = complex.last_mut() {
                    instruction.collecting = false;
                }
            }
            Some("end") => {
                if let Some(instruction) = complex.pop() {
                    collect_merge_field_name(&instruction.text, names);
                }
            }
            _ => {}
        }
    }
    Ok(local == b"instrText")
}

fn collect_merge_field_name(instruction: &str, names: &mut HashSet<String>) {
    let field = Field::new(instruction, "");
    if field.instruction.name == "MERGEFIELD"
        && let Some(FieldArgument::Text(name)) = field.instruction.arguments.first()
    {
        names.insert(name.clone());
    }
}

#[derive(Default)]
struct BodyIdentityValues {
    numeric_ids: Vec<String>,
    bookmark_names: Vec<String>,
    reference_names: Vec<String>,
}

struct BodyIdentityState {
    used_ids: HashSet<u32>,
    used_names: HashSet<String>,
    next_id: u32,
    next_name: u32,
}

impl BodyIdentityState {
    fn from_documents(documents: &[Document]) -> Result<Self> {
        let mut state = Self {
            used_ids: HashSet::new(),
            used_names: HashSet::new(),
            next_id: 1,
            next_name: 1,
        };
        for document in documents {
            let xml = document.document.to_xml()?;
            let values = body_identity_values(&xml)?;
            state.used_ids.extend(
                values
                    .numeric_ids
                    .iter()
                    .filter_map(|value| value.parse::<u32>().ok()),
            );
            state.used_names.extend(values.bookmark_names);
            state.used_names.extend(values.reference_names);
        }
        Ok(state)
    }

    fn allocate_id(&mut self) -> Result<String> {
        loop {
            let candidate = self.next_id;
            self.next_id = self.next_id.checked_add(1).ok_or_else(|| {
                Error::Other("mail merge exhausted the document identity range".to_owned())
            })?;
            if self.used_ids.insert(candidate) {
                return Ok(candidate.to_string());
            }
        }
    }

    fn allocate_name(&mut self) -> Result<String> {
        loop {
            let candidate = format!("MailMerge{}", self.next_name);
            self.next_name = self.next_name.checked_add(1).ok_or_else(|| {
                Error::Other("mail merge exhausted the bookmark name range".to_owned())
            })?;
            if self.used_names.insert(candidate.clone()) {
                return Ok(candidate);
            }
        }
    }
}

#[derive(Default)]
struct BodyIdentityRemap {
    numeric_ids: BTreeMap<String, String>,
    bookmark_names: BTreeMap<String, String>,
}

fn remap_body_identities(document: &mut Document, state: &mut BodyIdentityState) -> Result<()> {
    let xml = document.document.to_xml()?;
    let values = body_identity_values(&xml)?;
    let mut remap = BodyIdentityRemap::default();
    for value in values.numeric_ids {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.numeric_ids.entry(value) {
            entry.insert(state.allocate_id()?);
        }
    }
    for value in values.bookmark_names {
        if let std::collections::btree_map::Entry::Vacant(entry) = remap.bookmark_names.entry(value)
        {
            entry.insert(state.allocate_name()?);
        }
    }
    let updated = patch_body_identity_attributes(&xml, &remap)?;
    document.document = CT_Document::from_xml(&updated)?;
    Ok(())
}

fn body_identity_values(xml: &[u8]) -> Result<BodyIdentityValues> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut body_depth = None;
    let mut sdt_properties_depth = None;
    let mut values = BodyIdentityValues::default();
    let mut complex = Vec::<ComplexInstruction>::new();
    let mut in_instruction_text = false;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid document identity XML: {error}")))?;
        let word = match &event {
            Event::Start(element) | Event::Empty(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            Event::End(element) => {
                namespace_is_word(&reader.resolver().resolve_element(element.name()).0)
            }
            _ => false,
        };
        match event {
            Event::Start(element) => {
                if body_depth.is_none()
                    && word
                    && matches_local_name(element.name().as_ref(), b"body")
                {
                    body_depth = Some(depth);
                } else if body_depth.is_some() {
                    collect_body_identity_values(
                        &element,
                        reader.resolver(),
                        sdt_properties_depth.is_some(),
                        &mut values,
                    )?;
                    if collect_body_reference_values(
                        &element,
                        reader.resolver(),
                        &mut complex,
                        &mut values.reference_names,
                    )? {
                        in_instruction_text = true;
                    }
                    if word && matches_local_name(element.name().as_ref(), b"sdtPr") {
                        sdt_properties_depth = Some(depth);
                    }
                }
                depth += 1;
            }
            Event::Empty(element) if body_depth.is_some() => {
                collect_body_identity_values(
                    &element,
                    reader.resolver(),
                    sdt_properties_depth.is_some(),
                    &mut values,
                )?;
                collect_body_reference_values(
                    &element,
                    reader.resolver(),
                    &mut complex,
                    &mut values.reference_names,
                )?;
            }
            Event::Text(text) if body_depth.is_some() && in_instruction_text => {
                append_complex_instruction_text(
                    &text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?,
                    &mut complex,
                )?;
            }
            Event::CData(text) if body_depth.is_some() && in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    instruction.text.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?);
                }
            }
            Event::End(element) => {
                if word && matches_local_name(element.name().as_ref(), b"instrText") {
                    in_instruction_text = false;
                }
                depth = depth.saturating_sub(1);
                if sdt_properties_depth == Some(depth) {
                    sdt_properties_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
            }
            Event::Eof => return Ok(values),
            _ => {}
        }
        buffer.clear();
    }
}

fn collect_body_identity_values(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    in_sdt_properties: bool,
    values: &mut BodyIdentityValues,
) -> Result<()> {
    let name = element.name();
    let local = local_name(name.as_ref());
    let namespace = resolver.resolve_element(element.name()).0;
    if namespace_is_word(&namespace) && matches!(local, b"bookmarkStart" | b"bookmarkEnd") {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Word)?
        {
            values.numeric_ids.push(value);
        }
        if local == b"bookmarkStart"
            && let Some((_, value)) =
                resolved_element_attribute(element, resolver, b"name", AttributeNamespace::Word)?
        {
            values.bookmark_names.push(value);
        }
    } else if namespace_is_word(&namespace) && in_sdt_properties && local == b"id" {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"val", AttributeNamespace::Word)?
        {
            values.numeric_ids.push(value);
        }
    } else if namespace_matches(&namespace, WP_NS) && local == b"docPr" {
        if let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Unbound)?
        {
            values.numeric_ids.push(value);
        }
    } else if namespace_is_non_visual_drawing(&namespace)
        && local == b"cNvPr"
        && let Some((_, value)) =
            resolved_element_attribute(element, resolver, b"id", AttributeNamespace::Unbound)?
    {
        values.numeric_ids.push(value);
    }
    Ok(())
}

fn collect_body_reference_values(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    complex: &mut Vec<ComplexInstruction>,
    references: &mut Vec<String>,
) -> Result<bool> {
    let (namespace, local) = resolver.resolve_element(element.name());
    if !namespace_is_word(&namespace) {
        return Ok(false);
    }
    if local.as_ref() == b"hyperlink" {
        if let Some((_, anchor)) =
            resolved_element_attribute(element, resolver, b"anchor", AttributeNamespace::Word)?
        {
            references.push(anchor);
        }
    } else if local.as_ref() == b"fldSimple" {
        if let Some((_, instruction)) =
            resolved_element_attribute(element, resolver, b"instr", AttributeNamespace::Word)?
        {
            collect_reference_field_name(&instruction, references);
        }
    } else if local.as_ref() == b"fldChar" {
        match resolved_element_attribute(
            element,
            resolver,
            b"fldCharType",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value)
        .as_deref()
        {
            Some("begin") => complex.push(ComplexInstruction {
                text: String::new(),
                collecting: true,
            }),
            Some("separate") => {
                if let Some(instruction) = complex.last_mut() {
                    instruction.collecting = false;
                }
            }
            Some("end") => {
                if let Some(instruction) = complex.pop() {
                    collect_reference_field_name(&instruction.text, references);
                }
            }
            _ => {}
        }
    }
    Ok(local.as_ref() == b"instrText")
}

fn append_complex_instruction_text(text: &str, complex: &mut [ComplexInstruction]) -> Result<()> {
    if let Some(instruction) = complex.last_mut()
        && instruction.collecting
    {
        let unescaped = quick_xml::escape::unescape(text)
            .map_err(|error| Error::Other(format!("invalid field instruction entity: {error}")))?;
        instruction.text.push_str(&unescaped);
    }
    Ok(())
}

fn collect_reference_field_name(instruction: &str, references: &mut Vec<String>) {
    if let Some(name) = reference_field_name(instruction) {
        references.push(name);
    }
}

fn reference_field_name(instruction: &str) -> Option<String> {
    let field = Field::new(instruction, "");
    if matches!(field.instruction.name.as_str(), "REF" | "PAGEREF") {
        field
            .instruction
            .arguments
            .first()
            .and_then(argument_text)
            .map(str::to_owned)
    } else {
        None
    }
}

#[derive(Clone, Copy)]
enum AttributeNamespace {
    Relationship,
    Word,
    Unbound,
}

fn resolved_element_attribute(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    local: &[u8],
    expected: AttributeNamespace,
) -> Result<Option<(Vec<u8>, String)>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| Error::Other(format!("invalid XML attribute: {error}")))?;
        let (namespace, resolved_local) = resolver.resolve_attribute(attribute.key);
        let namespace_matches = match expected {
            AttributeNamespace::Relationship => namespace_matches(&namespace, R_NS),
            AttributeNamespace::Word => namespace_is_word(&namespace),
            AttributeNamespace::Unbound => matches!(namespace, ResolveResult::Unbound),
        };
        if namespace_matches && resolved_local.as_ref() == local {
            let raw = std::str::from_utf8(attribute.value.as_ref())
                .map_err(|error| Error::Other(format!("invalid XML attribute value: {error}")))?;
            let decoded = quick_xml::escape::unescape(raw)
                .map_err(|error| Error::Other(format!("invalid XML attribute entity: {error}")))?;
            return Ok(Some((
                attribute.key.as_ref().to_vec(),
                decoded.into_owned(),
            )));
        }
    }
    Ok(None)
}

const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

fn namespace_matches(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected.as_bytes())
}

fn namespace_is_non_visual_drawing(namespace: &ResolveResult<'_>) -> bool {
    [PIC_NS, DRAWING_NS, WPS_NS]
        .iter()
        .any(|expected| namespace_matches(namespace, expected))
}

fn patch_body_identity_attributes(xml: &[u8], remap: &BodyIdentityRemap) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut body_depth = None;
    let mut sdt_properties_depth = None;
    let mut edits = Vec::new();
    let mut complex = Vec::<ComplexInstructionEdit>::new();
    let mut in_instruction_text = false;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid document identity XML: {error}")))?;
        let after = reader.buffer_position() as usize;
        let identity_namespace = match &event {
            Event::Start(element) | Event::Empty(element) => BodyIdentityNamespace::from_resolved(
                &reader.resolver().resolve_element(element.name()).0,
            ),
            Event::End(element) => BodyIdentityNamespace::from_resolved(
                &reader.resolver().resolve_element(element.name()).0,
            ),
            _ => BodyIdentityNamespace::default(),
        };
        match event {
            Event::Start(element) => {
                if body_depth.is_none()
                    && identity_namespace.word
                    && matches_local_name(element.name().as_ref(), b"body")
                {
                    body_depth = Some(depth);
                } else if body_depth.is_some() {
                    collect_body_identity_edits(
                        xml,
                        before,
                        after,
                        &element,
                        reader.resolver(),
                        identity_namespace,
                        sdt_properties_depth.is_some(),
                        remap,
                        &mut edits,
                    )?;
                    if collect_body_reference_edits(
                        xml,
                        before,
                        after,
                        &element,
                        reader.resolver(),
                        remap,
                        &mut complex,
                        &mut edits,
                    )? {
                        in_instruction_text = true;
                    }
                    if identity_namespace.word
                        && matches_local_name(element.name().as_ref(), b"sdtPr")
                    {
                        sdt_properties_depth = Some(depth);
                    }
                }
                depth += 1;
            }
            Event::Empty(element) if body_depth.is_some() => {
                collect_body_identity_edits(
                    xml,
                    before,
                    after,
                    &element,
                    reader.resolver(),
                    identity_namespace,
                    sdt_properties_depth.is_some(),
                    remap,
                    &mut edits,
                )?;
                collect_body_reference_edits(
                    xml,
                    before,
                    after,
                    &element,
                    reader.resolver(),
                    remap,
                    &mut complex,
                    &mut edits,
                )?;
            }
            Event::Text(text) if body_depth.is_some() && in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    let decoded = text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::Other(format!("invalid field instruction entity: {error}"))
                    })?;
                    instruction.text.push_str(&decoded);
                    instruction.spans.push((before, after));
                }
            }
            Event::CData(text) if body_depth.is_some() && in_instruction_text => {
                if let Some(instruction) = complex.last_mut()
                    && instruction.collecting
                {
                    instruction.text.push_str(&text.decode().map_err(|error| {
                        Error::Other(format!("invalid field instruction text: {error}"))
                    })?);
                    instruction.spans.push((before, after));
                }
            }
            Event::End(element) => {
                if identity_namespace.word
                    && matches_local_name(element.name().as_ref(), b"instrText")
                {
                    in_instruction_text = false;
                }
                depth = depth.saturating_sub(1);
                if sdt_properties_depth == Some(depth) {
                    sdt_properties_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let mut updated = xml.to_vec();
    edits.sort_by_key(|edit: &FieldSourceEdit| edit.start);
    for edit in edits.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

#[allow(clippy::too_many_arguments)]
fn collect_body_identity_edits(
    xml: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    namespace: BodyIdentityNamespace,
    in_sdt_properties: bool,
    remap: &BodyIdentityRemap,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<()> {
    let name = element.name();
    let local = local_name(name.as_ref());
    if namespace.word && matches!(local, b"bookmarkStart" | b"bookmarkEnd") {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"id",
            AttributeNamespace::Word,
            &remap.numeric_ids,
            edits,
        )?;
        if local == b"bookmarkStart" {
            add_identity_attribute_edit(
                xml,
                start,
                end,
                element,
                resolver,
                b"name",
                AttributeNamespace::Word,
                &remap.bookmark_names,
                edits,
            )?;
        }
    } else if namespace.word && in_sdt_properties && local == b"id" {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"val",
            AttributeNamespace::Word,
            &remap.numeric_ids,
            edits,
        )?;
    } else if namespace.wordprocessing_drawing && local == b"docPr"
        || namespace.non_visual_drawing && local == b"cNvPr"
    {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"id",
            AttributeNamespace::Unbound,
            &remap.numeric_ids,
            edits,
        )?;
    }
    Ok(())
}

#[derive(Clone, Copy, Default)]
struct BodyIdentityNamespace {
    word: bool,
    wordprocessing_drawing: bool,
    non_visual_drawing: bool,
}

impl BodyIdentityNamespace {
    fn from_resolved(namespace: &ResolveResult<'_>) -> Self {
        Self {
            word: namespace_is_word(namespace),
            wordprocessing_drawing: namespace_matches(namespace, WP_NS),
            non_visual_drawing: namespace_is_non_visual_drawing(namespace),
        }
    }
}

fn add_identity_attribute_edit(
    xml: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    local: &[u8],
    expected: AttributeNamespace,
    replacements: &BTreeMap<String, String>,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<()> {
    let Some((key, old)) = resolved_element_attribute(element, resolver, local, expected)? else {
        return Ok(());
    };
    let Some(replacement) = replacements.get(&old) else {
        return Ok(());
    };
    let Some((relative_start, relative_end)) = attribute_value_span(&xml[start..end], &key) else {
        return Err(Error::Other(
            "document identity attribute source was not found".to_owned(),
        ));
    };
    edits.push(FieldSourceEdit {
        start: start + relative_start,
        end: start + relative_end,
        replacement: replacement.as_bytes().to_vec(),
    });
    Ok(())
}

struct ComplexInstructionEdit {
    text: String,
    collecting: bool,
    spans: Vec<(usize, usize)>,
}

#[allow(clippy::too_many_arguments)]
fn collect_body_reference_edits(
    xml: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    remap: &BodyIdentityRemap,
    complex: &mut Vec<ComplexInstructionEdit>,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<bool> {
    let (namespace, local) = resolver.resolve_element(element.name());
    if !namespace_is_word(&namespace) {
        return Ok(false);
    }
    if local.as_ref() == b"hyperlink" {
        add_identity_attribute_edit(
            xml,
            start,
            end,
            element,
            resolver,
            b"anchor",
            AttributeNamespace::Word,
            &remap.bookmark_names,
            edits,
        )?;
    } else if local.as_ref() == b"fldSimple" {
        if let Some((key, instruction)) =
            resolved_element_attribute(element, resolver, b"instr", AttributeNamespace::Word)?
            && let Some(updated) = remap_reference_instruction(&instruction, &remap.bookmark_names)
        {
            add_attribute_value_edit(xml, start, end, &key, &updated, true, edits)?;
        }
    } else if local.as_ref() == b"fldChar" {
        match resolved_element_attribute(
            element,
            resolver,
            b"fldCharType",
            AttributeNamespace::Word,
        )?
        .map(|(_, value)| value)
        .as_deref()
        {
            Some("begin") => complex.push(ComplexInstructionEdit {
                text: String::new(),
                collecting: true,
                spans: Vec::new(),
            }),
            Some("separate") => {
                if let Some(instruction) = complex.last_mut() {
                    instruction.collecting = false;
                }
            }
            Some("end") => {
                if let Some(instruction) = complex.pop()
                    && let Some(updated) =
                        remap_reference_instruction(&instruction.text, &remap.bookmark_names)
                    && let Some((first, remaining)) = instruction.spans.split_first()
                {
                    edits.push(FieldSourceEdit {
                        start: first.0,
                        end: first.1,
                        replacement: quick_xml::escape::escape(&updated)
                            .into_owned()
                            .into_bytes(),
                    });
                    edits.extend(remaining.iter().map(|(start, end)| FieldSourceEdit {
                        start: *start,
                        end: *end,
                        replacement: Vec::new(),
                    }));
                }
            }
            _ => {}
        }
    }
    Ok(local.as_ref() == b"instrText")
}

fn add_attribute_value_edit(
    xml: &[u8],
    start: usize,
    end: usize,
    key: &[u8],
    replacement: &str,
    escape: bool,
    edits: &mut Vec<FieldSourceEdit>,
) -> Result<()> {
    let Some((relative_start, relative_end)) = attribute_value_span(&xml[start..end], key) else {
        return Err(Error::Other(
            "document identity attribute source was not found".to_owned(),
        ));
    };
    let replacement = if escape {
        quick_xml::escape::escape(replacement)
            .into_owned()
            .into_bytes()
    } else {
        replacement.as_bytes().to_vec()
    };
    edits.push(FieldSourceEdit {
        start: start + relative_start,
        end: start + relative_end,
        replacement,
    });
    Ok(())
}

fn remap_reference_instruction(
    instruction: &str,
    names: &BTreeMap<String, String>,
) -> Option<String> {
    let old = reference_field_name(instruction)?;
    let replacement = names.get(&old)?;
    let command_end = instruction
        .find(char::is_whitespace)
        .unwrap_or(instruction.len());
    let relative = instruction[command_end..].find(&old)?;
    let start = command_end + relative;
    let mut updated = instruction.to_owned();
    updated.replace_range(start..start + old.len(), replacement);
    Some(updated)
}

fn attribute_value_span(element: &[u8], attribute_name: &[u8]) -> Option<(usize, usize)> {
    let mut index = 1usize;
    while index < element.len() && !element[index].is_ascii_whitespace() {
        index += 1;
    }
    while index < element.len() {
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= element.len() || matches!(element[index], b'>' | b'/') {
            return None;
        }
        let name_start = index;
        while index < element.len()
            && !element[index].is_ascii_whitespace()
            && element[index] != b'='
        {
            index += 1;
        }
        let name_end = index;
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        if element.get(index) != Some(&b'=') {
            return None;
        }
        index += 1;
        while index < element.len() && element[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *element.get(index)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        index += 1;
        let value_start = index;
        while index < element.len() && element[index] != quote {
            index += 1;
        }
        let value_end = index;
        if &element[name_start..name_end] == attribute_name {
            return Some((value_start, value_end));
        }
        index += 1;
    }
    None
}

fn empty_section_properties() -> CT_SectPr {
    CT_SectPr {
        page_width: None,
        page_height: None,
        orientation: None,
        margin_top: None,
        margin_right: None,
        margin_bottom: None,
        margin_left: None,
        gutter: None,
        header_distance: None,
        footer_distance: None,
        section_type: None,
        columns: None,
        title_pg: None,
        header_refs: Vec::new(),
        footer_refs: Vec::new(),
        extra_xml: Vec::new(),
        change: None,
    }
}

struct CachedFieldUpdate {
    cached_result: String,
    dirty: bool,
}

fn valid_xml_character(value: char) -> bool {
    matches!(value, '\u{0009}' | '\u{000A}' | '\u{000D}')
        || ('\u{0020}'..='\u{D7FF}').contains(&value)
        || ('\u{E000}'..='\u{FFFD}').contains(&value)
        || ('\u{10000}'..='\u{10FFFF}').contains(&value)
}

#[derive(Debug, Default)]
struct SequenceState {
    value: Option<i64>,
    heading_anchor: Option<usize>,
}

#[derive(Debug, Clone, Copy)]
struct MailMergeStoryState {
    record_number: Option<u32>,
    sequence_number: Option<u32>,
}

struct Evaluator<'a> {
    document: &'a Document,
    context: &'a FieldEvaluationContext,
    bookmarks: BTreeMap<String, String>,
    results: Vec<FieldEvaluation>,
    sequences: BTreeMap<(String, String), SequenceState>,
    mail_merge_stories: BTreeMap<String, MailMergeStoryState>,
    nested_outcomes: Vec<BTreeMap<usize, FieldOutcome>>,
    missing_merge_fields_as_empty: bool,
}

impl<'a> Evaluator<'a> {
    fn new(document: &'a Document, context: &'a FieldEvaluationContext) -> Self {
        let bookmarks = document
            .bookmarks()
            .into_iter()
            .filter(|bookmark| bookmark.issue().is_none())
            .filter_map(|bookmark| Some((bookmark.name()?.to_owned(), bookmark.text().to_owned())))
            .collect();
        Self {
            document,
            context,
            bookmarks,
            results: Vec::new(),
            sequences: BTreeMap::new(),
            mail_merge_stories: BTreeMap::new(),
            nested_outcomes: Vec::new(),
            missing_merge_fields_as_empty: false,
        }
    }

    fn for_mail_merge(document: &'a Document, context: &'a FieldEvaluationContext) -> Self {
        let mut evaluator = Self::new(document, context);
        evaluator.missing_merge_fields_as_empty = true;
        evaluator
    }

    fn evaluate_story(&mut self, story: &str, paragraphs: &[&CT_P]) {
        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            for run in paragraph.runs() {
                for content in &run.content {
                    if let RunContent::Field(field) = content {
                        self.evaluate_field(field, story, paragraphs, paragraph_index);
                    }
                }
            }
        }
    }

    fn evaluate_field(
        &mut self,
        field: &Field,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let instruction = field.effective_instruction();
        let result_index = self.results.len();
        self.results.push(FieldEvaluation {
            field_index: result_index,
            instruction: instruction.raw.clone(),
            cached_result: field.cached_result.clone(),
            outcome: keep("field evaluation did not complete"),
        });

        self.nested_outcomes.push(BTreeMap::new());
        self.evaluate_nested_fields(field, &instruction, story, paragraphs, paragraph_index);
        let outcome = self.evaluate_instruction(&instruction, story, paragraphs, paragraph_index);
        self.nested_outcomes.pop();
        self.results[result_index].outcome = outcome.clone();
        outcome
    }

    fn evaluate_instruction(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        if instruction.name.is_empty() {
            return keep("field instruction has no name");
        }
        if let Some(name) = unsupported_switch(instruction) {
            return keep(&format!(
                "field {} uses unsupported switch \\{name}",
                instruction.name
            ));
        }
        if let Err(diagnostic) = validate_instruction_shape(instruction) {
            return keep(&diagnostic);
        }

        let outcome = match instruction.name.as_str() {
            "PAGE" | "NUMPAGES" => FieldOutcome::DeferredPagination,
            "PAGEREF" => self.evaluate_pageref(instruction),
            "REF" => self.evaluate_ref(instruction),
            "IF" => self.evaluate_if(instruction, story, paragraphs, paragraph_index),
            "SEQ" => self.evaluate_seq(instruction, story, paragraphs, paragraph_index),
            "DOCPROPERTY" => self.evaluate_docproperty(instruction),
            "DOCVARIABLE" => self.evaluate_docvariable(instruction),
            "STYLEREF" => self.evaluate_styleref(instruction, paragraphs, paragraph_index),
            "INCLUDETEXT" => self.evaluate_includetext(instruction),
            "DATE" | "TIME" => self.evaluate_date_time(instruction),
            "FILENAME" => self.evaluate_filename(instruction),
            "AUTHOR" => self.evaluate_author(),
            "MERGEFIELD" => self.evaluate_mergefield(instruction),
            "=" => self.evaluate_formula(instruction, story, paragraphs, paragraph_index),
            "TOC" => self.evaluate_toc(instruction, story, paragraphs, paragraph_index),
            "TC" => self.evaluate_tc(instruction, story, paragraphs, paragraph_index),
            "NEXT" | "NEXTIF" | "SKIPIF" | "MERGEREC" | "MERGESEQ" => {
                self.evaluate_mail_merge_control(instruction, story, paragraphs, paragraph_index)
            }
            "DISPLAYBARCODE" | "MERGEBARCODE" => {
                self.evaluate_barcode(instruction, story, paragraphs, paragraph_index)
            }
            name => keep(&format!("field {name} is unsupported")),
        };

        match outcome {
            FieldOutcome::Resolved(value) => {
                match apply_formats(instruction, &value, self.context.now) {
                    Ok(value) => FieldOutcome::Resolved(value),
                    Err(diagnostic) => keep(&diagnostic),
                }
            }
            other => other,
        }
    }

    fn evaluate_nested_fields(
        &mut self,
        field: &Field,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) {
        for nested in field.effective_nested_fields_in_source_order(instruction) {
            let key = std::ptr::from_ref(nested) as usize;
            if !self
                .nested_outcomes
                .last()
                .is_some_and(|outcomes| outcomes.contains_key(&key))
            {
                let outcome = self.evaluate_field(nested, story, paragraphs, paragraph_index);
                self.nested_outcomes
                    .last_mut()
                    .expect("field evaluation frame exists")
                    .insert(key, outcome);
            }
        }
    }

    fn evaluate_ref(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(target) = text_argument(instruction, 0) else {
            return keep("REF requires a bookmark name");
        };
        match self.bookmarks.get(target) {
            Some(text) => FieldOutcome::Resolved(text.clone()),
            None => keep(&format!("REF target {target} was not found")),
        }
    }

    fn evaluate_pageref(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(target) = text_argument(instruction, 0) else {
            return keep("PAGEREF requires a bookmark name");
        };
        if self.bookmarks.contains_key(target) {
            FieldOutcome::DeferredPagination
        } else {
            keep(&format!("PAGEREF target {target} was not found"))
        }
    }

    fn evaluate_if(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        if instruction.arguments.len() != 5 {
            return keep("IF requires two operands, an operator, and two results");
        }
        let arguments = instruction
            .arguments
            .iter()
            .map(|argument| self.resolve_argument(argument, story, paragraphs, paragraph_index))
            .collect::<Vec<_>>();
        let left = match &arguments[0] {
            Ok(value) => value,
            Err(diagnostic) => return keep(diagnostic),
        };
        let operator = match &arguments[1] {
            Ok(value) => value,
            Err(diagnostic) => return keep(diagnostic),
        };
        let right = match &arguments[2] {
            Ok(value) => value,
            Err(diagnostic) => return keep(diagnostic),
        };
        let Some(condition) = compare_if(left, operator, right) else {
            return keep(&format!("IF operator {operator} is unsupported"));
        };
        let selected = if condition { 3 } else { 4 };
        match &arguments[selected] {
            Ok(value) => FieldOutcome::Resolved(value.clone()),
            Err(diagnostic) => keep(diagnostic),
        }
    }

    fn resolve_argument(
        &mut self,
        argument: &FieldArgument,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> std::result::Result<String, String> {
        match argument {
            FieldArgument::Text(value) => Ok(value.clone()),
            FieldArgument::Nested(field) => {
                let key = std::ptr::from_ref(field.as_ref()) as usize;
                let outcome = if let Some(outcome) = self
                    .nested_outcomes
                    .last()
                    .and_then(|outcomes| outcomes.get(&key))
                {
                    outcome.clone()
                } else {
                    let outcome = self.evaluate_field(field, story, paragraphs, paragraph_index);
                    self.nested_outcomes
                        .last_mut()
                        .expect("field evaluation frame exists")
                        .insert(key, outcome.clone());
                    outcome
                };
                match outcome {
                    FieldOutcome::Resolved(value) => Ok(value),
                    FieldOutcome::DeferredPagination => {
                        Err("nested field requires deferred pagination".to_owned())
                    }
                    FieldOutcome::TableOfContents(_)
                    | FieldOutcome::TableOfContentsEntry(_)
                    | FieldOutcome::MailMergeControl(_)
                    | FieldOutcome::Barcode(_) => {
                        Err("nested field produced a non-text result".to_owned())
                    }
                    FieldOutcome::KeepStored { diagnostic } => {
                        Err(format!("nested field was not resolved: {diagnostic}"))
                    }
                }
            }
        }
    }

    fn evaluate_seq(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(identifier) = text_argument(instruction, 0) else {
            return keep("SEQ requires an identifier");
        };
        let heading_anchor = switch_text(instruction, "s").and_then(|level| {
            let level = level.parse::<u32>().ok()?.checked_sub(1)?;
            paragraphs[..=paragraph_index]
                .iter()
                .enumerate()
                .rev()
                .find(|(_, paragraph)| {
                    let style_id = paragraph
                        .properties
                        .as_ref()
                        .and_then(|properties| properties.style_id.as_deref());
                    let mut effective =
                        style::resolve_paragraph_properties(style_id, &self.document.styles);
                    if let Some(properties) = &paragraph.properties {
                        effective.merge_from(properties);
                    }
                    effective.outline_lvl == Some(level)
                })
                .map(|(index, _)| index)
        });
        if has_switch(instruction, "s") && heading_anchor.is_none() {
            return keep("SEQ heading restart has no matching heading");
        }

        let state = self
            .sequences
            .entry((story.to_owned(), identifier.to_ascii_lowercase()))
            .or_default();
        if heading_anchor.is_some() && state.heading_anchor != heading_anchor {
            state.value = None;
            state.heading_anchor = heading_anchor;
        }

        let value = if let Some(reset) = switch_text(instruction, "r") {
            let Ok(reset) = reset.parse::<i64>() else {
                return keep("SEQ reset value is not an integer");
            };
            state.value = Some(reset);
            reset
        } else if has_switch(instruction, "c") {
            let Some(value) = state.value else {
                return keep("SEQ repeat has no preceding value");
            };
            value
        } else {
            let Some(value) = state.value.unwrap_or(0).checked_add(1) else {
                return keep("SEQ value overflowed");
            };
            state.value = Some(value);
            value
        };
        if has_switch(instruction, "h") {
            FieldOutcome::Resolved(String::new())
        } else {
            FieldOutcome::Resolved(value.to_string())
        }
    }

    fn evaluate_docproperty(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(name) = text_argument(instruction, 0) else {
            return keep("DOCPROPERTY requires a property name");
        };
        if let Some(value) = self.core_property(name) {
            return FieldOutcome::Resolved(value.to_owned());
        }
        let matches = self
            .document
            .custom_properties
            .iter()
            .flat_map(|properties| &properties.properties)
            .filter(|property| {
                property
                    .name
                    .as_deref()
                    .is_some_and(|candidate| candidate.eq_ignore_ascii_case(name))
            })
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return keep(&format!(
                "DOCPROPERTY target {name} was not found or is ambiguous"
            ));
        }
        match custom_property_text(&matches[0].value) {
            Some(value) => FieldOutcome::Resolved(value),
            None => keep(&format!(
                "DOCPROPERTY target {name} has an unsupported value"
            )),
        }
    }

    fn core_property(&self, name: &str) -> Option<&str> {
        let properties = self.document.core_properties.as_ref()?;
        match normalized_name(name).as_str() {
            "title" => properties.title.as_deref(),
            "subject" => properties.subject.as_deref(),
            "author" | "creator" => properties.creator.as_deref(),
            "comments" | "description" => properties.description.as_deref(),
            "keywords" => properties.keywords.as_deref(),
            "lastsavedby" | "lastmodifiedby" => properties.last_modified_by.as_deref(),
            "createtime" | "created" => properties.created.as_deref(),
            "savedate" | "modified" => properties.modified.as_deref(),
            _ => None,
        }
    }

    fn evaluate_docvariable(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(name) = text_argument(instruction, 0) else {
            return keep("DOCVARIABLE requires a variable name");
        };
        let matches = self
            .document
            .settings
            .iter()
            .flat_map(|settings| settings.document_variables())
            .filter(|variable| variable.name.eq_ignore_ascii_case(name))
            .collect::<Vec<_>>();
        if matches.len() == 1 {
            FieldOutcome::Resolved(matches[0].value.clone())
        } else {
            keep(&format!(
                "DOCVARIABLE target {name} was not found or is ambiguous"
            ))
        }
    }

    fn evaluate_styleref(
        &self,
        instruction: &FieldInstruction,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(target) = text_argument(instruction, 0) else {
            return keep("STYLEREF requires a style name");
        };
        let style_ids = self
            .document
            .styles
            .styles
            .iter()
            .filter(|style| {
                style.style_id.eq_ignore_ascii_case(target)
                    || style
                        .name
                        .as_deref()
                        .is_some_and(|name| name.eq_ignore_ascii_case(target))
            })
            .map(|style| style.style_id.as_str())
            .collect::<Vec<_>>();
        let matches = |paragraph: &&CT_P| {
            paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.style_id.as_deref())
                .is_some_and(|id| style_ids.contains(&id))
        };
        let source = if has_switch(instruction, "l") {
            paragraphs
                .iter()
                .enumerate()
                .rev()
                .find(|(index, paragraph)| *index != paragraph_index && matches(paragraph))
        } else {
            paragraphs[..paragraph_index]
                .iter()
                .enumerate()
                .rev()
                .find(|(_, paragraph)| matches(paragraph))
                .or_else(|| {
                    paragraphs[paragraph_index + 1..]
                        .iter()
                        .enumerate()
                        .find(|(_, paragraph)| matches(paragraph))
                        .map(|(index, paragraph)| (paragraph_index + 1 + index, paragraph))
                })
        };
        let Some((source_index, source)) = source else {
            return keep(&format!("STYLEREF target {target} was not found"));
        };
        let mut effective = style::resolve_paragraph_properties(
            source
                .properties
                .as_ref()
                .and_then(|properties| properties.style_id.as_deref()),
            &self.document.styles,
        );
        if let Some(properties) = source.properties.as_ref() {
            effective.merge_from(properties);
        }
        if ["n", "r", "t", "w"]
            .iter()
            .any(|name| has_switch(instruction, name))
            && effective.num_id.is_some_and(|num_id| num_id != 0)
        {
            return keep("STYLEREF numbered source formatting is unsupported");
        }
        let mut value = source.text();
        if has_switch(instruction, "p") {
            value.push_str(if source_index < paragraph_index {
                " above"
            } else {
                " below"
            });
        }
        FieldOutcome::Resolved(value)
    }

    fn evaluate_includetext(&self, instruction: &FieldInstruction) -> FieldOutcome {
        if has_switch(instruction, "c") {
            return keep("INCLUDETEXT converter selection is unsupported");
        }
        let Some(source) = text_argument(instruction, 0) else {
            return keep("INCLUDETEXT requires a source name");
        };
        let key = text_argument(instruction, 1)
            .map(|bookmark| format!("{source}#{bookmark}"))
            .unwrap_or_else(|| source.to_owned());
        match self.context.included_text.get(&key) {
            Some(value) => FieldOutcome::Resolved(value.clone()),
            None => keep(&format!("INCLUDETEXT input {key} was not supplied")),
        }
    }

    fn evaluate_date_time(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(now) = self.context.now else {
            return keep(&format!(
                "{} requires an explicit date and time",
                instruction.name
            ));
        };
        if !valid_date_time(now) {
            return keep("field date and time is invalid");
        }
        let default_picture = if instruction.name == "DATE" {
            "M/d/yyyy"
        } else {
            "h:mm:ss AM/PM"
        };
        let picture = switch_text(instruction, "@").unwrap_or(default_picture);
        match format_date_time(now, picture) {
            Ok(value) => FieldOutcome::Resolved(value),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_filename(&self, instruction: &FieldInstruction) -> FieldOutcome {
        if has_switch(instruction, "p") {
            return self
                .context
                .file_path
                .clone()
                .map(FieldOutcome::Resolved)
                .unwrap_or_else(|| keep("FILENAME path was not supplied"));
        }
        self.context
            .file_name
            .clone()
            .or_else(|| {
                self.context
                    .file_path
                    .as_deref()
                    .and_then(lexical_file_name)
                    .map(str::to_owned)
            })
            .map(FieldOutcome::Resolved)
            .unwrap_or_else(|| keep("FILENAME name was not supplied"))
    }

    fn evaluate_author(&self) -> FieldOutcome {
        self.document
            .core_properties
            .as_ref()
            .and_then(|properties| properties.creator.clone())
            .map(FieldOutcome::Resolved)
            .unwrap_or_else(|| keep("AUTHOR metadata was not found"))
    }

    fn evaluate_mergefield(&self, instruction: &FieldInstruction) -> FieldOutcome {
        let Some(name) = text_argument(instruction, 0) else {
            return keep("MERGEFIELD requires a field name");
        };
        let Some(value) = self.context.merge_fields.get(name) else {
            if self.missing_merge_fields_as_empty {
                return FieldOutcome::Resolved(String::new());
            }
            return keep(&format!("MERGEFIELD input {name} was not supplied"));
        };
        if value.is_empty() {
            return FieldOutcome::Resolved(String::new());
        }
        let mut result = String::new();
        if let Some(prefix) = switch_text(instruction, "b") {
            result.push_str(prefix);
        }
        result.push_str(value);
        if let Some(suffix) = switch_text(instruction, "f") {
            result.push_str(suffix);
        }
        FieldOutcome::Resolved(result)
    }

    fn evaluate_formula(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        if instruction.arguments.is_empty() {
            return keep("formula requires an expression");
        }
        let mut expression = String::new();
        for argument in &instruction.arguments {
            let value = match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                Ok(value) => value,
                Err(diagnostic) => return keep(&diagnostic),
            };
            if !expression.is_empty() {
                expression.push(' ');
            }
            expression.push_str(&value);
            if expression.len() > MAX_FORMULA_BYTES {
                return keep("formula exceeds the 4096-byte limit");
            }
        }
        match FormulaParser::new(&expression).and_then(FormulaParser::parse) {
            Ok(value) => FieldOutcome::Resolved(format_formula_value(value)),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_mail_merge_control(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        match instruction.name.as_str() {
            "MERGESEQ" if self.context.merge_sequence_number == Some(0) => {
                return keep("MERGESEQ merge sequence number must be one-based");
            }
            "NEXT" | "NEXTIF" | "SKIPIF" | "MERGEREC"
                if self.context.merge_record_number == Some(0) =>
            {
                return keep(&format!(
                    "{} merge record number must be one-based",
                    instruction.name
                ));
            }
            _ => {}
        }
        let condition = match instruction.name.as_str() {
            "NEXTIF" | "SKIPIF" => {
                let values = instruction
                    .arguments
                    .iter()
                    .map(|argument| {
                        self.resolve_argument(argument, story, paragraphs, paragraph_index)
                    })
                    .collect::<std::result::Result<Vec<_>, _>>();
                let values = match values {
                    Ok(values) => values,
                    Err(diagnostic) => return keep(&diagnostic),
                };
                let Some(condition) = compare_if(&values[0], &values[1], &values[2]) else {
                    return keep(&format!(
                        "{} operator {} is unsupported",
                        instruction.name, values[1]
                    ));
                };
                condition
            }
            _ => false,
        };
        let state =
            self.mail_merge_stories
                .entry(story.to_owned())
                .or_insert(MailMergeStoryState {
                    record_number: self.context.merge_record_number,
                    sequence_number: self.context.merge_sequence_number,
                });
        match instruction.name.as_str() {
            "NEXT" => {
                let Some(current_record) = state.record_number else {
                    return keep("NEXT requires an explicit merge record number");
                };
                let Some(record_number) = current_record.checked_add(1) else {
                    return keep("NEXT record number overflowed");
                };
                state.record_number = Some(record_number);
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecord { record_number })
            }
            "NEXTIF" => {
                let Some(mut record_number) = state.record_number else {
                    return keep("NEXTIF requires an explicit merge record number");
                };
                if condition {
                    let Some(next_record_number) = record_number.checked_add(1) else {
                        return keep("NEXTIF record number overflowed");
                    };
                    record_number = next_record_number;
                    state.record_number = Some(record_number);
                }
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecordIf {
                    condition,
                    record_number,
                })
            }
            "SKIPIF" => match state.record_number {
                Some(record_number) => {
                    FieldOutcome::MailMergeControl(MailMergeControl::SkipRecordIf {
                        condition,
                        record_number,
                    })
                }
                None => keep("SKIPIF requires an explicit merge record number"),
            },
            "MERGEREC" => state
                .record_number
                .map(MailMergeControl::RecordNumber)
                .map(FieldOutcome::MailMergeControl)
                .unwrap_or_else(|| keep("MERGEREC requires an explicit merge record number")),
            "MERGESEQ" => state
                .sequence_number
                .map(MailMergeControl::SequenceNumber)
                .map(FieldOutcome::MailMergeControl)
                .unwrap_or_else(|| keep("MERGESEQ requires an explicit merge sequence number")),
            _ => unreachable!(),
        }
    }

    fn evaluate_barcode(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(source) = instruction.arguments.first() else {
            return keep(&format!("{} requires a value", instruction.name));
        };
        let value = match source {
            FieldArgument::Text(source) if instruction.name == "MERGEBARCODE" => {
                match self.context.merge_fields.get(source) {
                    Some(value) => value.clone(),
                    None => return keep(&format!("MERGEBARCODE input {source} was not supplied")),
                }
            }
            _ => match self.resolve_argument(source, story, paragraphs, paragraph_index) {
                Ok(value) => value,
                Err(diagnostic) => return keep(&diagnostic),
            },
        };
        let Some(kind) = instruction.arguments.get(1) else {
            return keep(&format!("{} requires a barcode type", instruction.name));
        };
        let kind = match self.resolve_argument(kind, story, paragraphs, paragraph_index) {
            Ok(value) => value,
            Err(diagnostic) => return keep(&diagnostic),
        };
        let mut switches = Vec::with_capacity(instruction.switches.len());
        for field_switch in &instruction.switches {
            let argument = match &field_switch.argument {
                Some(argument) => {
                    match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                        Ok(value) => Some(value),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                }
                None => None,
            };
            switches.push((field_switch.name.clone(), argument));
        }
        match parse_barcode(&instruction.name, &switches, value, &kind) {
            Ok(barcode) => FieldOutcome::Barcode(barcode),
            Err(diagnostic) => keep(&diagnostic),
        }
    }

    fn evaluate_toc(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let mut toc = TocField {
            heading_levels: None,
            custom_styles: Vec::new(),
            entries: TocEntrySelection::None,
            sequence_identifier: None,
            bookmark: None,
            hyperlink: false,
            use_outline_levels: false,
            omit_page_number_levels: None,
            page_number_separator: None,
            entry_page_separator: None,
        };
        let mut has_explicit_source = false;
        for field_switch in &instruction.switches {
            let argument = match &field_switch.argument {
                Some(argument) => {
                    match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                        Ok(value) => Some(value),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                }
                None => None,
            };
            match field_switch.name.as_str() {
                "h" if argument.is_none() => toc.hyperlink = true,
                "u" if argument.is_none() => {
                    toc.use_outline_levels = true;
                    has_explicit_source = true;
                }
                "o" => {
                    has_explicit_source = true;
                    toc.heading_levels = match argument {
                        Some(value) => match parse_level_range(&value, "TOC heading") {
                            Ok(levels) => Some(levels),
                            Err(diagnostic) => return keep(&diagnostic),
                        },
                        None => Some((1, 9)),
                    };
                }
                "n" => {
                    toc.omit_page_number_levels = match argument {
                        Some(value) => match parse_level_range(&value, "TOC omitted page-number") {
                            Ok(levels) => Some(levels),
                            Err(diagnostic) => return keep(&diagnostic),
                        },
                        None => Some((1, 9)),
                    };
                }
                "t" => {
                    let Some(value) = argument else {
                        return keep("field TOC switch \\t requires a text argument");
                    };
                    has_explicit_source = true;
                    toc.custom_styles = match parse_custom_styles(&value) {
                        Ok(styles) => styles,
                        Err(diagnostic) => return keep(&diagnostic),
                    };
                }
                "f" => {
                    has_explicit_source = true;
                    toc.entries = argument.map_or(TocEntrySelection::All, |value| {
                        TocEntrySelection::Identifier(value)
                    });
                }
                "b" | "p" | "s" | "d" => {
                    let Some(value) = argument else {
                        return keep(&format!(
                            "field TOC switch \\{} requires a text argument",
                            field_switch.name
                        ));
                    };
                    match field_switch.name.as_str() {
                        "b" => toc.bookmark = Some(value),
                        "p" if value.chars().count() == 1 => {
                            toc.page_number_separator = Some(value)
                        }
                        "p" => {
                            return keep(
                                "TOC page-number separator must contain exactly one character",
                            );
                        }
                        "s" if !value.is_empty() => toc.sequence_identifier = Some(value),
                        "s" => return keep("TOC sequence identifier must not be empty"),
                        "d" => toc.entry_page_separator = Some(value),
                        _ => unreachable!(),
                    }
                }
                name if argument.is_some() => {
                    return keep(&format!(
                        "field TOC switch \\{name} does not take an argument"
                    ));
                }
                name => return keep(&format!("field TOC uses unsupported switch \\{name}")),
            }
        }
        if !has_explicit_source {
            toc.heading_levels = Some((1, 9));
        }
        FieldOutcome::TableOfContents(toc)
    }

    fn evaluate_tc(
        &mut self,
        instruction: &FieldInstruction,
        story: &str,
        paragraphs: &[&CT_P],
        paragraph_index: usize,
    ) -> FieldOutcome {
        let Some(entry) = instruction.arguments.first() else {
            return keep("TC requires entry text");
        };
        let entry = match self.resolve_argument(entry, story, paragraphs, paragraph_index) {
            Ok(entry) => entry,
            Err(diagnostic) => return keep(&diagnostic),
        };
        if entry.is_empty() {
            return keep("TC entry text must not be empty");
        }
        let mut tc = TcField {
            entry,
            level: 1,
            table_identifier: None,
            omit_page_number: false,
        };
        for field_switch in &instruction.switches {
            let argument = match &field_switch.argument {
                Some(argument) => {
                    match self.resolve_argument(argument, story, paragraphs, paragraph_index) {
                        Ok(value) => Some(value),
                        Err(diagnostic) => return keep(&diagnostic),
                    }
                }
                None => None,
            };
            match field_switch.name.as_str() {
                "n" if argument.is_none() => tc.omit_page_number = true,
                "f" => match argument {
                    Some(value) if !value.is_empty() => tc.table_identifier = Some(value),
                    Some(_) => return keep("TC table identifier must not be empty"),
                    None => return keep("field TC switch \\f requires a text argument"),
                },
                "l" => {
                    let Some(value) = argument else {
                        return keep("field TC switch \\l requires a text argument");
                    };
                    tc.level = match parse_toc_level(&value, "TC") {
                        Ok(level) => level,
                        Err(diagnostic) => return keep(&diagnostic),
                    };
                }
                name if argument.is_some() => {
                    return keep(&format!(
                        "field TC switch \\{name} does not take an argument"
                    ));
                }
                name => return keep(&format!("field TC uses unsupported switch \\{name}")),
            }
        }
        FieldOutcome::TableOfContentsEntry(tc)
    }
}

fn referenced_header_footer_parts(document: &Document, is_header: bool) -> Vec<(String, Vec<u8>)> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Vec::new();
    };
    let relationship_type = if is_header {
        rel_types::HEADER
    } else {
        rel_types::FOOTER
    };
    let sections = document
        .document
        .body
        .content
        .iter()
        .filter_map(|content| match content {
            BodyContent::Paragraph(paragraph) => paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.sect_pr.as_ref()),
            BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => None,
        })
        .chain(document.document.body.sect_pr.iter());
    let mut seen = HashSet::new();
    let mut parts = Vec::new();
    for section in sections {
        let references = section_header_footer_references(section, is_header);
        for reference in references {
            let Some(relationship) = relationships.get_by_id(&reference.rel_id) else {
                continue;
            };
            if relationship.rel_type != relationship_type
                || relationship.target_mode.as_deref() == Some("External")
            {
                continue;
            }
            let part_name =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            if seen.insert(part_name.clone())
                && let Some(xml) = document.package.get_part(&part_name)
            {
                parts.push((part_name, xml.to_vec()));
            }
        }
    }
    parts
}

fn section_header_footer_references(
    section: &CT_SectPr,
    is_header: bool,
) -> &[rdocx_oxml::header_footer::HdrFtrRef] {
    if is_header {
        &section.header_refs
    } else {
        &section.footer_refs
    }
}

fn relationship_parts(document: &Document, relationship_type: &str) -> Vec<(String, Vec<u8>)> {
    let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) else {
        return Vec::new();
    };
    let mut seen = HashSet::new();
    relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == relationship_type)
        .filter(|relationship| relationship.target_mode.as_deref() != Some("External"))
        .filter_map(|relationship| {
            let part_name =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            if !seen.insert(part_name.clone()) {
                return None;
            }
            Some((
                part_name.clone(),
                document.package.get_part(&part_name)?.to_vec(),
            ))
        })
        .collect()
}

fn normal_note_paragraphs(notes: &CT_Footnotes) -> Vec<&CT_P> {
    notes
        .footnotes
        .iter()
        .filter(|note| note.note_type == NoteType::Normal)
        .flat_map(|note| note.paragraphs.iter())
        .collect()
}

#[derive(Clone, Copy)]
enum PackageStoryKind {
    HeaderFooter,
    Footnotes,
    Endnotes,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum StoryElementKind {
    HeaderFooterRoot,
    FootnotesRoot,
    EndnotesRoot,
    NormalNote,
    Paragraph,
    Other,
}

struct StoryElement {
    kind: StoryElementKind,
    start: usize,
}

struct FieldSourceEdit {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

fn patch_story_field_sources(
    xml: &[u8],
    paragraphs: &[&CT_P],
    story_kind: PackageStoryKind,
) -> Result<Vec<u8>> {
    let paragraph_spans = story_paragraph_spans(xml, story_kind)?;
    if paragraph_spans.len() != paragraphs.len() {
        return Err(Error::Other(format!(
            "package story paragraph scan found {} of {} typed paragraphs",
            paragraph_spans.len(),
            paragraphs.len()
        )));
    }
    let mut edits = Vec::new();
    for (paragraph, (paragraph_start, paragraph_end)) in paragraphs.iter().zip(paragraph_spans) {
        let mut search_start = 0usize;
        for run in paragraph.runs() {
            for content in &run.content {
                let RunContent::Field(field) = content else {
                    continue;
                };
                let Some((source, replacement)) = field.source_replacement()? else {
                    return Err(Error::Other(
                        "parsed package story field has no source XML".to_owned(),
                    ));
                };
                let Some(start) = find_typed_field_source(
                    xml,
                    paragraph_start,
                    paragraph_end,
                    source,
                    search_start,
                )?
                else {
                    return Err(Error::Other(
                        "package story field source was not found at its typed paragraph boundary"
                            .to_owned(),
                    ));
                };
                let end = start + source.len();
                edits.push(FieldSourceEdit {
                    start: paragraph_start + start,
                    end: paragraph_start + end,
                    replacement,
                });
                search_start = end;
            }
        }
    }

    let mut updated = xml.to_vec();
    for edit in edits.into_iter().rev() {
        updated.splice(edit.start..edit.end, edit.replacement);
    }
    Ok(updated)
}

fn story_paragraph_spans(xml: &[u8], story_kind: PackageStoryKind) -> Result<Vec<(usize, usize)>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<StoryElement>::new();
    let mut paragraphs = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package story XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        match event {
            Event::Start(element) => {
                let parent = stack.last().map(|element| element.kind);
                let kind = match story_kind {
                    PackageStoryKind::HeaderFooter
                        if stack.is_empty()
                            && word
                            && (matches_local_name(element.name().as_ref(), b"hdr")
                                || matches_local_name(element.name().as_ref(), b"ftr")) =>
                    {
                        StoryElementKind::HeaderFooterRoot
                    }
                    PackageStoryKind::Endnotes
                        if stack.is_empty()
                            && word
                            && matches_local_name(element.name().as_ref(), b"endnotes") =>
                    {
                        StoryElementKind::EndnotesRoot
                    }
                    PackageStoryKind::Footnotes
                        if stack.is_empty()
                            && word
                            && matches_local_name(element.name().as_ref(), b"footnotes") =>
                    {
                        StoryElementKind::FootnotesRoot
                    }
                    PackageStoryKind::Footnotes
                        if parent == Some(StoryElementKind::FootnotesRoot)
                            && word
                            && matches_local_name(element.name().as_ref(), b"footnote")
                            && note_is_normal(&element) =>
                    {
                        StoryElementKind::NormalNote
                    }
                    PackageStoryKind::Endnotes
                        if parent == Some(StoryElementKind::EndnotesRoot)
                            && word
                            && matches_local_name(element.name().as_ref(), b"endnote")
                            && note_is_normal(&element) =>
                    {
                        StoryElementKind::NormalNote
                    }
                    PackageStoryKind::HeaderFooter
                        if parent == Some(StoryElementKind::HeaderFooterRoot)
                            && word
                            && matches_local_name(element.name().as_ref(), b"p") =>
                    {
                        StoryElementKind::Paragraph
                    }
                    PackageStoryKind::Endnotes
                        if parent == Some(StoryElementKind::NormalNote)
                            && word
                            && matches_local_name(element.name().as_ref(), b"p") =>
                    {
                        StoryElementKind::Paragraph
                    }
                    PackageStoryKind::Footnotes
                        if parent == Some(StoryElementKind::NormalNote)
                            && word
                            && matches_local_name(element.name().as_ref(), b"p") =>
                    {
                        StoryElementKind::Paragraph
                    }
                    _ => StoryElementKind::Other,
                };
                stack.push(StoryElement {
                    kind,
                    start: before,
                });
            }
            Event::End(_) => {
                let Some(element) = stack.pop() else {
                    return Err(Error::Other(
                        "package story XML has an unmatched end element".to_owned(),
                    ));
                };
                if element.kind == StoryElementKind::Paragraph {
                    paragraphs.push((element.start, reader.buffer_position() as usize));
                }
            }
            Event::Eof => {
                if !stack.is_empty() {
                    return Err(Error::Other(
                        "package story XML has an unclosed element".to_owned(),
                    ));
                }
                return Ok(paragraphs);
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn note_is_normal(element: &BytesStart<'_>) -> bool {
    let mut id = 0i32;
    let mut note_type = None;
    for attribute in element.attributes().flatten() {
        if matches_local_name(attribute.key.as_ref(), b"id") {
            id = std::str::from_utf8(&attribute.value)
                .unwrap_or("0")
                .parse()
                .unwrap_or(0);
        } else if matches_local_name(attribute.key.as_ref(), b"type") {
            note_type = Some(String::from_utf8_lossy(&attribute.value).into_owned());
        }
    }
    !matches!(
        note_type.as_deref(),
        Some("separator" | "continuationSeparator" | "continuationNotice")
    ) && id > 0
}

fn find_typed_field_source(
    xml: &[u8],
    paragraph_start: usize,
    paragraph_end: usize,
    source: &[u8],
    search_start: usize,
) -> Result<Option<usize>> {
    let paragraph_xml = &xml[paragraph_start..paragraph_end];
    let mut candidate_start = search_start;
    while let Some(relative) = find_bytes(&paragraph_xml[candidate_start..], source) {
        let candidate = candidate_start + relative;
        if field_source_has_typed_ancestors(xml, paragraph_start + candidate, paragraph_start)? {
            return Ok(Some(candidate));
        }
        candidate_start = candidate + source.len().max(1);
    }
    Ok(None)
}

fn field_source_has_typed_ancestors(
    xml: &[u8],
    source_start: usize,
    paragraph_start: usize,
) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut ancestors = Vec::<(usize, bool, Vec<u8>)>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid package paragraph XML: {error}")))?;
        let word = namespace_is_word(&namespace);
        match event {
            Event::Start(element) => {
                if before == source_start {
                    let name = element.name();
                    let local = local_name(name.as_ref());
                    return Ok(word
                        && matches!(local, b"fldSimple" | b"r")
                        && typed_field_ancestors(&ancestors, paragraph_start));
                }
                ancestors.push((before, word, local_name(element.name().as_ref()).to_vec()));
            }
            Event::Empty(element) => {
                if before == source_start {
                    let name = element.name();
                    let local = local_name(name.as_ref());
                    return Ok(word
                        && matches!(local, b"fldSimple" | b"r")
                        && typed_field_ancestors(&ancestors, paragraph_start));
                }
            }
            Event::End(_) => {
                ancestors.pop();
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn typed_field_ancestors(ancestors: &[(usize, bool, Vec<u8>)], paragraph_start: usize) -> bool {
    let Some(paragraph_index) = ancestors
        .iter()
        .position(|(start, _, _)| *start == paragraph_start)
    else {
        return false;
    };
    let (_, paragraph_word, paragraph) = &ancestors[paragraph_index];
    *paragraph_word
        && paragraph.as_slice() == b"p"
        && ancestors
            .iter()
            .skip(paragraph_index + 1)
            .all(|(_, word, local)| {
                *word
                    && matches!(
                        local.as_slice(),
                        b"hyperlink"
                            | b"sdt"
                            | b"sdtContent"
                            | b"ins"
                            | b"del"
                            | b"moveFrom"
                            | b"moveTo"
                    )
            })
}

fn namespace_is_word(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == W_NS.as_bytes())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        None
    } else {
        haystack
            .windows(needle.len())
            .position(|window| window == needle)
    }
}

fn apply_updates_to_notes(
    notes: &mut CT_Footnotes,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for note in &mut notes.footnotes {
        if note.note_type == NoteType::Normal {
            apply_updates_to_paragraphs(&mut note.paragraphs, updates, update_index);
        }
    }
}

fn apply_updates_to_paragraphs(
    paragraphs: &mut [CT_P],
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for paragraph in paragraphs {
        apply_updates_to_paragraph(paragraph, updates, update_index);
    }
}

fn apply_updates_to_body(
    body: &mut CT_Body,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for content in &mut body.content {
        match content {
            BodyContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index);
            }
            BodyContent::Table(table) => apply_updates_to_table(table, updates, update_index),
            BodyContent::ContentControl(control) => {
                apply_updates_to_block_control(control, updates, update_index);
            }
            BodyContent::RawXml(_) => {}
        }
    }
}

fn apply_updates_to_table(
    table: &mut CT_Tbl,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter_mut()
            .filter(|(position, _, _)| *position == boundary)
        {
            apply_updates_to_block_control(control, updates, update_index);
        }
        if let Some(row) = table.rows.get_mut(boundary) {
            apply_updates_to_row(row, updates, update_index);
        }
    }
}

fn apply_updates_to_row(row: &mut CT_Row, updates: &[CachedFieldUpdate], update_index: &mut usize) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter_mut()
            .filter(|(position, _, _)| *position == boundary)
        {
            apply_updates_to_block_control(control, updates, update_index);
        }
        if let Some(cell) = row.cells.get_mut(boundary) {
            apply_updates_to_cell(cell, updates, update_index);
        }
    }
}

fn apply_updates_to_cell(
    cell: &mut CT_Tc,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index);
            }
            CellContent::Table(table) => apply_updates_to_table(table, updates, update_index),
            CellContent::ContentControl(control) => {
                apply_updates_to_block_control(control, updates, update_index);
            }
        }
    }
}

fn apply_updates_to_block_control(
    control: &mut CT_Sdt,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index);
            }
            SdtContent::Table(table) => apply_updates_to_table(table, updates, update_index),
            SdtContent::Row(row) => apply_updates_to_row(row, updates, update_index),
            SdtContent::Cell(cell) => apply_updates_to_cell(cell, updates, update_index),
            SdtContent::ContentControl(control) => {
                apply_updates_to_block_control(control, updates, update_index);
            }
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
}

fn apply_updates_to_run_control(
    control: &mut CT_Sdt,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for content in &mut control.content {
        match content {
            SdtContent::Run(run) => apply_updates_to_run(run, updates, update_index),
            SdtContent::Paragraph(paragraph) => {
                apply_updates_to_paragraph(paragraph, updates, update_index);
            }
            SdtContent::ContentControl(control) => {
                apply_updates_to_run_control(control, updates, update_index);
            }
            SdtContent::Table(_)
            | SdtContent::Row(_)
            | SdtContent::Cell(_)
            | SdtContent::RawXml(_) => {}
        }
    }
}

fn apply_updates_to_paragraph(
    paragraph: &mut CT_P,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for boundary in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter_mut()
            .filter(|(position, _, _, _)| *position == boundary)
        {
            apply_updates_to_run_control(control, updates, update_index);
        }
        if let Some(run) = paragraph.runs.get_mut(boundary) {
            apply_updates_to_run(run, updates, update_index);
        }
    }
}

fn apply_updates_to_run(
    run: &mut rdocx_oxml::text::CT_R,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    for content in &mut run.content {
        if let RunContent::Field(field) = content {
            apply_updates_to_field(field, updates, update_index);
        }
    }
}

fn apply_updates_to_field(
    field: &mut Field,
    updates: &[CachedFieldUpdate],
    update_index: &mut usize,
) {
    let Some(update) = updates.get(*update_index) else {
        return;
    };
    field.cached_result.clone_from(&update.cached_result);
    field.dirty = Some(update.dirty);
    *update_index += 1;

    let nested_pointers = field
        .nested_fields_in_source_order()
        .into_iter()
        .map(|nested| std::ptr::from_ref(nested) as usize)
        .collect::<Vec<_>>();
    for pointer in nested_pointers {
        if let Some(nested) = nested_field_mut(field, pointer) {
            apply_updates_to_field(nested, updates, update_index);
        }
    }
}

fn nested_field_mut(field: &mut Field, pointer: usize) -> Option<&mut Field> {
    for argument in &mut field.instruction.arguments {
        if let FieldArgument::Nested(nested) = argument
            && std::ptr::from_ref(nested.as_ref()) as usize == pointer
        {
            return Some(nested.as_mut());
        }
    }
    for field_switch in &mut field.instruction.switches {
        if let Some(FieldArgument::Nested(nested)) = &mut field_switch.argument
            && std::ptr::from_ref(nested.as_ref()) as usize == pointer
        {
            return Some(nested.as_mut());
        }
    }
    None
}

fn collect_body_paragraphs<'a>(body: &'a CT_Body, output: &mut Vec<&'a CT_P>) {
    for content in &body.content {
        match content {
            BodyContent::Paragraph(paragraph) => output.push(paragraph),
            BodyContent::Table(table) => collect_table_paragraphs(table, output),
            BodyContent::ContentControl(control) => {
                collect_control_paragraphs(control, BlockControlOwner::Body, output)
            }
            BodyContent::RawXml(_) => {}
        }
    }
}

fn collect_table_paragraphs<'a>(table: &'a CT_Tbl, output: &mut Vec<&'a CT_P>) {
    for boundary in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter()
            .filter(|(position, _, _)| *position == boundary)
        {
            collect_control_paragraphs(control, BlockControlOwner::Table, output);
        }
        if let Some(row) = table.rows.get(boundary) {
            collect_row_paragraphs(row, output);
        }
    }
}

fn collect_row_paragraphs<'a>(row: &'a CT_Row, output: &mut Vec<&'a CT_P>) {
    for boundary in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter()
            .filter(|(position, _, _)| *position == boundary)
        {
            collect_control_paragraphs(control, BlockControlOwner::Row, output);
        }
        if let Some(cell) = row.cells.get(boundary) {
            collect_cell_paragraphs(cell, output);
        }
    }
}

fn collect_cell_paragraphs<'a>(cell: &'a CT_Tc, output: &mut Vec<&'a CT_P>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => output.push(paragraph),
            CellContent::Table(table) => collect_table_paragraphs(table, output),
            CellContent::ContentControl(control) => {
                collect_control_paragraphs(control, BlockControlOwner::Cell, output)
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BlockControlOwner {
    Body,
    Table,
    Row,
    Cell,
}

fn collect_control_paragraphs<'a>(
    control: &'a CT_Sdt,
    owner: BlockControlOwner,
    output: &mut Vec<&'a CT_P>,
) {
    for content in &control.content {
        match (owner, content) {
            (
                BlockControlOwner::Body | BlockControlOwner::Cell,
                SdtContent::Paragraph(paragraph),
            ) => output.push(paragraph),
            (BlockControlOwner::Body | BlockControlOwner::Cell, SdtContent::Table(table)) => {
                collect_table_paragraphs(table, output)
            }
            (BlockControlOwner::Table, SdtContent::Row(row)) => collect_row_paragraphs(row, output),
            (BlockControlOwner::Row, SdtContent::Cell(cell)) => {
                collect_cell_paragraphs(cell, output)
            }
            (_, SdtContent::ContentControl(control)) => {
                collect_control_paragraphs(control, owner, output)
            }
            _ => {}
        }
    }
}

fn keep(diagnostic: &str) -> FieldOutcome {
    FieldOutcome::KeepStored {
        diagnostic: format!("{diagnostic}, stored display retained"),
    }
}

const MAX_FORMULA_BYTES: usize = 4096;
const MAX_FORMULA_TOKENS: usize = 512;
const MAX_FORMULA_DEPTH: usize = 32;

fn parse_toc_level(value: &str, field: &str) -> std::result::Result<u8, String> {
    value
        .trim()
        .parse::<u8>()
        .ok()
        .filter(|level| (1..=9).contains(level))
        .ok_or_else(|| format!("{field} level must be from 1 through 9"))
}

fn parse_level_range(value: &str, field: &str) -> std::result::Result<(u8, u8), String> {
    let Some((start, end)) = value.split_once('-') else {
        let level = parse_toc_level(value, field)?;
        return Ok((level, level));
    };
    let start = parse_toc_level(start, field)?;
    let end = parse_toc_level(end, field)?;
    if start > end {
        return Err(format!("{field} range starts after it ends"));
    }
    Ok((start, end))
}

fn parse_custom_styles(value: &str) -> std::result::Result<Vec<(String, u8)>, String> {
    let parts = value.split(',').collect::<Vec<_>>();
    if parts.len() % 2 != 0 || parts.is_empty() {
        return Err("TOC custom styles require style and level pairs".to_owned());
    }
    let mut styles = Vec::with_capacity(parts.len() / 2);
    for pair in parts.chunks_exact(2) {
        let name = pair[0].trim();
        if name.is_empty() {
            return Err("TOC custom style name must not be empty".to_owned());
        }
        styles.push((name.to_owned(), parse_toc_level(pair[1], "TOC style")?));
    }
    Ok(styles)
}

fn parse_barcode(
    field: &str,
    switches: &[(String, Option<String>)],
    value: String,
    kind: &str,
) -> std::result::Result<BarcodeField, String> {
    let kind = parse_barcode_kind(kind).ok_or_else(|| {
        format!(
            "{field} barcode type {} is unsupported",
            kind.to_ascii_uppercase()
        )
    })?;
    if value.is_empty() || value.chars().count() > 1024 {
        return Err(format!(
            "{field} value must contain from 1 through 1024 characters"
        ));
    }
    validate_barcode_value(field, kind, &value)?;
    let mut barcode = BarcodeField {
        value,
        kind,
        height: None,
        scale: None,
        error_correction: None,
        point_of_sale_style: None,
        case_style: None,
        fix_check_digit: false,
        rotation: None,
        foreground_color: None,
        background_color: None,
        display_text: false,
        add_start_stop: false,
    };
    for (name, argument) in switches {
        match name.as_str() {
            "t" if argument.is_none() => barcode.display_text = true,
            "x" if argument.is_none() => barcode.fix_check_digit = true,
            "d" => {
                if argument.is_some() {
                    return Err(format!(
                        "field {field} switch \\d does not take an argument"
                    ));
                }
                if !matches!(kind, BarcodeKind::Nw7 | BarcodeKind::Code39) {
                    return Err(format!("{field} switch \\d requires NW7 or CODE39"));
                }
                barcode.add_start_stop = true;
            }
            "h" | "s" | "q" | "p" | "c" | "r" | "f" | "b" => {
                let Some(value) = argument.as_deref() else {
                    return Err(format!(
                        "field {field} switch \\{name} requires a text argument"
                    ));
                };
                match name.as_str() {
                    "h" => barcode.height = Some(parse_unsigned_integer(value, field, "height")?),
                    "s" => {
                        barcode.scale = Some(
                            u16::try_from(parse_bounded_integer(value, 10, 1000, field, "scale")?)
                                .expect("bounded barcode scale fits u16"),
                        )
                    }
                    "q" => {
                        if kind != BarcodeKind::Qr {
                            return Err(format!("{field} switch \\q requires QR"));
                        }
                        barcode.error_correction = Some(
                            u8::try_from(parse_bounded_integer(
                                value,
                                0,
                                3,
                                field,
                                "error correction",
                            )?)
                            .expect("bounded error correction fits u8"),
                        );
                    }
                    "p" => {
                        if !matches!(
                            kind,
                            BarcodeKind::Upca
                                | BarcodeKind::Upce
                                | BarcodeKind::Ean13
                                | BarcodeKind::Ean8
                        ) {
                            return Err(format!(
                                "{field} switch \\p requires UPCA, UPCE, EAN13, or EAN8"
                            ));
                        }
                        barcode.point_of_sale_style =
                            Some(parse_point_of_sale_style(value).ok_or_else(|| {
                                format!(
                                    "{field} point-of-sale style must be STD, SUP2, SUP5, or CASE"
                                )
                            })?);
                    }
                    "c" => {
                        if !matches!(kind, BarcodeKind::Case | BarcodeKind::Itf14) {
                            return Err(format!("{field} switch \\c requires ITF14"));
                        }
                        barcode.case_style = Some(parse_case_style(value).ok_or_else(|| {
                            format!("{field} case style must be STD, EXT, or ADD")
                        })?);
                    }
                    "r" => {
                        let rotation = parse_bounded_integer(value, 0, 3, field, "rotation")?;
                        barcode.rotation =
                            Some(u8::try_from(rotation).expect("bounded barcode rotation fits u8"));
                    }
                    "f" => barcode.foreground_color = Some(parse_barcode_color(value, field)?),
                    "b" => barcode.background_color = Some(parse_barcode_color(value, field)?),
                    _ => unreachable!(),
                }
            }
            name if argument.is_some() => {
                return Err(format!(
                    "field {field} switch \\{name} does not take an argument"
                ));
            }
            _ => return Err(format!("field {field} uses unsupported switch \\{name}")),
        }
    }
    Ok(barcode)
}

fn parse_barcode_kind(value: &str) -> Option<BarcodeKind> {
    match value.to_ascii_uppercase().as_str() {
        "UPCA" => Some(BarcodeKind::Upca),
        "UPCE" => Some(BarcodeKind::Upce),
        "JAN13" => Some(BarcodeKind::Jan13),
        "JAN8" => Some(BarcodeKind::Jan8),
        "EAN13" => Some(BarcodeKind::Ean13),
        "EAN8" => Some(BarcodeKind::Ean8),
        "CASE" => Some(BarcodeKind::Case),
        "ITF14" => Some(BarcodeKind::Itf14),
        "NW7" => Some(BarcodeKind::Nw7),
        "CODE39" => Some(BarcodeKind::Code39),
        "CODE128" => Some(BarcodeKind::Code128),
        "JPPOST" => Some(BarcodeKind::JpPost),
        "QR" => Some(BarcodeKind::Qr),
        _ => None,
    }
}

fn barcode_kind_name(kind: BarcodeKind) -> &'static str {
    match kind {
        BarcodeKind::Upca => "UPCA",
        BarcodeKind::Upce => "UPCE",
        BarcodeKind::Jan13 => "JAN13",
        BarcodeKind::Jan8 => "JAN8",
        BarcodeKind::Ean13 => "EAN13",
        BarcodeKind::Ean8 => "EAN8",
        BarcodeKind::Case => "CASE",
        BarcodeKind::Itf14 => "ITF14",
        BarcodeKind::Nw7 => "NW7",
        BarcodeKind::Code39 => "CODE39",
        BarcodeKind::Code128 => "CODE128",
        BarcodeKind::JpPost => "JPPOST",
        BarcodeKind::Qr => "QR",
    }
}

fn parse_point_of_sale_style(value: &str) -> Option<BarcodePointOfSaleStyle> {
    match value.to_ascii_uppercase().as_str() {
        "STD" => Some(BarcodePointOfSaleStyle::Standard),
        "SUP2" => Some(BarcodePointOfSaleStyle::SupplementalTwoDigit),
        "SUP5" => Some(BarcodePointOfSaleStyle::SupplementalFiveDigit),
        "CASE" => Some(BarcodePointOfSaleStyle::Case),
        _ => None,
    }
}

fn parse_case_style(value: &str) -> Option<BarcodeCaseStyle> {
    match value.to_ascii_uppercase().as_str() {
        "STD" => Some(BarcodeCaseStyle::Standard),
        "EXT" => Some(BarcodeCaseStyle::Extended),
        "ADD" => Some(BarcodeCaseStyle::Add),
        _ => None,
    }
}

fn parse_unsigned_integer(
    value: &str,
    field: &str,
    label: &str,
) -> std::result::Result<u32, String> {
    value
        .parse::<u32>()
        .map_err(|_| format!("{field} {label} must be a nonnegative integer"))
}

fn parse_bounded_integer(
    value: &str,
    minimum: u32,
    maximum: u32,
    field: &str,
    label: &str,
) -> std::result::Result<u32, String> {
    value
        .parse::<u32>()
        .ok()
        .filter(|value| (minimum..=maximum).contains(value))
        .ok_or_else(|| format!("{field} {label} must be from {minimum} through {maximum}"))
}

fn validate_barcode_value(
    field: &str,
    kind: BarcodeKind,
    value: &str,
) -> std::result::Result<(), String> {
    let digit_range = match kind {
        BarcodeKind::Ean8 | BarcodeKind::Jan8 => Some(7..=8),
        BarcodeKind::Ean13 | BarcodeKind::Jan13 => Some(12..=13),
        BarcodeKind::Upca => Some(11..=12),
        BarcodeKind::Upce => Some(6..=8),
        BarcodeKind::Case | BarcodeKind::Itf14 => Some(13..=14),
        _ => None,
    };
    if let Some(range) = digit_range
        && (!value.bytes().all(|byte| byte.is_ascii_digit()) || !range.contains(&value.len()))
    {
        return Err(format!(
            "{field} {} value has an invalid digit count or character",
            barcode_kind_name(kind)
        ));
    }
    if kind == BarcodeKind::Code39
        && !value.bytes().all(|byte| {
            byte.is_ascii_uppercase()
                || byte.is_ascii_digit()
                || matches!(byte, b' ' | b'-' | b'.' | b'$' | b'/' | b'+' | b'%')
        })
    {
        return Err(format!(
            "{field} CODE39 value contains an unsupported character"
        ));
    }
    Ok(())
}

fn parse_barcode_color(value: &str, field: &str) -> std::result::Result<u32, String> {
    let parsed = if let Some(hexadecimal) = value
        .strip_prefix("0x")
        .or_else(|| value.strip_prefix("0X"))
    {
        u32::from_str_radix(hexadecimal, 16).ok()
    } else {
        value.parse::<u32>().ok()
    };
    parsed
        .filter(|value| *value <= 0xFF_FFFF)
        .ok_or_else(|| format!("{field} barcode colour must be from 0 through 0xFFFFFF"))
}

struct FormulaParser {
    characters: Vec<char>,
    index: usize,
    token_count: usize,
    depth: usize,
}

impl FormulaParser {
    fn new(input: &str) -> std::result::Result<Self, String> {
        if input.len() > MAX_FORMULA_BYTES {
            return Err("formula exceeds the 4096-byte limit".to_owned());
        }
        if !input.is_ascii() {
            return Err("formula contains unsupported non-ASCII syntax".to_owned());
        }
        Ok(Self {
            characters: input.chars().collect(),
            index: 0,
            token_count: 0,
            depth: 0,
        })
    }

    fn parse(mut self) -> std::result::Result<f64, String> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if self.index != self.characters.len() {
            return Err("formula contains unsupported or trailing syntax".to_owned());
        }
        if !value.is_finite() {
            return Err("formula result is outside the finite numeric range".to_owned());
        }
        Ok(value)
    }

    fn parse_comparison(&mut self) -> std::result::Result<f64, String> {
        let left = self.parse_additive()?;
        self.skip_whitespace();
        let operator = ["<=", ">=", "<>", "=", "<", ">"]
            .into_iter()
            .find(|operator| self.remaining().starts_with(operator));
        let Some(operator) = operator else {
            return Ok(left);
        };
        self.index += operator.len();
        self.count_token()?;
        let right = self.parse_additive()?;
        Ok(
            if match operator {
                "=" => left == right,
                "<>" => left != right,
                "<" => left < right,
                "<=" => left <= right,
                ">" => left > right,
                ">=" => left >= right,
                _ => unreachable!(),
            } {
                1.0
            } else {
                0.0
            },
        )
    }

    fn parse_additive(&mut self) -> std::result::Result<f64, String> {
        let mut value = self.parse_multiplicative()?;
        loop {
            self.skip_whitespace();
            let operator = self.peek();
            if !matches!(operator, Some('+') | Some('-')) {
                return Ok(value);
            }
            self.index += 1;
            self.count_token()?;
            let right = self.parse_multiplicative()?;
            value = if operator == Some('+') {
                value + right
            } else {
                value - right
            };
            self.ensure_finite(value)?;
        }
    }

    fn parse_multiplicative(&mut self) -> std::result::Result<f64, String> {
        let mut value = self.parse_power()?;
        loop {
            self.skip_whitespace();
            let operator = self.peek();
            if !matches!(operator, Some('*') | Some('/')) {
                return Ok(value);
            }
            self.index += 1;
            self.count_token()?;
            let right = self.parse_power()?;
            if operator == Some('/') && right == 0.0 {
                return Err("formula divides by zero".to_owned());
            }
            value = match operator {
                Some('*') => value * right,
                Some('/') => value / right,
                _ => unreachable!(),
            };
            self.ensure_finite(value)?;
        }
    }

    fn parse_power(&mut self) -> std::result::Result<f64, String> {
        let value = self.parse_unary()?;
        self.skip_whitespace();
        if self.peek() != Some('^') {
            return Ok(value);
        }
        self.index += 1;
        self.count_token()?;
        let exponent = self.parse_power()?;
        let result = value.powf(exponent);
        self.ensure_finite(result)?;
        Ok(result)
    }

    fn parse_unary(&mut self) -> std::result::Result<f64, String> {
        self.skip_whitespace();
        match self.peek() {
            Some('+') => {
                self.index += 1;
                self.count_token()?;
                self.parse_unary()
            }
            Some('-') => {
                self.index += 1;
                self.count_token()?;
                Ok(-self.parse_unary()?)
            }
            _ => self.parse_percentage(),
        }
    }

    fn parse_percentage(&mut self) -> std::result::Result<f64, String> {
        let mut value = self.parse_primary()?;
        loop {
            self.skip_whitespace();
            if self.peek() != Some('%') {
                return Ok(value);
            }
            self.index += 1;
            self.count_token()?;
            value /= 100.0;
            self.ensure_finite(value)?;
        }
    }

    fn parse_primary(&mut self) -> std::result::Result<f64, String> {
        self.skip_whitespace();
        if self.peek() == Some('(') {
            self.index += 1;
            self.count_token()?;
            self.depth += 1;
            if self.depth > MAX_FORMULA_DEPTH {
                return Err("formula exceeds the 32-level nesting limit".to_owned());
            }
            let value = self.parse_comparison()?;
            self.skip_whitespace();
            if self.peek() != Some(')') {
                return Err("formula has an unclosed parenthesis".to_owned());
            }
            self.index += 1;
            self.count_token()?;
            self.depth -= 1;
            return Ok(value);
        }
        let start = self.index;
        while self
            .peek()
            .is_some_and(|character| character.is_ascii_digit() || matches!(character, '.' | ','))
        {
            self.index += 1;
        }
        if start == self.index {
            return Err(
                if self
                    .peek()
                    .is_some_and(|character| character.is_ascii_alphabetic())
                {
                    "formula functions are unsupported".to_owned()
                } else {
                    "formula requires a numeric operand".to_owned()
                },
            );
        }
        self.count_token()?;
        let number = self.characters[start..self.index]
            .iter()
            .filter(|character| **character != ',')
            .collect::<String>();
        number
            .parse::<f64>()
            .ok()
            .filter(|value| value.is_finite())
            .ok_or_else(|| "formula contains an invalid or out-of-range number".to_owned())
    }

    fn remaining(&self) -> String {
        self.characters[self.index..].iter().collect()
    }

    fn peek(&self) -> Option<char> {
        self.characters.get(self.index).copied()
    }

    fn skip_whitespace(&mut self) {
        while self.peek().is_some_and(char::is_whitespace) {
            self.index += 1;
        }
    }

    fn count_token(&mut self) -> std::result::Result<(), String> {
        self.token_count += 1;
        if self.token_count > MAX_FORMULA_TOKENS {
            Err("formula exceeds the 512-token limit".to_owned())
        } else {
            Ok(())
        }
    }

    fn ensure_finite(&self, value: f64) -> std::result::Result<(), String> {
        if value.is_finite() {
            Ok(())
        } else {
            Err("formula result is outside the finite numeric range".to_owned())
        }
    }
}

fn format_formula_value(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        return format!("{value:.0}");
    }
    format!("{value:.14e}")
        .parse::<f64>()
        .unwrap_or(value)
        .to_string()
}

fn text_argument(instruction: &FieldInstruction, index: usize) -> Option<&str> {
    instruction.arguments.get(index).and_then(argument_text)
}

fn argument_text(argument: &FieldArgument) -> Option<&str> {
    match argument {
        FieldArgument::Text(value) => Some(value),
        FieldArgument::Nested(_) => None,
    }
}

fn has_switch(instruction: &FieldInstruction, name: &str) -> bool {
    instruction
        .switches
        .iter()
        .any(|switch| switch.name == name)
}

fn switch_text<'a>(instruction: &'a FieldInstruction, name: &str) -> Option<&'a str> {
    instruction.switches.iter().find_map(|switch| {
        (switch.name == name)
            .then_some(switch.argument.as_ref())
            .flatten()
            .and_then(argument_text)
    })
}

fn unsupported_switch(instruction: &FieldInstruction) -> Option<&str> {
    let allowed: &[&str] = match instruction.name.as_str() {
        "PAGE" | "NUMPAGES" => &["*", "#"],
        "REF" => &["h", "*", "#"],
        "PAGEREF" => &["h", "p", "*", "#"],
        "IF" => &["*", "#"],
        "SEQ" => &["n", "c", "h", "r", "s", "*", "#"],
        "DOCPROPERTY" | "DOCVARIABLE" => &["*", "#", "@"],
        "STYLEREF" => &["l", "n", "r", "t", "w", "p", "*", "#"],
        "INCLUDETEXT" => &["!", "c", "*", "#"],
        "DATE" | "TIME" => &["@", "*"],
        "FILENAME" => &["p", "*"],
        "AUTHOR" => &["*"],
        "MERGEFIELD" => &["b", "f", "m", "v", "*", "#", "@"],
        "=" => &["*", "#"],
        "NEXT" | "NEXTIF" | "SKIPIF" | "MERGEREC" | "MERGESEQ" => &[],
        _ => return None,
    };
    instruction
        .switches
        .iter()
        .find(|switch| !allowed.contains(&switch.name.as_str()))
        .map(|switch| switch.name.as_str())
}

fn validate_instruction_shape(instruction: &FieldInstruction) -> std::result::Result<(), String> {
    if !instruction_quotes_are_balanced(&instruction.raw) {
        return Err(format!("field {} has unclosed quoting", instruction.name));
    }
    let argument_range = match instruction.name.as_str() {
        "PAGE" | "NUMPAGES" | "DATE" | "TIME" | "FILENAME" | "AUTHOR" => 0..=0,
        "REF" | "PAGEREF" | "SEQ" | "DOCPROPERTY" | "DOCVARIABLE" | "STYLEREF" | "MERGEFIELD" => {
            1..=1
        }
        "IF" => 5..=5,
        "INCLUDETEXT" => 1..=2,
        "=" => 1..=usize::MAX,
        "TOC" => 0..=0,
        "TC" => 1..=1,
        "DISPLAYBARCODE" | "MERGEBARCODE" => 2..=2,
        "NEXT" | "MERGEREC" | "MERGESEQ" => 0..=0,
        "NEXTIF" | "SKIPIF" => 3..=3,
        _ => return Ok(()),
    };
    if !argument_range.contains(&instruction.arguments.len()) {
        let expected = if argument_range.start() == argument_range.end() {
            argument_range.start().to_string()
        } else {
            format!(
                "{} through {}",
                argument_range.start(),
                argument_range.end()
            )
        };
        return Err(format!(
            "field {} requires {expected} positional operands",
            instruction.name
        ));
    }

    if matches!(
        instruction.name.as_str(),
        "TOC" | "TC" | "DISPLAYBARCODE" | "MERGEBARCODE"
    ) {
        return Ok(());
    }

    for switch in &instruction.switches {
        let requires_text = matches!(
            switch.name.as_str(),
            "*" | "#" | "@" | "r" | "s" | "b" | "f"
        ) || instruction.name == "INCLUDETEXT" && switch.name == "c";
        match (&switch.argument, requires_text) {
            (Some(FieldArgument::Text(_)), true) | (None, false) => {}
            (_, true) => {
                return Err(format!(
                    "field {} switch \\{} requires a text argument",
                    instruction.name, switch.name
                ));
            }
            (Some(_), false) => {
                return Err(format!(
                    "field {} switch \\{} does not take an argument",
                    instruction.name, switch.name
                ));
            }
        }
    }
    Ok(())
}

fn instruction_quotes_are_balanced(raw: &str) -> bool {
    let mut characters = raw.chars().peekable();
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

fn compare_if(left: &str, operator: &str, right: &str) -> Option<bool> {
    let numeric = left
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
        .zip(right.parse::<f64>().ok().filter(|value| value.is_finite()));
    let ordering = numeric
        .map(|(left, right)| left.total_cmp(&right))
        .unwrap_or_else(|| left.to_lowercase().cmp(&right.to_lowercase()));
    match operator {
        "=" => Some(if right.contains(['*', '?']) {
            wildcard_matches(right, left)
        } else {
            ordering.is_eq()
        }),
        "<>" => Some(if right.contains(['*', '?']) {
            !wildcard_matches(right, left)
        } else {
            !ordering.is_eq()
        }),
        "<" => Some(ordering.is_lt()),
        "<=" => Some(!ordering.is_gt()),
        ">" => Some(ordering.is_gt()),
        ">=" => Some(!ordering.is_lt()),
        _ => None,
    }
}

fn wildcard_matches(pattern: &str, value: &str) -> bool {
    let mut expression = String::from("(?is)^");
    for character in pattern.chars() {
        match character {
            '*' => expression.push_str(".*"),
            '?' => expression.push('.'),
            other => expression.push_str(&regex::escape(&other.to_string())),
        }
    }
    expression.push('$');
    regex::Regex::new(&expression).is_ok_and(|regex| regex.is_match(value))
}

fn custom_property_text(value: &CustomPropertyValue) -> Option<String> {
    match value {
        CustomPropertyValue::Lpstr(value)
        | CustomPropertyValue::Lpwstr(value)
        | CustomPropertyValue::FileTime(value) => Some(value.clone()),
        CustomPropertyValue::I4(value) => Some(value.to_string()),
        CustomPropertyValue::R8(value) if value.is_finite() => Some(value.to_string()),
        CustomPropertyValue::Bool(value) => Some(value.to_string()),
        CustomPropertyValue::Empty => Some(String::new()),
        CustomPropertyValue::R8(_) | CustomPropertyValue::Raw(_) => None,
    }
}

fn normalized_name(value: &str) -> String {
    value
        .chars()
        .filter(|character| character.is_ascii_alphanumeric())
        .flat_map(char::to_lowercase)
        .collect()
}

fn lexical_file_name(path: &str) -> Option<&str> {
    path.rsplit(['/', '\\'])
        .find(|component| !component.is_empty())
}

fn apply_formats(
    instruction: &FieldInstruction,
    value: &str,
    date_time: Option<FieldDateTime>,
) -> std::result::Result<String, String> {
    let mut output = value.to_owned();
    if let Some(picture) = switch_text(instruction, "#") {
        let number = value
            .parse::<f64>()
            .ok()
            .filter(|number| number.is_finite())
            .ok_or_else(|| "numeric field value is not a finite number".to_owned())?;
        output = format_numeric_picture(number, picture)?;
    }
    if let Some(picture) = switch_text(instruction, "@") {
        let date_time = if matches!(instruction.name.as_str(), "DATE" | "TIME") {
            date_time.ok_or_else(|| {
                "date-time formatting requires an explicit date and time".to_owned()
            })?
        } else {
            parse_field_date_time(value).ok_or_else(|| {
                "date-time field value is not a supported civil date and time".to_owned()
            })?
        };
        output = format_date_time(date_time, picture)?;
    }
    for format in instruction.switches.iter().filter_map(|switch| {
        (switch.name == "*")
            .then_some(switch.argument.as_ref())
            .flatten()
            .and_then(argument_text)
    }) {
        output = apply_general_format(&output, format)?;
    }
    Ok(output)
}

fn apply_general_format(value: &str, format: &str) -> std::result::Result<String, String> {
    if format == "ALPHABETIC" {
        return parse_positive_integer(value).map(|value| alphabetic(value, true));
    }
    if format == "ROMAN" {
        return parse_positive_integer(value).and_then(|value| roman(value, true));
    }
    match format.to_ascii_lowercase().as_str() {
        "upper" => Ok(value.to_uppercase()),
        "lower" => Ok(value.to_lowercase()),
        "firstcap" => Ok(capitalize_first(value)),
        "caps" => Ok(value
            .split_inclusive(char::is_whitespace)
            .map(capitalize_first)
            .collect()),
        "arabic" => parse_positive_integer(value).map(|value| value.to_string()),
        "alphabetic" => parse_positive_integer(value).map(|value| alphabetic(value, false)),
        "roman" => parse_positive_integer(value).and_then(|value| roman(value, false)),
        "ordinal" => parse_positive_integer(value).map(ordinal),
        "mergeformat" | "charformat" => Ok(value.to_owned()),
        other => Err(format!("general format {other} is unsupported")),
    }
}

fn parse_positive_integer(value: &str) -> std::result::Result<u32, String> {
    value
        .trim()
        .parse::<u32>()
        .ok()
        .filter(|value| *value > 0)
        .ok_or_else(|| "general numeric format requires a positive integer".to_owned())
}

fn capitalize_first(value: &str) -> String {
    let mut characters = value.chars();
    characters
        .next()
        .map(|first| first.to_uppercase().chain(characters).collect())
        .unwrap_or_default()
}

fn alphabetic(mut value: u32, upper: bool) -> String {
    let mut output = Vec::new();
    while value > 0 {
        value -= 1;
        let base = if upper { b'A' } else { b'a' };
        output.push((base + (value % 26) as u8) as char);
        value /= 26;
    }
    output.iter().rev().collect()
}

fn roman(mut value: u32, upper: bool) -> std::result::Result<String, String> {
    if value > 3999 {
        return Err("Roman format supports values from 1 through 3999".to_owned());
    }
    let values = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut output = String::new();
    for (number, numeral) in values {
        while value >= number {
            output.push_str(numeral);
            value -= number;
        }
    }
    Ok(if upper { output } else { output.to_lowercase() })
}

fn ordinal(value: u32) -> String {
    let suffix = if (11..=13).contains(&(value % 100)) {
        "th"
    } else {
        match value % 10 {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th",
        }
    };
    format!("{value}{suffix}")
}

fn split_picture_sections(picture: &str) -> std::result::Result<Vec<String>, String> {
    let mut sections = vec![String::new()];
    let mut quote = None;
    for character in picture.chars() {
        match (character, quote) {
            ('"' | '\'', None) => {
                quote = Some(character);
                sections.last_mut().unwrap().push(character);
            }
            (character, Some(active)) if character == active => {
                quote = None;
                sections.last_mut().unwrap().push(character);
            }
            (';', None) => sections.push(String::new()),
            (other, _) => sections.last_mut().unwrap().push(other),
        }
    }
    if quote.is_some() || sections.len() > 3 {
        Err("numeric picture has invalid sections or quoting".to_owned())
    } else {
        Ok(sections)
    }
}

fn format_numeric_picture(value: f64, picture: &str) -> std::result::Result<String, String> {
    let sections = split_picture_sections(picture)?;
    let (section, magnitude, implicit_negative) = if value < 0.0 {
        if let Some(section) = sections.get(1) {
            (section.as_str(), -value, false)
        } else {
            (sections[0].as_str(), -value, true)
        }
    } else if value == 0.0 && sections.len() == 3 {
        (sections[2].as_str(), 0.0, false)
    } else {
        (sections[0].as_str(), value, false)
    };
    let placeholders = numeric_placeholder_indices(section);
    let Some(first) = placeholders.first().copied() else {
        return unquote_picture(section);
    };
    let last = placeholders.last().copied().unwrap() + 1;
    let prefix = unquote_picture(&section[..first])?;
    let suffix = unquote_picture(&section[last..])?;
    let (core, literals) = extract_numeric_literals(&section[first..last])?;
    let mut halves = core.split('.');
    let integer_picture = halves.next().unwrap_or_default();
    let decimal_picture = halves.next().unwrap_or_default();
    if halves.next().is_some()
        || integer_picture
            .chars()
            .any(|character| !matches!(character, '0' | '#' | ','))
        || decimal_picture
            .chars()
            .any(|character| !matches!(character, '0' | '#'))
    {
        return Err("numeric picture contains unsupported tokens".to_owned());
    }
    let maximum_decimals = decimal_picture.chars().count();
    let minimum_decimals = decimal_picture
        .chars()
        .filter(|character| *character == '0')
        .count();
    let integer_placeholders = integer_picture
        .chars()
        .filter(|character| matches!(character, '0' | '#'))
        .collect::<Vec<_>>();
    let decimal_placeholders = decimal_picture.chars().collect::<Vec<_>>();
    let formatted = format!("{magnitude:.maximum_decimals$}");
    let (integer, mut decimal) = formatted
        .split_once('.')
        .map(|(integer, decimal)| (integer.to_owned(), decimal.to_owned()))
        .unwrap_or((formatted, String::new()));
    while decimal.len() > minimum_decimals && decimal.ends_with('0') {
        decimal.pop();
    }
    let mut integer = integer;
    if !integer_placeholders.contains(&'0') && integer == "0" {
        integer.clear();
    }
    let missing_integer = integer_placeholders.len().saturating_sub(integer.len());
    let padding = integer_placeholders[..missing_integer]
        .iter()
        .map(|placeholder| if *placeholder == '0' { '0' } else { ' ' })
        .collect::<String>();
    if integer_picture.contains(',') {
        integer = group_digits(&integer);
    }
    let decimal_padding = decimal_placeholders[decimal.len()..]
        .iter()
        .map(|placeholder| if *placeholder == '0' { '0' } else { ' ' })
        .collect::<String>();
    let number = if decimal_placeholders.is_empty() {
        format!("{padding}{integer}")
    } else {
        format!("{padding}{integer}.{decimal}{decimal_padding}")
    };
    let number = insert_numeric_literals(
        &number,
        integer_picture
            .chars()
            .chain(decimal_picture.chars())
            .filter(|character| matches!(character, '0' | '#'))
            .count(),
        &literals,
    );
    Ok(format!(
        "{}{}{}{}",
        if implicit_negative { "-" } else { "" },
        prefix,
        number,
        suffix
    ))
}

fn numeric_placeholder_indices(value: &str) -> Vec<usize> {
    let mut quote = None;
    let mut indices = Vec::new();
    for (index, character) in value.char_indices() {
        match (character, quote) {
            ('"' | '\'', None) => quote = Some(character),
            (character, Some(active)) if character == active => quote = None,
            ('0' | '#', None) => indices.push(index),
            _ => {}
        }
    }
    indices
}

fn extract_numeric_literals(
    value: &str,
) -> std::result::Result<(String, Vec<(usize, String)>), String> {
    let mut pattern = String::new();
    let mut literals = Vec::new();
    let mut literal = String::new();
    let mut quote = None;
    let mut placeholder_count = 0usize;
    for character in value.chars() {
        match (character, quote) {
            ('"' | '\'', None) => quote = Some(character),
            (character, Some(active)) if character == active => {
                literals.push((placeholder_count, std::mem::take(&mut literal)));
                quote = None;
            }
            (character, Some(_)) => literal.push(character),
            (character @ ('0' | '#'), None) => {
                pattern.push(character);
                placeholder_count += 1;
            }
            (character, None) => pattern.push(character),
        }
    }
    if quote.is_some() {
        Err("numeric picture has unclosed quoting".to_owned())
    } else {
        Ok((pattern, literals))
    }
}

fn insert_numeric_literals(
    number: &str,
    placeholder_count: usize,
    literals: &[(usize, String)],
) -> String {
    if literals.is_empty() {
        return number.to_owned();
    }
    let slot_count = number
        .chars()
        .filter(|character| character.is_ascii_digit() || *character == ' ')
        .count();
    let extra_leading = slot_count.saturating_sub(placeholder_count);
    let mut output = String::new();
    let mut digits_written = 0usize;
    for (position, literal) in literals
        .iter()
        .filter(|(position, _)| extra_leading + position == 0)
    {
        let _ = position;
        output.push_str(literal);
    }
    for character in number.chars() {
        output.push(character);
        if character.is_ascii_digit() || character == ' ' {
            digits_written += 1;
            for (_, literal) in literals
                .iter()
                .filter(|(position, _)| extra_leading + position == digits_written)
            {
                output.push_str(literal);
            }
        }
    }
    output
}

fn unquote_picture(value: &str) -> std::result::Result<String, String> {
    let mut output = String::new();
    let mut quote = None;
    for character in value.chars() {
        match (character, quote) {
            ('"' | '\'', None) => quote = Some(character),
            (character, Some(active)) if character == active => quote = None,
            (character, _) => output.push(character),
        }
    }
    if quote.is_some() {
        Err("numeric picture has unclosed quoting".to_owned())
    } else {
        Ok(output)
    }
}

fn group_digits(value: &str) -> String {
    let mut output = String::new();
    for (index, character) in value.chars().enumerate() {
        if index > 0 && (value.len() - index).is_multiple_of(3) {
            output.push(',');
        }
        output.push(character);
    }
    output
}

fn valid_date_time(value: FieldDateTime) -> bool {
    value.month >= 1
        && value.month <= 12
        && value.day >= 1
        && value.day <= days_in_month(value.year, value.month)
        && value.hour < 24
        && value.minute < 60
        && value.second < 60
}

fn parse_field_date_time(value: &str) -> Option<FieldDateTime> {
    let value = value.trim().strip_suffix('Z').unwrap_or(value.trim());
    let (date, time) = value
        .split_once('T')
        .or_else(|| value.split_once(' '))
        .map_or((value, None), |(date, time)| (date, Some(time)));
    let mut date_parts = date.split('-');
    let year = date_parts.next()?.parse().ok()?;
    let month = date_parts.next()?.parse().ok()?;
    let day = date_parts.next()?.parse().ok()?;
    if date_parts.next().is_some() {
        return None;
    }
    let (hour, minute, second) = if let Some(time) = time {
        let mut time_parts = time.split(':');
        let hour = time_parts.next()?.parse().ok()?;
        let minute = time_parts.next()?.parse().ok()?;
        let second = time_parts.next()?.split('.').next()?.parse().ok()?;
        if time_parts.next().is_some() {
            return None;
        }
        (hour, minute, second)
    } else {
        (0, 0, 0)
    };
    let parsed = FieldDateTime {
        year,
        month,
        day,
        hour,
        minute,
        second,
    };
    valid_date_time(parsed).then_some(parsed)
}

fn days_in_month(year: i32, month: u8) -> u8 {
    match month {
        4 | 6 | 9 | 11 => 30,
        2 if year % 400 == 0 || year % 4 == 0 && year % 100 != 0 => 29,
        2 => 28,
        _ => 31,
    }
}

fn format_date_time(value: FieldDateTime, picture: &str) -> std::result::Result<String, String> {
    if !valid_date_time(value) {
        return Err("field date and time is invalid".to_owned());
    }
    let months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    let weekdays = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let characters = picture.chars().collect::<Vec<_>>();
    let mut output = String::new();
    let mut index = 0;
    while index < characters.len() {
        if characters[index] == '"' {
            index += 1;
            while index < characters.len() && characters[index] != '"' {
                output.push(characters[index]);
                index += 1;
            }
            if index == characters.len() {
                return Err("date-time picture has unclosed quoting".to_owned());
            }
            index += 1;
            continue;
        }
        if characters[index] == '\\' {
            index += 1;
            let Some(character) = characters.get(index) else {
                return Err("date-time picture ends with an escape".to_owned());
            };
            output.push(*character);
            index += 1;
            continue;
        }
        let rest = characters[index..].iter().collect::<String>();
        if rest.to_ascii_uppercase().starts_with("AM/PM") {
            output.push_str(if value.hour < 12 { "AM" } else { "PM" });
            index += 5;
            continue;
        }
        let token = characters[index];
        if !matches!(token, 'y' | 'M' | 'd' | 'H' | 'h' | 'm' | 's') {
            output.push(token);
            index += 1;
            continue;
        }
        let mut count = 1;
        while index + count < characters.len() && characters[index + count] == token {
            count += 1;
        }
        match token {
            'y' if count == 2 => output.push_str(&format!("{:02}", value.year.rem_euclid(100))),
            'y' => output.push_str(&format!("{:04}", value.year)),
            'M' if count == 1 => output.push_str(&value.month.to_string()),
            'M' if count == 2 => output.push_str(&format!("{:02}", value.month)),
            'M' if count == 3 => output.push_str(&months[value.month as usize - 1][..3]),
            'M' => output.push_str(months[value.month as usize - 1]),
            'd' if count == 1 => output.push_str(&value.day.to_string()),
            'd' if count == 2 => output.push_str(&format!("{:02}", value.day)),
            'd' if count == 3 => output.push_str(&weekdays[weekday(value)][..3]),
            'd' => output.push_str(weekdays[weekday(value)]),
            'H' if count == 1 => output.push_str(&value.hour.to_string()),
            'H' => output.push_str(&format!("{:02}", value.hour)),
            'h' if count == 1 => output.push_str(&(value.hour % 12).max(1).to_string()),
            'h' => output.push_str(&format!("{:02}", (value.hour % 12).max(1))),
            'm' if count == 1 => output.push_str(&value.minute.to_string()),
            'm' => output.push_str(&format!("{:02}", value.minute)),
            's' if count == 1 => output.push_str(&value.second.to_string()),
            's' => output.push_str(&format!("{:02}", value.second)),
            _ => unreachable!(),
        }
        index += count;
    }
    Ok(output)
}

fn weekday(value: FieldDateTime) -> usize {
    let mut year = i64::from(value.year);
    let mut month = i64::from(value.month);
    if month < 3 {
        month += 12;
        year -= 1;
    }
    let year_of_century = year.rem_euclid(100);
    let century = year.div_euclid(100);
    let h = (i64::from(value.day)
        + (13 * (month + 1)) / 5
        + year_of_century
        + year_of_century / 4
        + century / 4
        + 5 * century)
        .rem_euclid(7);
    ((h + 6) % 7) as usize
}

#[cfg(test)]
mod tests {
    use rdocx_oxml::document::BodyContent;
    use rdocx_oxml::properties::CT_PPr;
    use rdocx_oxml::text::{CT_P, CT_R, Field, FieldSwitch, RunContent};

    use super::*;

    fn document_with_fields(fields: &[(&str, &str)]) -> Document {
        let mut document = Document::new();
        let mut paragraph = CT_P::new();
        for (instruction, cached_result) in fields {
            paragraph.runs.push(CT_R {
                properties: None,
                content: vec![RunContent::Field(Field::new(instruction, cached_result))],
                extra_xml: Vec::new(),
                extra_xml_positions: Vec::new(),
                alt_drawings: Vec::new(),
            });
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        document
    }

    fn document_with_parsed_paragraph(xml: &str) -> Document {
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let paragraph = loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Event::Eof => panic!("missing paragraph"),
                _ => {}
            }
            buffer.clear();
        };
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        document
    }

    #[test]
    fn a_typed_field_is_reported_without_mutating_its_cache() {
        let document = document_with_fields(&[("AUTHOR", "stored")]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 1, "the typed field must be evaluated");
    }

    #[test]
    fn formula_fields_use_bounded_precedence_and_stable_failures() {
        let document = document_with_fields(&[
            ("= 2 + 3 * 4", "precedence"),
            ("= (2 + 3) * 4", "parentheses"),
            ("= 2 ^ 3 ^ 2", "power"),
            ("= 1,000 + .5", "grouping"),
            ("= 50%", "percentage"),
            ("= (50 + 50)%", "grouped percentage"),
            (r#"= 7 / 2 \# "0.00""#, "picture"),
            ("= 1 / 0", "zero"),
            ("= SUM(1, 2)", "function"),
            ("= (1 + 2", "malformed"),
            ("= 1e309", "bounds"),
        ]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results
                .iter()
                .take(7)
                .map(|result| result.outcome.clone())
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("14".to_owned()),
                FieldOutcome::Resolved("20".to_owned()),
                FieldOutcome::Resolved("512".to_owned()),
                FieldOutcome::Resolved("1000.5".to_owned()),
                FieldOutcome::Resolved("0.5".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("3.50".to_owned()),
            ]
        );
        for result in &results[7..] {
            assert!(
                matches!(result.outcome, FieldOutcome::KeepStored { .. }),
                "{} unexpectedly resolved as {:?}",
                result.instruction,
                result.outcome
            );
        }
        assert_eq!(results[7].outcome, keep("formula divides by zero"));
        assert_eq!(
            results[8].outcome,
            keep("formula functions are unsupported")
        );
        assert_eq!(
            results[9].outcome,
            keep("formula has an unclosed parenthesis")
        );
        let invalid_percentage = document_with_fields(&[("= 5 % 2", "stored")]);
        assert_eq!(
            invalid_percentage
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            keep("formula contains unsupported or trailing syntax")
        );
        let normalized_decimal = document_with_fields(&[("= 0.1 + 0.2", "stored")]);
        assert_eq!(
            normalized_decimal
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("0.3".to_owned())
        );
        let compact = format!(
            "= {}",
            std::iter::repeat_n("1", 129).collect::<Vec<_>>().join("+")
        );
        let spaced = format!(
            "= {}",
            std::iter::repeat_n("1", 129)
                .collect::<Vec<_>>()
                .join(" + ")
        );
        for instruction in [compact, spaced] {
            let equivalent = document_with_fields(&[(&instruction, "stored")]);
            assert_eq!(
                equivalent
                    .evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::Resolved("129".to_owned())
            );
        }
        for instruction in [
            format!("= {}", "1".repeat(MAX_FORMULA_BYTES + 1)),
            format!(
                "= {}1{}",
                "(".repeat(MAX_FORMULA_DEPTH + 1),
                ")".repeat(MAX_FORMULA_DEPTH + 1)
            ),
            format!("= {}1", "1+".repeat(MAX_FORMULA_TOKENS)),
        ] {
            let bounded = document_with_fields(&[(&instruction, "bounded fallback")]);
            let result = bounded
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap();
            assert!(matches!(result[0].outcome, FieldOutcome::KeepStored { .. }));
            assert_eq!(result[0].cached_result, "bounded fallback");
        }

        let mut nested = document_with_fields(&[("= 1 + 2", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut nested.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(formula) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        formula.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Amount", "nested")));
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Amount".to_owned(), "3".to_owned())]),
            ..Default::default()
        };
        let results = nested.evaluate_fields(&context).unwrap();
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("5".to_owned()));
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("3".to_owned()));
    }

    #[test]
    fn mail_merge_control_state_is_story_and_record_scoped() {
        let document = document_with_fields(&[
            ("NEXT", "next"),
            ("MERGEREC", "record"),
            (r#"NEXTIF "A" = "B""#, "conditional next"),
            (r#"SKIPIF "A" = "A""#, "conditional skip"),
            ("MERGESEQ", "sequence"),
        ]);
        let context = FieldEvaluationContext {
            merge_record_number: Some(4),
            merge_sequence_number: Some(2),
            ..Default::default()
        };
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(
            results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecord { record_number: 5 }),
                FieldOutcome::MailMergeControl(MailMergeControl::RecordNumber(5)),
                FieldOutcome::MailMergeControl(MailMergeControl::NextRecordIf {
                    condition: false,
                    record_number: 5,
                }),
                FieldOutcome::MailMergeControl(MailMergeControl::SkipRecordIf {
                    condition: true,
                    record_number: 5,
                }),
                FieldOutcome::MailMergeControl(MailMergeControl::SequenceNumber(2)),
            ]
        );

        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            unreachable!()
        };
        let paragraphs = [paragraph];
        let mut evaluator = Evaluator::new(&document, &context);
        evaluator.evaluate_story("main", &paragraphs);
        evaluator.evaluate_story("header:one", &paragraphs);
        assert_eq!(
            evaluator.results[0].outcome, evaluator.results[5].outcome,
            "each story must start from the explicit record context"
        );

        let unavailable = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(
            unavailable
                .iter()
                .all(|result| matches!(result.outcome, FieldOutcome::KeepStored { .. }))
        );

        let record_only = FieldEvaluationContext {
            merge_record_number: Some(9),
            ..Default::default()
        };
        let record = document_with_fields(&[("MERGEREC", "stored")]);
        assert_eq!(
            record.evaluate_fields(&record_only).unwrap()[0].outcome,
            FieldOutcome::MailMergeControl(MailMergeControl::RecordNumber(9))
        );

        for (instruction, context, diagnostic) in [
            (
                "MERGEREC",
                FieldEvaluationContext {
                    merge_record_number: Some(0),
                    ..Default::default()
                },
                "MERGEREC merge record number must be one-based",
            ),
            (
                "MERGESEQ",
                FieldEvaluationContext {
                    merge_sequence_number: Some(0),
                    ..Default::default()
                },
                "MERGESEQ merge sequence number must be one-based",
            ),
        ] {
            let invalid = document_with_fields(&[(instruction, "stored")]);
            assert_eq!(
                invalid.evaluate_fields(&context).unwrap()[0].outcome,
                keep(diagnostic)
            );
        }
    }

    #[test]
    fn toc_tc_and_barcode_fields_preserve_non_text_results() {
        let document = document_with_fields(&[
            (
                r#"TOC \o "1-3" \t "Heading 1,1,Appendix,2" \f C \b Main \h \u \n "2-3" \p " " \d ":""#,
                "stored toc",
            ),
            (r#"TC "Entry" \l 2 \f C \n"#, "stored tc"),
            (
                r#"DISPLAYBARCODE "0123456789012" EAN13 \h 100 \s 200 \f 0xFF0000 \b 0xFFFFFF \t"#,
                "stored barcode",
            ),
        ]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results[0].outcome,
            FieldOutcome::TableOfContents(TocField {
                heading_levels: Some((1, 3)),
                custom_styles: vec![("Heading 1".to_owned(), 1), ("Appendix".to_owned(), 2)],
                entries: TocEntrySelection::Identifier("C".to_owned()),
                sequence_identifier: None,
                bookmark: Some("Main".to_owned()),
                hyperlink: true,
                use_outline_levels: true,
                omit_page_number_levels: Some((2, 3)),
                page_number_separator: Some(" ".to_owned()),
                entry_page_separator: Some(":".to_owned()),
            })
        );
        assert_eq!(
            results[1].outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: "Entry".to_owned(),
                level: 2,
                table_identifier: Some("C".to_owned()),
                omit_page_number: true,
            })
        );
        assert_eq!(
            results[2].outcome,
            FieldOutcome::Barcode(BarcodeField {
                value: "0123456789012".to_owned(),
                kind: BarcodeKind::Ean13,
                height: Some(100),
                scale: Some(200),
                error_correction: None,
                point_of_sale_style: None,
                case_style: None,
                fix_check_digit: false,
                rotation: None,
                foreground_color: Some(0xFF0000),
                background_color: Some(0xFFFFFF),
                display_text: true,
                add_start_stop: false,
            })
        );

        let unsupported = document_with_fields(&[
            (r#"TOC \o "9-1""#, "toc"),
            (r#"TC "Entry" \l 10"#, "tc"),
            ("DISPLAYBARCODE value UNKNOWN", "barcode"),
            ("DISPLAYBARCODE 123 EAN13", "barcode digits"),
        ]);
        let outcomes = unsupported
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            outcomes
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                keep("TOC heading range starts after it ends"),
                keep("TC level must be from 1 through 9"),
                keep("DISPLAYBARCODE barcode type UNKNOWN is unsupported"),
                keep("DISPLAYBARCODE EAN13 value has an invalid digit count or character"),
            ]
        );

        for (instruction, heading_levels, entries, use_outline_levels) in [
            ("TOC", Some((1, 9)), TocEntrySelection::None, false),
            (r"TOC \o", Some((1, 9)), TocEntrySelection::None, false),
            (r"TOC \f", None, TocEntrySelection::All, false),
            (
                r"TOC \f C",
                None,
                TocEntrySelection::Identifier("C".to_owned()),
                false,
            ),
            (r"TOC \u", None, TocEntrySelection::None, true),
        ] {
            let toc = document_with_fields(&[(instruction, "stored toc")]);
            assert_eq!(
                toc.evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::TableOfContents(TocField {
                    heading_levels,
                    custom_styles: Vec::new(),
                    entries,
                    sequence_identifier: None,
                    bookmark: None,
                    hyperlink: false,
                    use_outline_levels,
                    omit_page_number_levels: None,
                    page_number_separator: None,
                    entry_page_separator: None,
                })
            );
        }

        let normalized_toc =
            document_with_fields(&[(r#"TOC \t "Heading 1,1, Appendix,2" \p ":""#, "stored toc")]);
        assert_eq!(
            normalized_toc
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                heading_levels: None,
                custom_styles: vec![("Heading 1".to_owned(), 1), ("Appendix".to_owned(), 2)],
                entries: TocEntrySelection::None,
                sequence_identifier: None,
                bookmark: None,
                hyperlink: false,
                use_outline_levels: false,
                omit_page_number_levels: None,
                page_number_separator: Some(":".to_owned()),
                entry_page_separator: None,
            })
        );
        let decorated_default = document_with_fields(&[(r"TOC \h", "stored toc")]);
        assert!(matches!(
            decorated_default
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                heading_levels: Some((1, 9)),
                hyperlink: true,
                ..
            })
        ));
        let invalid_toc = document_with_fields(&[(r#"TOC \p "ab""#, "stored toc")]);
        assert_eq!(
            invalid_toc
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            keep("TOC page-number separator must contain exactly one character")
        );
        let sequenced_toc =
            document_with_fields(&[(r#"TOC \o "1-3" \s chapter \d ":""#, "stored toc")]);
        assert!(matches!(
            sequenced_toc
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContents(TocField {
                sequence_identifier: Some(ref identifier),
                entry_page_separator: Some(ref separator),
                ..
            }) if identifier == "chapter" && separator == ":"
        ));

        let barcode_grammar = document_with_fields(&[
            (r"DISPLAYBARCODE 0123456789012 EAN13 \x", "fix"),
            (r"DISPLAYBARCODE 0123456789012 EAN13 \p STD", "pos"),
            (r"DISPLAYBARCODE 1234567890123 ITF14 \c EXT", "case"),
            (r"DISPLAYBARCODE value QR \q 3", "correction"),
            (r"DISPLAYBARCODE 0123456789012 EAN13 \p OTHER", "bad pos"),
            (r"DISPLAYBARCODE value QR \q H", "bad correction"),
        ]);
        let outcomes = barcode_grammar
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        let FieldOutcome::Barcode(fix) = &outcomes[0].outcome else {
            panic!("expected check-digit barcode")
        };
        assert!(fix.fix_check_digit);
        let FieldOutcome::Barcode(pos) = &outcomes[1].outcome else {
            panic!("expected point-of-sale barcode")
        };
        assert_eq!(
            pos.point_of_sale_style,
            Some(BarcodePointOfSaleStyle::Standard)
        );
        let FieldOutcome::Barcode(case) = &outcomes[2].outcome else {
            panic!("expected ITF14 case barcode")
        };
        assert_eq!(case.case_style, Some(BarcodeCaseStyle::Extended));
        let FieldOutcome::Barcode(correction) = &outcomes[3].outcome else {
            panic!("expected corrected QR barcode")
        };
        assert_eq!(correction.error_correction, Some(3));
        assert_eq!(
            outcomes[4].outcome,
            keep("DISPLAYBARCODE point-of-sale style must be STD, SUP2, SUP5, or CASE")
        );
        assert_eq!(
            outcomes[5].outcome,
            keep("DISPLAYBARCODE error correction must be from 0 through 3")
        );

        let case_alias = document_with_fields(&[
            (r"DISPLAYBARCODE 1234567890123 CASE \c EXT", "case"),
            (r"DISPLAYBARCODE not-digits CASE", "invalid case"),
        ]);
        let outcomes = case_alias
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(matches!(
            outcomes[0].outcome,
            FieldOutcome::Barcode(BarcodeField {
                kind: BarcodeKind::Case,
                case_style: Some(BarcodeCaseStyle::Extended),
                ..
            })
        ));
        assert_eq!(
            outcomes[1].outcome,
            keep("DISPLAYBARCODE CASE value has an invalid digit count or character")
        );

        let extra_operands = document_with_fields(&[
            (r#"TC "Entry" unexpected"#, "tc"),
            (r"DISPLAYBARCODE value QR unexpected \t", "barcode"),
        ]);
        let outcomes = extra_operands
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            outcomes[0].outcome,
            keep("field TC requires 1 positional operands")
        );
        assert_eq!(
            outcomes[1].outcome,
            keep("field DISPLAYBARCODE requires 2 positional operands")
        );

        let escaped = document_with_fields(&[
            (r#"TC "A\"B""#, "stored quote"),
            (r#"TC "\Entry""#, "stored slash"),
        ]);
        assert_eq!(
            escaped
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: "A\"B".to_owned(),
                level: 1,
                table_identifier: None,
                omit_page_number: false,
            })
        );
        assert_eq!(
            escaped
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[1]
                .outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                entry: r"\Entry".to_owned(),
                level: 1,
                table_identifier: None,
                omit_page_number: false,
            })
        );

        let long_value = "x".repeat(1025);
        let barcode_bounds = document_with_fields(&[
            (&format!("DISPLAYBARCODE {long_value} QR"), "value"),
            (r"DISPLAYBARCODE value QR \h 4294967296", "height"),
            (r"DISPLAYBARCODE value QR \s 9", "scale low"),
            (r"DISPLAYBARCODE value QR \s 1001", "scale high"),
            (r"DISPLAYBARCODE value QR \r 4", "rotation"),
            (r"DISPLAYBARCODE value QR \f 0x1000000", "foreground"),
            (r"DISPLAYBARCODE value QR \b 16777216", "background"),
        ]);
        let outcomes = barcode_bounds
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            outcomes
                .into_iter()
                .map(|outcome| outcome.outcome)
                .collect::<Vec<_>>(),
            [
                keep("DISPLAYBARCODE value must contain from 1 through 1024 characters"),
                keep("DISPLAYBARCODE height must be a nonnegative integer"),
                keep("DISPLAYBARCODE scale must be from 10 through 1000"),
                keep("DISPLAYBARCODE scale must be from 10 through 1000"),
                keep("DISPLAYBARCODE rotation must be from 0 through 3"),
                keep("DISPLAYBARCODE barcode colour must be from 0 through 0xFFFFFF"),
                keep("DISPLAYBARCODE barcode colour must be from 0 through 0xFFFFFF"),
            ]
        );
        let maximum_value = "x".repeat(1024);
        let boundary_instruction = format!(
            "DISPLAYBARCODE {maximum_value} QR \\h 4294967295 \\s 10 \\r 3 \\f 0 \\b 0xFFFFFF"
        );
        let boundaries = document_with_fields(&[(&boundary_instruction, "stored")]);
        assert!(matches!(
            boundaries
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Barcode(BarcodeField {
                height: Some(u32::MAX),
                scale: Some(10),
                rotation: Some(3),
                foreground_color: Some(0),
                background_color: Some(0xFF_FFFF),
                ..
            })
        ));

        let mut nested = document_with_fields(&[("DISPLAYBARCODE placeholder QR", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut nested.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(barcode) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        barcode.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Code", "nested")));
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Code".to_owned(), "nested value".to_owned())]),
            ..Default::default()
        };
        assert!(matches!(
            nested.evaluate_fields(&context).unwrap()[0].outcome,
            FieldOutcome::Barcode(BarcodeField { ref value, .. }) if value == "nested value"
        ));

        let mut nested_tc = document_with_fields(&[("TC placeholder", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut nested_tc.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(tc) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        tc.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Entry", "nested")));
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Entry".to_owned(), "Nested entry".to_owned())]),
            ..Default::default()
        };
        assert!(matches!(
            nested_tc.evaluate_fields(&context).unwrap()[0].outcome,
            FieldOutcome::TableOfContentsEntry(TcField { ref entry, .. })
                if entry == "Nested entry"
        ));

        let mut nested_switches = document_with_fields(&[
            (r"TOC \b placeholder", "toc"),
            (r#"TC "Entry" \f placeholder"#, "tc"),
            (r"DISPLAYBARCODE value QR \s 100", "barcode"),
        ]);
        let BodyContent::Paragraph(paragraph) = &mut nested_switches.document.body.content[0]
        else {
            unreachable!()
        };
        for (index, instruction) in ["MERGEFIELD Scope", "MERGEFIELD Kind", "MERGEFIELD Scale"]
            .into_iter()
            .enumerate()
        {
            let RunContent::Field(field) = &mut paragraph.runs[index].content[0] else {
                unreachable!()
            };
            field.instruction.switches[0].argument = Some(FieldArgument::Nested(Box::new(
                Field::new(instruction, "nested"),
            )));
        }
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([
                ("Scope".to_owned(), "Main".to_owned()),
                ("Kind".to_owned(), "C".to_owned()),
                ("Scale".to_owned(), "250".to_owned()),
            ]),
            ..Default::default()
        };
        let outcomes = nested_switches.evaluate_fields(&context).unwrap();
        assert!(matches!(
            outcomes[0].outcome,
            FieldOutcome::TableOfContents(TocField {
                bookmark: Some(ref bookmark),
                ..
            }) if bookmark == "Main"
        ));
        assert!(matches!(
            outcomes[2].outcome,
            FieldOutcome::TableOfContentsEntry(TcField {
                table_identifier: Some(ref identifier),
                ..
            }) if identifier == "C"
        ));
        assert!(matches!(
            outcomes[4].outcome,
            FieldOutcome::Barcode(BarcodeField {
                scale: Some(250),
                ..
            })
        ));
    }

    #[test]
    fn raw_only_instruction_edits_use_the_serialized_instruction() {
        let mut document = document_with_fields(&[("DATE", "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        field.instruction.raw = "AUTHOR".to_owned();
        document.set_author("Ada Lovelace");

        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results[0].instruction, "AUTHOR");
        assert_eq!(
            results[0].outcome,
            FieldOutcome::Resolved("Ada Lovelace".to_owned())
        );
    }

    #[test]
    fn same_run_nested_raw_instruction_edit_updates_saves_and_reopens() {
        let word_namespace = rdocx_oxml::namespace::W_NS;
        let xml = format!(
            concat!(
                r#"<w:p xmlns:w="{0}"><q:r xmlns:q="{0}" xmlns:x="urn:producer">"#,
                r#"<q:fldChar q:fldCharType="begin"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText>MERGEFIELD Old</q:instrText><x:nestedInstruction/><q:fldChar q:fldCharType="separate"/><q:t>stored nested</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"</q:r></w:p>"#,
            ),
            word_namespace,
        );
        let mut document = document_with_parsed_paragraph(&xml);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        let Some(FieldArgument::Nested(nested)) = outer.instruction.arguments.first_mut() else {
            panic!("expected nested field")
        };
        nested.instruction.raw = "MERGEFIELD New".to_owned();

        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("New".to_owned(), "new value".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        assert_eq!(document.update_fields(&context).unwrap(), 2);
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains("MERGEFIELD New"), "{saved_xml}");
        assert!(!saved_xml.contains("MERGEFIELD Old"), "{saved_xml}");
        assert_eq!(saved_xml.matches(">new value<").count(), 1, "{saved_xml}");
        assert!(saved_xml.contains(r#"w:dirty="0""#), "{saved_xml}");
        assert!(saved_xml.contains("<x:nestedInstruction/>"), "{saved_xml}");

        let reopened = Document::from_bytes(&saved).unwrap();
        let evaluations = reopened.evaluate_fields(&context).unwrap();
        assert_eq!(evaluations.len(), 2);
        assert_eq!(evaluations[0].cached_result, "yes");
        assert_eq!(evaluations[1].instruction, "MERGEFIELD New");
        assert_eq!(evaluations[1].cached_result, "new value");
        assert_eq!(
            evaluations[1].outcome,
            FieldOutcome::Resolved("new value".to_owned())
        );
    }

    #[test]
    fn raw_only_nested_edit_suppresses_its_stale_nested_operand() {
        let word_namespace = rdocx_oxml::namespace::W_NS;
        let xml = format!(
            concat!(
                r#"<w:p xmlns:w="{0}"><q:r xmlns:q="{0}" xmlns:x="urn:producer">"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText><x:middleBefore/>"#,
                r#"<q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText>MERGEFIELD Stale</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored grandchild</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<x:middleAfter/><q:instrText xml:space="preserve"> = &quot;stale&quot; &quot;old yes&quot; &quot;old no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored middle</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"<q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/>"#,
                r#"</q:r></w:p>"#,
            ),
            word_namespace,
        );
        let mut document = document_with_parsed_paragraph(&xml);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        let Some(FieldArgument::Nested(middle)) = outer.instruction.arguments.first_mut() else {
            panic!("expected middle field")
        };
        assert!(matches!(
            middle.instruction.arguments.first(),
            Some(FieldArgument::Nested(_))
        ));
        middle.instruction.raw = "MERGEFIELD New".to_owned();

        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("New".to_owned(), "new value".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        let before_save = document.evaluate_fields(&context).unwrap();
        assert_eq!(before_save.len(), 2);
        assert_eq!(before_save[0].field_index, 0);
        assert_eq!(before_save[1].field_index, 1);
        assert_eq!(before_save[1].instruction, "MERGEFIELD New");
        assert_eq!(
            before_save[1].outcome,
            FieldOutcome::Resolved("new value".to_owned())
        );

        assert_eq!(document.update_fields(&context).unwrap(), 2);
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains("MERGEFIELD New"), "{saved_xml}");
        assert!(!saved_xml.contains("MERGEFIELD Stale"), "{saved_xml}");
        assert!(!saved_xml.contains("stored grandchild"), "{saved_xml}");
        assert!(!saved_xml.contains("stored middle"), "{saved_xml}");
        assert!(saved_xml.contains("<x:middleBefore/>"), "{saved_xml}");
        assert!(saved_xml.contains("<x:middleAfter/>"), "{saved_xml}");
        assert!(
            saved_xml.find("<x:middleBefore/>").unwrap()
                < saved_xml.find("<x:middleAfter/>").unwrap(),
            "{saved_xml}"
        );
        assert_eq!(saved_xml.matches(r#"w:dirty="0""#).count(), 2);
        assert_eq!(saved_xml.matches(">new value<").count(), 1, "{saved_xml}");

        let reopened = Document::from_bytes(&saved).unwrap();
        let evaluations = reopened.evaluate_fields(&context).unwrap();
        assert_eq!(evaluations.len(), 2);
        assert_eq!(evaluations[0].field_index, 0);
        assert_eq!(evaluations[0].cached_result, "yes");
        assert_eq!(evaluations[1].field_index, 1);
        assert_eq!(evaluations[1].instruction, "MERGEFIELD New");
        assert_eq!(evaluations[1].cached_result, "new value");
        assert_eq!(
            evaluations[1].outcome,
            FieldOutcome::Resolved("new value".to_owned())
        );
    }

    #[test]
    fn multi_run_raw_only_nested_edit_preserves_every_run_scaffold() {
        let word_namespace = rdocx_oxml::namespace::W_NS;
        let xml = format!(
            concat!(
                r#"<w:p xmlns:w="{0}">"#,
                r#"<q:r xmlns:q="{0}" data-run="outer-start"><q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:a="urn:start" data-run="start"><q:rPr><q:i/></q:rPr><q:fldChar q:fldCharType="begin" q:dirty="on"/><q:instrText xml:space="preserve">IF </q:instrText><a:prefix/><q:fldChar q:fldCharType="begin" q:dirty="on"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:m="urn:middle" data-run="middle"><q:rPr><q:u q:val="single"/></q:rPr><q:instrText>MERGEFIELD Stale</q:instrText><m:inside/></q:r>"#,
                r#"<q:r xmlns:q="{0}" xmlns:z="urn:end" data-run="end"><q:rPr><q:b/></q:rPr><q:fldChar q:fldCharType="separate"/><q:t>stored grandchild</q:t><q:fldChar q:fldCharType="end"/><z:suffix/><q:instrText xml:space="preserve"> = &quot;stale&quot; &quot;old yes&quot; &quot;old no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored middle</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
                r#"<q:r xmlns:q="{0}" data-run="outer-end"><q:instrText xml:space="preserve"> = &quot;new value&quot; &quot;yes&quot; &quot;no&quot;</q:instrText><q:fldChar q:fldCharType="separate"/><q:t>stored outer</q:t><q:fldChar q:fldCharType="end"/></q:r>"#,
                r#"</w:p>"#,
            ),
            word_namespace,
        );
        let mut document = document_with_parsed_paragraph(&xml);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        let Some(FieldArgument::Nested(middle)) = outer.instruction.arguments.first_mut() else {
            panic!("expected middle field")
        };
        assert!(matches!(
            middle.instruction.arguments.first(),
            Some(FieldArgument::Nested(_))
        ));
        middle.instruction.raw = "MERGEFIELD New".to_owned();

        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("New".to_owned(), "new value".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        assert_eq!(document.update_fields(&context).unwrap(), 2);
        let saved = document.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(&saved)).unwrap();
        let saved_xml =
            std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(saved_xml.contains("MERGEFIELD New"), "{saved_xml}");
        assert!(!saved_xml.contains("MERGEFIELD Stale"), "{saved_xml}");
        assert!(!saved_xml.contains("stored grandchild"), "{saved_xml}");
        for preserved in [
            r#"data-run="start""#,
            r#"data-run="middle""#,
            r#"data-run="end""#,
            r#"xmlns:a="urn:start""#,
            r#"xmlns:m="urn:middle""#,
            r#"xmlns:z="urn:end""#,
            "<q:i/>",
            r#"<q:u q:val="single"/>"#,
            "<q:b/>",
            "<a:prefix/>",
            "<m:inside/>",
            "<z:suffix/>",
        ] {
            assert!(
                saved_xml.contains(preserved),
                "missing {preserved}: {saved_xml}"
            );
        }
        assert!(
            saved_xml.find("<a:prefix/>").unwrap() < saved_xml.find("<m:inside/>").unwrap()
                && saved_xml.find("<m:inside/>").unwrap() < saved_xml.find("<z:suffix/>").unwrap(),
            "{saved_xml}"
        );

        let reopened = Document::from_bytes(&saved).unwrap();
        let evaluations = reopened.evaluate_fields(&context).unwrap();
        assert_eq!(evaluations.len(), 2);
        assert_eq!(evaluations[0].field_index, 0);
        assert_eq!(evaluations[0].cached_result, "yes");
        assert_eq!(evaluations[1].field_index, 1);
        assert_eq!(evaluations[1].instruction, "MERGEFIELD New");
        assert_eq!(evaluations[1].cached_result, "new value");
    }

    #[test]
    fn nested_if_and_comparison_operators_evaluate_recursively() {
        let mut document = document_with_fields(&[(r#"IF left = right yes no"#, "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("MERGEFIELD Score", "stored score")));
        outer.instruction.arguments[1] = FieldArgument::Text(">=".to_owned());
        outer.instruction.arguments[2] = FieldArgument::Text("2".to_owned());
        outer.instruction.arguments[4] =
            FieldArgument::Nested(Box::new(Field::new("UNKNOWN", "stored branch")));
        let mut context = FieldEvaluationContext::default();
        context
            .merge_fields
            .insert("Score".to_owned(), "10".to_owned());
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(results[0].field_index, 0);
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("yes".to_owned()));
        assert_eq!(results[1].field_index, 1);
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("10".to_owned()));
        assert_eq!(results[2].field_index, 2);
        assert!(matches!(
            results[2].outcome,
            FieldOutcome::KeepStored { .. }
        ));

        for (operator, expected) in [
            ("=", "yes"),
            ("<>", "no"),
            ("<", "no"),
            ("<=", "yes"),
            (">", "no"),
            (">=", "yes"),
        ] {
            let document = document_with_fields(&[(
                &format!(r#"IF "2" {operator} "2" "yes" "no""#),
                "stored",
            )]);
            assert_eq!(
                document
                    .evaluate_fields(&FieldEvaluationContext::default())
                    .unwrap()[0]
                    .outcome,
                FieldOutcome::Resolved(expected.to_owned())
            );
        }
        let wildcard =
            document_with_fields(&[(r#"IF "Alphabet" = "A?pha*" "yes" "no""#, "stored")]);
        assert_eq!(
            wildcard
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("yes".to_owned())
        );

        let mut unresolved = document_with_fields(&[(r#"IF left = right yes no"#, "outer")]);
        let BodyContent::Paragraph(paragraph) = &mut unresolved.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("DATE", "stored date")));
        let outcomes = unresolved
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert!(matches!(
            outcomes[0].outcome,
            FieldOutcome::KeepStored { .. }
        ));
        assert!(matches!(
            outcomes[1].outcome,
            FieldOutcome::KeepStored { .. }
        ));
    }

    #[test]
    fn nested_if_reuses_the_eager_effective_instruction_outcome() {
        let mut document = document_with_fields(&[(r#"IF left = 1 yes no"#, "outer")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("SEQ Figure", "stored sequence")));

        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 2);
        assert_eq!(results[0].outcome, FieldOutcome::Resolved("yes".to_owned()));
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("1".to_owned()));
    }

    #[test]
    fn nested_outcome_frames_do_not_leak_between_outer_fields() {
        let mut document = document_with_fields(&[
            (r#"IF left = 1 first no"#, "first outer"),
            (r#"IF left = 2 second no"#, "second outer"),
        ]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        for run in &mut paragraph.runs {
            let RunContent::Field(outer) = &mut run.content[0] else {
                unreachable!()
            };
            outer.instruction.arguments[0] =
                FieldArgument::Nested(Box::new(Field::new("SEQ Figure", "stored sequence")));
        }

        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 4);
        assert_eq!(
            results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("first".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("second".to_owned()),
                FieldOutcome::Resolved("2".to_owned()),
            ]
        );
    }

    #[test]
    fn every_nested_field_is_reported_before_outer_fallback() {
        let mut document = document_with_fields(&[("UNKNOWN", "outer")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer
            .instruction
            .arguments
            .push(FieldArgument::Nested(Box::new(Field::new(
                "MERGEFIELD Name",
                "stored name",
            ))));
        outer.instruction.switches.push(FieldSwitch {
            name: "*".to_owned(),
            argument: Some(FieldArgument::Nested(Box::new(Field::new(
                "AUTHOR",
                "stored author",
            )))),
        });
        let context = FieldEvaluationContext {
            merge_fields: BTreeMap::from([("Name".to_owned(), "Ada".to_owned())]),
            ..FieldEvaluationContext::default()
        };
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(results.len(), 3);
        assert!(matches!(
            results[0].outcome,
            FieldOutcome::KeepStored { .. }
        ));
        assert_eq!(results[1].outcome, FieldOutcome::Resolved("Ada".to_owned()));
        assert!(matches!(
            results[2].outcome,
            FieldOutcome::KeepStored { .. }
        ));
    }

    #[test]
    fn malformed_arity_and_extreme_inputs_keep_stored_without_panicking() {
        let document = document_with_fields(&[
            (r"DATE \@", "date"),
            (r"SEQ Figure \r", "reset"),
            ("PAGE extra", "page"),
            (r#"MERGEFIELD "Name"#, "merge"),
            (r##"MERGEFIELD Amount \# "$0.00"##, "picture"),
            (r"SEQ Figure \r 9223372036854775807", "max"),
            ("SEQ Figure", "overflow"),
        ]);
        let mut context = FieldEvaluationContext::default();
        context
            .merge_fields
            .insert("Name".to_owned(), "Ada".to_owned());
        context
            .merge_fields
            .insert("Amount".to_owned(), "12".to_owned());
        let results = document.evaluate_fields(&context).unwrap();
        assert!(
            results[..5]
                .iter()
                .all(|result| matches!(result.outcome, FieldOutcome::KeepStored { .. }))
        );
        assert_eq!(
            results[5].outcome,
            FieldOutcome::Resolved(i64::MAX.to_string())
        );
        assert_eq!(results[6].outcome, keep("SEQ value overflowed"));

        let date = document_with_fields(&[(r#"DATE \@ "dddd""#, "date")]);
        let context = FieldEvaluationContext {
            now: Some(FieldDateTime {
                year: i32::MIN,
                month: 1,
                day: 1,
                hour: 0,
                minute: 0,
                second: 0,
            }),
            ..FieldEvaluationContext::default()
        };
        assert!(matches!(
            date.evaluate_fields(&context).unwrap()[0].outcome,
            FieldOutcome::Resolved(_)
        ));
    }

    #[test]
    fn sequence_state_is_scoped_and_reset_by_supported_switches() {
        let document = document_with_fields(&[
            ("SEQ Figure", "0"),
            (r"SEQ Figure \c", "0"),
            (r"SEQ Figure \r 5", "0"),
            (r"SEQ Figure \h", "0"),
            (r"SEQ Figure \n \* ROMAN", "0"),
            ("SEQ Table", "0"),
        ]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(
            results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("5".to_owned()),
                FieldOutcome::Resolved(String::new()),
                FieldOutcome::Resolved("VII".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
            ]
        );

        let story_document = document_with_fields(&[("SEQ Shared", "0")]);
        let BodyContent::Paragraph(paragraph) = &story_document.document.body.content[0] else {
            unreachable!()
        };
        let paragraphs = [paragraph];
        let story_context = FieldEvaluationContext::default();
        let mut evaluator = Evaluator::new(&story_document, &story_context);
        evaluator.evaluate_story("header:one", &paragraphs);
        evaluator.evaluate_story("footer:one", &paragraphs);
        assert_eq!(
            evaluator
                .results
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
            ]
        );

        let mut heading_restart = Document::new();
        for (style_id, instruction) in [
            (Some("Heading1"), None),
            (None, Some(r"SEQ Figure \s 1")),
            (None, Some("SEQ Figure")),
            (Some("Heading1"), None),
            (None, Some(r"SEQ Figure \s 1")),
        ] {
            let mut paragraph = CT_P::new();
            paragraph.properties = style_id.map(|style_id| CT_PPr {
                style_id: Some(style_id.to_owned()),
                outline_lvl: Some(0),
                ..Default::default()
            });
            if let Some(instruction) = instruction {
                paragraph.runs.push(CT_R {
                    properties: None,
                    content: vec![RunContent::Field(Field::new(instruction, "stored"))],
                    extra_xml: Vec::new(),
                    extra_xml_positions: Vec::new(),
                    alt_drawings: Vec::new(),
                });
            }
            heading_restart
                .document
                .body
                .content
                .push(BodyContent::Paragraph(paragraph));
        }
        assert_eq!(
            heading_restart
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()
                .into_iter()
                .map(|result| result.outcome)
                .collect::<Vec<_>>(),
            [
                FieldOutcome::Resolved("1".to_owned()),
                FieldOutcome::Resolved("2".to_owned()),
                FieldOutcome::Resolved("1".to_owned()),
            ]
        );
    }

    #[test]
    fn missing_context_and_unsupported_fields_keep_their_cached_display() {
        let document = document_with_fields(&[("DATE", "stored"), ("UNKNOWN", "stored")]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 2, "fallback outcomes must still be reported");
    }

    #[test]
    fn document_properties_variables_and_author_use_package_values() {
        let document = document_with_fields(&[("AUTHOR", "stored")]);
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 1, "package-backed fields must be reported");
    }

    #[test]
    fn styleref_searches_the_approved_direction_and_scope() {
        let mut document = Document::new();
        document.add_style(style::StyleBuilder::paragraph("Heading1", "Heading 1"));
        let mut source = CT_P::new();
        source.properties = Some(CT_PPr {
            style_id: Some("Heading1".to_owned()),
            num_id: Some(7),
            num_ilvl: Some(0),
            ..Default::default()
        });
        source.runs.push(CT_R::new("numbered heading"));
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(source));
        let mut field = CT_P::new();
        field.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new(
                r#"STYLEREF "Heading 1" \n"#,
                "stored",
            ))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(field));
        let results = document
            .evaluate_fields(&FieldEvaluationContext::default())
            .unwrap();
        assert_eq!(results.len(), 1, "style fields must be reported");
        assert_eq!(
            results[0].outcome,
            keep("STYLEREF numbered source formatting is unsupported")
        );

        let BodyContent::Paragraph(source) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        source.properties.as_mut().unwrap().num_id = Some(0);
        assert_eq!(
            document
                .evaluate_fields(&FieldEvaluationContext::default())
                .unwrap()[0]
                .outcome,
            FieldOutcome::Resolved("numbered heading".to_owned())
        );
    }

    #[test]
    fn date_time_filename_mergefield_and_includetext_use_only_explicit_context() {
        let document = document_with_fields(&[("FILENAME", "stored")]);
        let context = FieldEvaluationContext {
            file_name: Some("report.docx".to_owned()),
            ..FieldEvaluationContext::default()
        };
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(
            results[0].outcome,
            FieldOutcome::Resolved("report.docx".to_owned())
        );
    }

    #[test]
    fn formatting_switches_match_the_pinned_word_matrix() {
        let document = document_with_fields(&[(r#"MERGEFIELD Name \* Upper"#, "stored")]);
        let mut context = FieldEvaluationContext::default();
        context
            .merge_fields
            .insert("Name".to_owned(), "field value".to_owned());
        let results = document.evaluate_fields(&context).unwrap();
        assert_eq!(
            results[0].outcome,
            FieldOutcome::Resolved("FIELD VALUE".to_owned())
        );
        assert_eq!(apply_general_format("FiELD", "Lower").unwrap(), "field");
        assert_eq!(
            apply_general_format("field VALUE", "FirstCap").unwrap(),
            "Field VALUE"
        );
        assert_eq!(
            apply_general_format("field value", "Caps").unwrap(),
            "Field Value"
        );
        assert_eq!(apply_general_format("27", "Arabic").unwrap(), "27");
        assert_eq!(apply_general_format("same", "MERGEFORMAT").unwrap(), "same");
        assert_eq!(apply_general_format("same", "Charformat").unwrap(), "same");
        assert_eq!(apply_general_format("27", "alphabetic").unwrap(), "aa");
        assert_eq!(apply_general_format("27", "ALPHABETIC").unwrap(), "AA");
        assert_eq!(apply_general_format("14", "roman").unwrap(), "xiv");
        assert_eq!(apply_general_format("14", "ROMAN").unwrap(), "XIV");
        assert_eq!(apply_general_format("22", "Ordinal").unwrap(), "22nd");
        assert_eq!(
            format_numeric_picture(1234.5, "#,##0.00").unwrap(),
            "1,234.50"
        );
        assert_eq!(
            format_numeric_picture(-12.0, "$0.00;($0.00);\"zero\"").unwrap(),
            "($12.00)"
        );
        assert_eq!(
            format_numeric_picture(0.0, "$0.00;($0.00);\"zero\"").unwrap(),
            "zero"
        );
        assert_eq!(format_numeric_picture(0.0, "#").unwrap(), " ");
        assert_eq!(format_numeric_picture(15.0, "$###").unwrap(), "$ 15");
        assert_eq!(
            format_numeric_picture(12.5, "$##0.00 'is sales tax'").unwrap(),
            "$ 12.50 is sales tax"
        );
        assert_eq!(format_numeric_picture(15.0, "#'x'##").unwrap(), " x15");
        assert_eq!(
            format_numeric_picture(123456.0, "000'-'000").unwrap(),
            "123-456"
        );
        let now = FieldDateTime {
            year: 2025,
            month: 12,
            day: 14,
            hour: 21,
            minute: 7,
            second: 5,
        };
        assert_eq!(
            format_date_time(now, "dddd, MMMM d, yyyy HH:mm:ss AM/PM").unwrap(),
            "Sunday, December 14, 2025 21:07:05 PM"
        );
        let property = Field::new(r#"DOCPROPERTY SavedAt \@ "MMMM d, yyyy""#, "stored");
        assert_eq!(
            apply_formats(&property.instruction, "2025-12-14T21:07:05Z", None).unwrap(),
            "December 14, 2025"
        );
        let merge = Field::new(r#"MERGEFIELD Date \@ "yyyy-MM-dd""#, "stored");
        assert_eq!(
            apply_formats(&merge.instruction, "2026-01-02", None).unwrap(),
            "2026-01-02"
        );
    }

    #[test]
    fn wildcard_if_matches_multiline_nested_ref_values() {
        let mut document =
            document_with_fields(&[(r#"IF left = "first*second" yes no"#, "stored")]);
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!()
        };
        let RunContent::Field(outer) = &mut paragraph.runs[0].content[0] else {
            unreachable!()
        };
        outer.instruction.arguments[0] =
            FieldArgument::Nested(Box::new(Field::new("REF Multi", "stored ref")));

        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            unreachable!()
        };
        let paragraphs = [paragraph];
        let context = FieldEvaluationContext::default();
        let mut evaluator = Evaluator::new(&document, &context);
        evaluator
            .bookmarks
            .insert("Multi".to_owned(), "first\nsecond".to_owned());
        evaluator.evaluate_story("main", &paragraphs);
        assert_eq!(
            evaluator.results[0].outcome,
            FieldOutcome::Resolved("yes".to_owned())
        );
        assert_eq!(
            evaluator.results[1].outcome,
            FieldOutcome::Resolved("first\nsecond".to_owned())
        );
    }
}
