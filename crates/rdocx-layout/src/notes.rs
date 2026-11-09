//! Footnote and endnote content, laid out once before pagination.
//!
//! Notes used to be laid out inside the post-pagination pass that drew them,
//! which meant pagination could not know how much room they would need and
//! drew body text straight over them. Laying them out here, ahead of
//! pagination, lets the paginator reserve exactly the height it will later
//! draw. Reserve and render read the same lines, so they cannot disagree.
//!
//! The marker is shaped here too. The paginator only holds `&FontManager` and
//! shaping needs `&mut`, so a note that arrives pre-shaped is a note the
//! paginator can place without touching a font.

use std::collections::HashMap;

use rdocx_oxml::footnotes::NoteType;
use rdocx_oxml::styles::CT_Styles;

use crate::WordStory;
use crate::block::ParagraphBlock;
use crate::engine::{
    MainStoryLayoutItem, SourceRegistry, body_layout_items,
    layout_paragraph_with_source_and_direction, sect_pr_to_geometry,
};
use crate::input::{LayoutInput, MediaRegistry};
use crate::style_resolver::NumberingState;
use oxml_layout::{
    Color, Diagnostic, FontManager, LayoutLine, NoteRef, NoteStream, Result, TextDirection,
    TextSegment,
};

/// Point size notes are set at.
const NOTE_FONT_SIZE: f64 = 8.0;
/// Legacy note indent constant. Actual note layout uses authored paragraph indents.
pub const NOTE_INDENT: f64 = 12.0;
/// Vertical gap between the separator rule and the first note line.
pub const NOTE_SEPARATOR_OFFSET: f64 = 6.0;
/// Width of the rule above a note that starts on its own page, as a fraction
/// of the content width.
pub const SEPARATOR_WIDTH_FRACTION: f64 = 0.33;

/// One note, laid out and ready to place.
#[derive(Debug, Clone)]
pub struct NoteLayout {
    /// The pre-shaped superscript number drawn at the start of the note.
    pub marker: TextSegment,
    /// Special separator and notice stories supply all visible content themselves.
    pub draw_marker: bool,
    /// How far above the baseline the marker sits.
    pub marker_rise: f64,
    /// The note's lines, flattened across its paragraphs.
    pub lines: Vec<LayoutLine>,
    /// Line ranges belonging to paragraphs with a visible tracked revision,
    /// with the colour of their change bar.
    pub revision_ranges: Vec<(std::ops::Range<usize>, Color)>,
}

impl NoteLayout {
    /// Height of a range of this note's lines.
    pub fn height_of(&self, first: usize, count: usize) -> f64 {
        self.lines
            .iter()
            .skip(first)
            .take(count)
            .map(|line| line.height)
            .sum()
    }

    /// Height of every line from `first` onward.
    pub fn height_from(&self, first: usize) -> f64 {
        self.height_of(first, self.lines.len())
    }

    /// Total height of every line.
    pub fn height(&self) -> f64 {
        self.height_from(0)
    }
}

/// A note, and the content width it was broken to.
///
/// The width is held as raw bits because `f64` is not `Hash`. Both the key and
/// every lookup come from `PageGeometry::content_width()` over the same
/// `sectPr`, so this is exact equality on a value that was computed the same
/// way twice, not a comparison that needs a tolerance.
type NoteKey = (NoteRef, u64);

/// Every note the document defines, laid out once per distinct width.
#[derive(Debug, Clone, Default)]
pub struct NoteRegistry {
    notes: HashMap<NoteKey, NoteEntry>,
    specials: HashMap<(NoteStream, NoteType, u64), NoteEntry>,
    continuation_separator: [bool; 2],
}

#[derive(Debug, Clone)]
struct NoteEntry {
    layout: NoteLayout,
    paragraphs: Vec<NoteRenderParagraph>,
}

#[derive(Debug, Clone)]
pub(crate) struct NoteRenderParagraph {
    pub paint: std::sync::Arc<crate::block::WordParagraphPaint>,
    pub border_gap_before: f64,
    pub block: ParagraphBlock,
    pub direction: TextDirection,
    pub lines: std::ops::Range<usize>,
}

impl NoteRegistry {
    /// Lay out every note in the footnote and endnote streams, once for each
    /// distinct content width the document paginates at.
    ///
    /// Notes use the full source width and resolved paragraph indents.
    /// The width that matters is the one belonging to the
    /// section carrying the reference rather than the document's last section.
    /// `content_widths` may repeat, and a repeat costs nothing: the common
    /// document, whose sections share a page size, lays each note out once.
    pub(crate) fn build(
        input: &LayoutInput,
        styles: &CT_Styles,
        media: &MediaRegistry,
        fm: &mut FontManager,
        num_state: &mut NumberingState,
        content_widths: &[f64],
        diagnostics: &mut Vec<Diagnostic>,
        sources: Option<&SourceRegistry>,
    ) -> Result<Self> {
        let previous_context = num_state.note_body_context;
        let result = (|| {
            let mut notes = HashMap::new();
            let mut specials = HashMap::new();
            let mut continuation_separator = [false; 2];

            // Each stream is keyed separately, so a document numbering a footnote
            // and an endnote alike keeps both.
            for (kind, stream) in [
                (NoteStream::Footnote, input.footnotes.as_ref()),
                (NoteStream::Endnote, input.endnotes.as_ref()),
            ]
            .into_iter()
            .filter_map(|(kind, stream)| stream.map(|stream| (kind, stream)))
            {
                let family = usize::from(kind == NoteStream::Endnote);
                let selected = input.note_defaults[family]
                    .as_ref()
                    .map(|properties| properties.special_references.as_slice())
                    .unwrap_or_default();
                for note in &stream.footnotes {
                    let special = (note.note_type != NoteType::Normal).then_some(note.note_type);
                    if let Some(record) = special {
                        let conventional = match record {
                            NoteType::Separator => -1,
                            NoteType::ContinuationSeparator => 0,
                            NoteType::ContinuationNotice => -2,
                            NoteType::Normal => unreachable!(),
                        };
                        if !selected.contains(&note.id)
                            && !(selected.is_empty() && note.id == conventional)
                        {
                            continue;
                        }
                        if record == NoteType::ContinuationSeparator {
                            continuation_separator[family] = true;
                        }
                    }
                    let note_ref = NoteRef {
                        stream: kind,
                        id: note.id,
                    };

                    // Laying the same note out again must not consume its list
                    // numbers again, so every width after the first starts from the
                    // counters the first one started from. Numbering does not
                    // depend on the width, so the state left behind is the state a
                    // single layout would have left.
                    let counters_before = num_state.clone();
                    let mut laid_out = false;

                    for &content_width in content_widths {
                        let key = (note_ref, content_width.to_bits());
                        if special.is_some_and(|record| {
                            specials.contains_key(&(kind, record, content_width.to_bits()))
                        }) || (special.is_none() && notes.contains_key(&key))
                        {
                            continue;
                        }
                        if laid_out {
                            *num_state = counters_before.clone();
                        }
                        laid_out = true;
                        let note_width = content_width.max(1.0);
                        num_state.note_body_context = Some(note_ref);

                        let mut lines = Vec::new();
                        let mut revision_ranges = Vec::new();
                        let mut render_paragraphs = Vec::new();
                        let story = match kind {
                            NoteStream::Footnote => WordStory::Footnote { id: note.id },
                            NoteStream::Endnote => WordStory::Endnote { id: note.id },
                        };
                        let items = input
                            .story_bodies
                            .get(&story)
                            .map(body_layout_items)
                            .unwrap_or_else(|| {
                                note.paragraphs
                                    .iter()
                                    .enumerate()
                                    .map(|(index, paragraph)| {
                                        MainStoryLayoutItem::Paragraph(paragraph, vec![index])
                                    })
                                    .collect()
                            });
                        for item in items {
                            let laid_out_blocks = match item {
                                MainStoryLayoutItem::Paragraph(paragraph, path) => {
                                    let source =
                                        sources.and_then(|sources| sources.id(&story, &path));
                                    let (mut block, direction, mut paint) =
                                        layout_paragraph_with_source_and_direction(
                                            paragraph,
                                            note_width,
                                            styles,
                                            input,
                                            media,
                                            fm,
                                            num_state,
                                            diagnostics,
                                            source,
                                            source,
                                            None,
                                        )?;
                                    if let Some(registry) = sources {
                                        registry.bind_text_boxes(&mut block, &mut paint, source)?;
                                    }
                                    vec![(block, direction, paint)]
                                }
                                MainStoryLayoutItem::Table(table, path) => {
                                    let (table, semantics) =
                                        crate::table::layout_table_with_provenance(
                                            table,
                                            note_width,
                                            styles,
                                            input,
                                            media,
                                            fm,
                                            num_state,
                                            diagnostics,
                                            sources,
                                            &story,
                                            &path,
                                            None,
                                        )?;
                                    crate::paginator::story_table_rows(
                                        &table,
                                        &semantics,
                                        &input
                                            .document
                                            .body
                                            .sect_pr
                                            .as_ref()
                                            .map(sect_pr_to_geometry)
                                            .unwrap_or_default(),
                                        media.media(),
                                    )?
                                    .into_iter()
                                    .map(|(block, paint)| (block, TextDirection::Auto, paint))
                                    .collect()
                                }
                            };
                            for (block, direction, paint) in laid_out_blocks {
                                let first = lines.len();
                                if let Some(color) = block.change_bar
                                    && !block.lines.is_empty()
                                {
                                    revision_ranges.push((first..first + block.lines.len(), color));
                                }
                                lines.extend(block.lines.iter().cloned());
                                let last = lines.len();
                                render_paragraphs.push(NoteRenderParagraph {
                                    paint: std::sync::Arc::new(paint),
                                    border_gap_before: 0.0,
                                    block,
                                    direction,
                                    lines: first..last,
                                });
                            }
                        }

                        // Group normalization happens after the complete story is laid out,
                        // leaving cached/source paragraphs and ordinary note paths intact.
                        let groups = crate::block::paragraph_border_groups(
                            &render_paragraphs
                                .iter()
                                .map(|paragraph| Some(&paragraph.block))
                                .collect::<Vec<_>>(),
                        );
                        for index in 0..render_paragraphs.len() {
                            let (top, bottom) = groups[index];
                            let gap = if top {
                                let previous = &render_paragraphs[index - 1].block;
                                let current = &render_paragraphs[index].block;
                                if input.do_not_use_html_paragraph_auto_spacing {
                                    previous.space_after + current.space_before
                                } else {
                                    previous.space_after.max(current.space_before)
                                }
                            } else {
                                0.0
                            };
                            let paragraph = &mut render_paragraphs[index];
                            paragraph.block.remove_group_border_edges(top, bottom);
                            paragraph.border_gap_before = gap;
                            if let Some(first) = paragraph.block.lines.first_mut() {
                                first.height += gap;
                                first.ascent += gap;
                            }
                            lines[paragraph.lines.clone()].clone_from_slice(&paragraph.block.lines);
                        }

                        let marker_text = if special.is_some() {
                            " ".to_owned()
                        } else {
                            num_state.note_label(note_ref)
                        };
                        let Some(marker) = shape_marker(&marker_text, fm)? else {
                            continue;
                        };

                        let entry = NoteEntry {
                            layout: NoteLayout {
                                marker,
                                draw_marker: false,
                                marker_rise: NOTE_FONT_SIZE * 0.33,
                                lines,
                                revision_ranges,
                            },
                            paragraphs: render_paragraphs,
                        };
                        if let Some(record) = special {
                            specials.insert((kind, record, content_width.to_bits()), entry);
                        } else {
                            notes.insert(key, entry);
                        }
                    }
                }
            }

            Ok(NoteRegistry {
                notes,
                specials,
                continuation_separator,
            })
        })();
        // The numbering state may be reused after an error or another width.
        num_state.note_body_context = previous_context;
        result
    }

    /// The note as broken for a section of this content width.
    pub fn get(&self, note: NoteRef, content_width: f64) -> Option<&NoteLayout> {
        self.notes
            .get(&(note, content_width.to_bits()))
            .map(|entry| &entry.layout)
    }

    pub(crate) fn get_render(
        &self,
        note: NoteRef,
        content_width: f64,
    ) -> Option<(&NoteLayout, &[NoteRenderParagraph])> {
        self.notes
            .get(&(note, content_width.to_bits()))
            .map(|entry| (&entry.layout, entry.paragraphs.as_slice()))
    }

    /// A special record laid out at this section's width.
    pub(crate) fn get_special(
        &self,
        stream: NoteStream,
        record: NoteType,
        content_width: f64,
    ) -> Option<(&NoteLayout, &[NoteRenderParagraph])> {
        self.specials
            .get(&(stream, record, content_width.to_bits()))
            .map(|entry| (&entry.layout, entry.paragraphs.as_slice()))
    }

    /// Whether the stream defines an authored continuation separator.
    pub fn has_continuation_separator(&self, stream: NoteStream) -> bool {
        self.continuation_separator[usize::from(stream == NoteStream::Endnote)]
    }
}

/// Read only actual direct Word note marks, under the existing writer's
/// namespace publication authority. Raw positions retain their original run
/// boundary. A one-child projection keeps producer shadows and root bindings
/// without assigning Word meaning to a lexical prefix or opaque descendants.
pub(crate) fn run_note_events(run: &rdocx_oxml::text::CT_R) -> Result<Vec<(usize, String)>> {
    use quick_xml::{Writer, events::Event, name::ResolveResult, reader::NsReader};
    use rdocx_oxml::text::CT_R;
    let mut events = Vec::new();
    for (index, raw) in run.extra_xml.iter().enumerate() {
        let encoded = run.extra_xml_positions.get(index).copied().unwrap_or(0);
        if CT_R::raw_child_is_root_attributes(encoded) {
            continue;
        }
        let mut view = CT_R::new("");
        view.properties = None;
        view.content.clear();
        view.alt_drawings.clear();
        view.extra_xml.clear();
        view.extra_xml_positions.clear();
        for (position, root) in run.extra_xml_positions.iter().zip(&run.extra_xml) {
            if CT_R::raw_child_is_root_attributes(*position) {
                view.extra_xml.push(root.clone());
                view.extra_xml_positions.push(*position);
            }
        }
        view.extra_xml.push(raw.clone());
        view.extra_xml_positions.push(0);
        let mut writer = Writer::new(Vec::new());
        view.to_xml(&mut writer)
            .map_err(|error| oxml_layout::LayoutError::Layout(error.to_string()))?;
        let mut xml = br#"<rdocx-note-scope xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#.to_vec();
        xml.extend(writer.into_inner());
        xml.extend_from_slice(b"</rdocx-note-scope>");
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut depth = 0;
        loop {
            let event = reader
                .read_event()
                .map_err(|error| oxml_layout::LayoutError::Layout(error.to_string()))?;
            if let Event::Start(element) | Event::Empty(element) = &event
                && depth == 2
                && matches!(reader.resolver().resolve_element(element.name()).0,ResolveResult::Bound(namespace) if namespace.as_ref()==rdocx_oxml::namespace::W_NS.as_bytes())
            {
                let local = element.local_name();
                if matches!(
                    local.as_ref(),
                    b"footnoteRef" | b"endnoteRef" | b"separator" | b"continuationSeparator"
                ) {
                    events.push((
                        CT_R::raw_child_position(encoded)
                            .saturating_sub(usize::from(run.properties.is_some()))
                            .min(run.content.len()),
                        String::from_utf8_lossy(local.as_ref()).into_owned(),
                    ));
                }
            }
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth -= 1,
                Event::Eof => break,
                _ => {}
            }
        }
    }
    events.sort_by_key(|(position, _)| *position);
    Ok(events)
}

/// Shape a note's number as the superscript marker drawn beside it.
fn shape_marker(text: &str, fm: &mut FontManager) -> Result<Option<TextSegment>> {
    let text = text.to_owned();
    let size = NOTE_FONT_SIZE * 0.58;

    let Ok(font_id) = fm.resolve_font(Some("serif"), false, false) else {
        return Ok(None);
    };
    let Ok(shaped) = fm.shape_text(font_id, &text, size) else {
        return Ok(None);
    };
    let metrics = fm.metrics(font_id, size)?;

    Ok(Some(TextSegment {
        text,
        direction: oxml_layout::TextDirection::Auto,
        source: None,
        font_id,
        font_size: size,
        glyph_ids: shaped.glyph_ids,
        advances: shaped.advances,
        width: shaped.width,
        ascent: metrics.ascent,
        descent: metrics.descent,
        line_gap: 0.0,
        color: Color::BLACK,
        bold: false,
        italic: false,
        underline: None,
        strike: false,
        dstrike: false,
        highlight: None,
        baseline_offset: 0.0,
        hyperlink_url: None,
        field_kind: None,
        field_source: None,
        note: None,
        note_reference_source: None,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
    use rdocx_oxml::text::CT_P;

    /// One footnote, numbered 1, whose text is long enough to wrap at any of
    /// the widths under test.
    fn input_with_one_note() -> LayoutInput {
        let mut note = CT_P::new();
        note.add_run(
            "A note long enough that the measure it is broken to decides how \
             many lines it occupies rather than leaving it on a single line.",
        );

        LayoutInput {
            sequence_snapshot: None,
            revision_view: crate::input::RevisionView::Accepted,
            revision_style: crate::input::RevisionStyle::Plain,
            automatic_hyphenation: false,
            mirror_margins: false,
            gutter_at_top: false,
            do_not_use_html_paragraph_auto_spacing: false,
            default_tab_stop: None,
            clamp_tabs_past_margin: false,
            legacy_table_positioning: false,
            modern_footnote_layout: false,
            footnote_layout_like_word8: false,
            math_properties: None,
            note_defaults: [None, None],
            document: rdocx_oxml::document::CT_Document::new(),
            styles: CT_Styles::new_default(),
            numbering: None,
            story_part_names: Default::default(),
            story_bodies: Default::default(),
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: Some(CT_Footnotes {
                footnotes: vec![CT_Footnote {
                    id: 1,
                    note_type: NoteType::Normal,
                    paragraphs: vec![note],
                }],
            }),
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
            embedded_fonts: Vec::new(),
        }
    }

    #[test]
    fn source_note_events_keep_namespace_and_actual_run_boundaries() {
        use rdocx_oxml::text::CT_R;
        for (xml, expected) in [
            (
                r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:b/></w:rPr><w:t>before</w:t><w:footnoteRef/><w:t>after</w:t><w:footnoteRef></w:footnoteRef></w:r>"#,
                vec![(1, "footnoteRef"), (2, "footnoteRef")],
            ),
            (
                r#"<q:r xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="urn:producer"><q:footnoteRef/><w:footnoteRef/><q:t>after</q:t></q:r>"#,
                vec![(0, "footnoteRef")],
            ),
            (
                r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><v:opaque xmlns:v="urn:producer"><w:footnoteRef/></v:opaque><v:footnoteRef xmlns:v="urn:producer"/></w:r>"#,
                vec![],
            ),
            (
                r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:separator/><w:continuationSeparator/><w:endnoteRef/></w:r>"#,
                vec![
                    (0, "separator"),
                    (0, "continuationSeparator"),
                    (0, "endnoteRef"),
                ],
            ),
        ] {
            let run = CT_R::from_xml_fragment_checked(xml.as_bytes()).unwrap();
            let before = run.clone();
            assert_eq!(
                run_note_events(&run).unwrap(),
                expected
                    .into_iter()
                    .map(|(n, s)| (n, s.into()))
                    .collect::<Vec<_>>(),
                "{xml}"
            );
            assert_eq!(run, before);
        }
    }

    #[test]
    fn note_source_context_is_restored_on_success_error_and_repeated_widths() {
        let mut input = input_with_one_note();
        let media = MediaRegistry::new(&HashMap::new());
        let mut fm = FontManager::new_deterministic().unwrap();
        let mut state = NumberingState::new();
        let sentinel = Some(NoteRef {
            stream: NoteStream::Endnote,
            id: 99,
        });
        state.note_body_context = sentinel;
        NoteRegistry::build(
            &input,
            &input.styles,
            &media,
            &mut fm,
            &mut state,
            &[468.0, 400.0, 468.0],
            &mut Vec::new(),
            None,
        )
        .unwrap();
        assert_eq!(state.note_body_context, sentinel);
        let run = &mut input.footnotes.as_mut().unwrap().footnotes[0].paragraphs[0].runs[0];
        run.extra_xml.push(b"<w:footnoteRef".to_vec());
        run.extra_xml_positions.push(0);
        assert!(
            NoteRegistry::build(
                &input,
                &input.styles,
                &media,
                &mut fm,
                &mut state,
                &[468.0],
                &mut Vec::new(),
                None
            )
            .is_err()
        );
        assert_eq!(state.note_body_context, sentinel);
    }

    fn build_at(widths: &[f64]) -> NoteRegistry {
        let input = input_with_one_note();
        let media = MediaRegistry::new(&HashMap::new());
        let mut fm = FontManager::new();
        let mut num_state = NumberingState::new();
        let mut diagnostics = Vec::new();
        NoteRegistry::build(
            &input,
            &input.styles,
            &media,
            &mut fm,
            &mut num_state,
            widths,
            &mut diagnostics,
            None,
        )
        .expect("the registry builds")
    }

    const NOTE_ONE: NoteRef = NoteRef {
        stream: NoteStream::Footnote,
        id: 1,
    };

    #[test]
    fn the_registry_lays_a_note_out_once_per_distinct_width() {
        let registry = build_at(&[468.0, 1044.0]);

        let narrow = registry.get(NOTE_ONE, 468.0).expect("narrow is registered");
        let wide = registry.get(NOTE_ONE, 1044.0).expect("wide is registered");

        assert!(
            wide.lines.len() < narrow.lines.len(),
            "one layout was reused for both widths, {} lines against {}",
            wide.lines.len(),
            narrow.lines.len()
        );
    }

    #[test]
    fn a_repeated_width_is_registered_once_and_still_found() {
        let repeated = build_at(&[468.0, 468.0]);
        let once = build_at(&[468.0]);

        let from_repeated = repeated.get(NOTE_ONE, 468.0).expect("still registered");
        let from_once = once.get(NOTE_ONE, 468.0).expect("registered");
        assert_eq!(from_repeated.lines.len(), from_once.lines.len());
    }

    #[test]
    fn an_unregistered_width_has_no_layout() {
        // The engine registers every width it paginates, so a miss means the
        // caller and the builder disagree, and silently drawing the wrong
        // measure would be worse than drawing nothing.
        let registry = build_at(&[468.0]);
        assert!(registry.get(NOTE_ONE, 1044.0).is_none());
    }
}
