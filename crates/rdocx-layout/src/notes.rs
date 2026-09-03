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

use rdocx_oxml::styles::CT_Styles;

use crate::WordStory;
use crate::block::ParagraphBlock;
use crate::engine::{SourceRegistry, layout_paragraph_with_source_and_direction};
use crate::input::{LayoutInput, MediaRegistry};
use crate::style_resolver::NumberingState;
use oxml_layout::{
    Color, Diagnostic, FontManager, LayoutLine, NoteRef, NoteStream, Result, TextDirection,
    TextSegment,
};

/// Point size notes are set at.
const NOTE_FONT_SIZE: f64 = 8.0;
/// Horizontal space reserved for the marker, to the left of note text.
///
/// Notes are both line-broken and drawn against this, so the two agree.
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
    /// How far above the baseline the marker sits.
    pub marker_rise: f64,
    /// The note's lines, flattened across its paragraphs.
    pub lines: Vec<LayoutLine>,
    /// Line ranges belonging to paragraphs with a visible tracked revision.
    pub revision_ranges: Vec<std::ops::Range<usize>>,
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
    continuation_separator: bool,
}

#[derive(Debug, Clone)]
struct NoteEntry {
    layout: NoteLayout,
    paragraphs: Vec<NoteRenderParagraph>,
}

#[derive(Debug, Clone)]
pub(crate) struct NoteRenderParagraph {
    pub block: ParagraphBlock,
    pub direction: TextDirection,
    pub lines: std::ops::Range<usize>,
}

impl NoteRegistry {
    /// Lay out every note in the footnote and endnote streams, once for each
    /// distinct content width the document paginates at.
    ///
    /// A note is broken at `content_width - NOTE_INDENT`, because that is where
    /// it is drawn, and the width that matters is the one belonging to the
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
        let mut notes = HashMap::new();
        let mut continuation_separator = false;

        // Each stream is keyed separately, so a document numbering a footnote
        // and an endnote alike keeps both.
        for (kind, stream) in [
            (NoteStream::Footnote, input.footnotes.as_ref()),
            (NoteStream::Endnote, input.endnotes.as_ref()),
        ]
        .into_iter()
        .filter_map(|(kind, stream)| stream.map(|stream| (kind, stream)))
        {
            if stream.has_continuation_separator() {
                continuation_separator = true;
            }

            for note in &stream.footnotes {
                // `get_by_id` is the authority on what counts as a real note,
                // so separators never reach the registry.
                if stream.get_by_id(note.id).is_none() {
                    continue;
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
                    if notes.contains_key(&key) {
                        continue;
                    }
                    if laid_out {
                        *num_state = counters_before.clone();
                    }
                    laid_out = true;
                    let note_width = (content_width - NOTE_INDENT).max(1.0);

                    let mut lines = Vec::new();
                    let mut revision_ranges = Vec::new();
                    let mut render_paragraphs = Vec::new();
                    let story = match kind {
                        NoteStream::Footnote => WordStory::Footnote { id: note.id },
                        NoteStream::Endnote => WordStory::Endnote { id: note.id },
                    };
                    for (paragraph_index, paragraph) in note.paragraphs.iter().enumerate() {
                        let source =
                            sources.and_then(|sources| sources.id(&story, &[paragraph_index]));
                        let (block, direction) = layout_paragraph_with_source_and_direction(
                            paragraph,
                            note_width,
                            styles,
                            input,
                            media,
                            fm,
                            num_state,
                            diagnostics,
                            source,
                        )?;
                        let first = lines.len();
                        if block.has_visible_revision && !block.lines.is_empty() {
                            revision_ranges.push(first..first + block.lines.len());
                        }
                        lines.extend(block.lines.iter().cloned());
                        let last = lines.len();
                        render_paragraphs.push(NoteRenderParagraph {
                            block,
                            direction,
                            lines: first..last,
                        });
                    }

                    let Some(marker) = shape_marker(note.id, fm)? else {
                        continue;
                    };

                    notes.insert(
                        key,
                        NoteEntry {
                            layout: NoteLayout {
                                marker,
                                marker_rise: NOTE_FONT_SIZE * 0.33,
                                lines,
                                revision_ranges,
                            },
                            paragraphs: render_paragraphs,
                        },
                    );
                }
            }
        }

        Ok(NoteRegistry {
            notes,
            continuation_separator,
        })
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

    /// Whether either stream defined the rule drawn above a carried note.
    pub fn has_continuation_separator(&self) -> bool {
        self.continuation_separator
    }
}

/// Shape a note's number as the superscript marker drawn beside it.
fn shape_marker(id: i32, fm: &mut FontManager) -> Result<Option<TextSegment>> {
    let text = id.to_string();
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
        note: None,
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
            revision_view: crate::input::RevisionView::Accepted,
            automatic_hyphenation: false,
            math_properties: None,
            document: rdocx_oxml::document::CT_Document::new(),
            styles: CT_Styles::new_default(),
            numbering: None,
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
        }
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
