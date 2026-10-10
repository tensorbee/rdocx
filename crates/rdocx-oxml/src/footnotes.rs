//! Footnote and endnote elements: `CT_Footnotes`, `CT_Footnote`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::word_prefixes_at;
use crate::properties::is_word_element;
use crate::text::{CT_P, ROOT_R_BINDING, declare_w14_on_part_root, root_binding_scope};

/// `ST_FtnEdn` — what a note in the stream is for.
///
/// The stream holds the document's real notes alongside the separator marks
/// Word draws above them. The distinction is carried by `w:type`, not by the
/// id: the conventional ids 0 and 1 are a convention, not a guarantee, and
/// reading a `continuationSeparator` as if it were note number 1 is how a
/// separator ends up rendered as body content.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum NoteType {
    /// A real note, the only kind a reference can resolve to.
    #[default]
    Normal,
    /// The rule drawn above the notes on a page.
    Separator,
    /// The rule drawn above a note carried over from the previous page.
    ContinuationSeparator,
    /// The notice Word can place when a note continues.
    ContinuationNotice,
}

impl NoteType {
    pub(crate) fn from_str(s: &str) -> Self {
        match s {
            "separator" => NoteType::Separator,
            "continuationSeparator" => NoteType::ContinuationSeparator,
            "continuationNotice" => NoteType::ContinuationNotice,
            _ => NoteType::Normal,
        }
    }

    fn to_str(self) -> Option<&'static str> {
        match self {
            NoteType::Normal => None,
            NoteType::Separator => Some("separator"),
            NoteType::ContinuationSeparator => Some("continuationSeparator"),
            NoteType::ContinuationNotice => Some("continuationNotice"),
        }
    }
}

/// A single footnote or endnote.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Footnote {
    /// Footnote ID (matches w:footnoteReference w:id in the document body).
    pub id: i32,
    /// What this entry is for. Separators are retained so a round trip does
    /// not discard them, and filtered out of `get_by_id` so no reference can
    /// resolve to one.
    pub note_type: NoteType,
    /// Paragraphs making up the footnote content.
    pub paragraphs: Vec<CT_P>,
}

/// Collection of footnotes parsed from `word/footnotes.xml`.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Footnotes {
    pub footnotes: Vec<CT_Footnote>,
}

#[allow(non_snake_case)]
impl CT_Footnotes {
    pub fn new() -> Self {
        CT_Footnotes {
            footnotes: Vec::new(),
        }
    }

    /// Get a real note by its ID.
    ///
    /// Separator entries are never returned. A reference can only ever mean a
    /// real note, and some documents place a `continuationSeparator` at an id
    /// a reference could otherwise collide with.
    pub fn get_by_id(&self, id: i32) -> Option<&CT_Footnote> {
        self.footnotes
            .iter()
            .find(|f| f.id == id && f.note_type == NoteType::Normal)
    }

    /// Whether the stream defines the rule drawn above a carried-over note.
    pub fn has_continuation_separator(&self) -> bool {
        self.footnotes
            .iter()
            .any(|f| f.note_type == NoteType::ContinuationSeparator)
    }

    /// Parse from XML bytes (the content of footnotes.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        crate::content_control::validate_story_part_nesting(xml)?;
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut footnotes = Vec::new();
        let mut buf = Vec::new();
        let mut word_prefixes = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"footnote", &prefixes)
                        || is_word_element(name.as_ref(), b"endnote", &prefixes)
                    {
                        let mut id: i32 = 0;
                        let mut note_type = NoteType::Normal;
                        for attr in e.attributes().flatten() {
                            if matches_local_name(attr.key.as_ref(), b"id") {
                                id = std::str::from_utf8(&attr.value)
                                    .unwrap_or("0")
                                    .parse()
                                    .unwrap_or(0);
                            } else if matches_local_name(attr.key.as_ref(), b"type") {
                                note_type = NoteType::from_str(
                                    std::str::from_utf8(&attr.value).unwrap_or(""),
                                );
                            }
                        }

                        // `w:type` decides what an entry is, because the ids
                        // separators conventionally use are a convention
                        // rather than a rule, and a `continuationSeparator`
                        // sitting at id 1 must not read as note number one.
                        //
                        // An untyped entry at id 0 or below is still treated
                        // as a separator. That is the older convention, and
                        // producers that predate writing `w:type` rely on it.
                        if note_type == NoteType::Normal && id <= 0 {
                            note_type = NoteType::Separator;
                        }

                        // Separators are kept rather than dropped, so a round
                        // trip preserves them.
                        let paragraphs = parse_footnote_content(&mut reader, &prefixes)?;
                        footnotes.push(CT_Footnote {
                            id,
                            note_type,
                            paragraphs,
                        });
                    } else if is_word_element(name.as_ref(), b"footnotes", &prefixes)
                        || is_word_element(name.as_ref(), b"endnotes", &prefixes)
                    {
                        // Continue into the root element
                        word_prefixes = prefixes;
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Footnotes { footnotes })
    }

    /// Serialize to XML bytes as footnotes.
    pub fn to_xml_footnotes(&self) -> Result<Vec<u8>> {
        self.to_xml_root("w:footnotes", "w:footnote")
    }

    /// Serialize to XML bytes as endnotes.
    pub fn to_xml_endnotes(&self) -> Result<Vec<u8>> {
        self.to_xml_root("w:endnotes", "w:endnote")
    }

    fn to_xml_root(&self, root_tag: &str, item_tag: &str) -> Result<Vec<u8>> {
        let _binding_scope = root_binding_scope(ROOT_R_BINDING);
        let mut writer = Writer::new(Vec::new());

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut start = BytesStart::new(root_tag);
        start.push_attribute(("xmlns:w", W_NS));
        start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(start))?;

        let mut buf = itoa::Buffer::new();
        for footnote in &self.footnotes {
            let mut fn_start = BytesStart::new(item_tag);
            // `w:type` precedes `w:id` in the schema's attribute listing, and
            // is omitted entirely for a normal note, which is what Word writes.
            if let Some(type_str) = footnote.note_type.to_str() {
                fn_start.push_attribute(("w:type", type_str));
            }
            fn_start.push_attribute(("w:id", buf.format(footnote.id)));
            writer.write_event(Event::Start(fn_start))?;

            for p in &footnote.paragraphs {
                p.to_xml(&mut writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new(item_tag)))?;
        }

        writer.write_event(Event::End(BytesEnd::new(root_tag)))?;

        let mut xml = writer.into_inner();
        declare_w14_on_part_root(&mut xml)?;
        Ok(xml)
    }
}

impl Default for CT_Footnotes {
    fn default() -> Self {
        Self::new()
    }
}

/// Parse the content of a single footnote/endnote (paragraphs until closing tag).
fn parse_footnote_content(
    reader: &mut Reader<&[u8]>,
    word_prefixes: &[String],
) -> Result<Vec<CT_P>> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let prefixes = word_prefixes_at(e, word_prefixes)?;
                if is_word_element(name.as_ref(), b"p", &prefixes) {
                    paragraphs.push(CT_P::from_xml_with_prefixes_and_root(
                        reader,
                        &prefixes,
                        Some(e),
                    )?);
                } else {
                    reader.read_to_end_into(name, &mut Vec::new())?;
                }
            }
            Ok(Event::Empty(ref e)) => {
                let prefixes = word_prefixes_at(e, word_prefixes)?;
                if is_word_element(e.name().as_ref(), b"p", &prefixes) {
                    paragraphs.push(CT_P::from_empty_root(e, &prefixes)?);
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                if matches_local_name(name.as_ref(), b"footnote")
                    || matches_local_name(name.as_ref(), b"endnote")
                {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(paragraphs)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rewritten_notes_are_compact_and_declare_w_once() {
        let xml = br#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p w:rsidR="00A1B2C3"><w:r w:rsidRPr="00A1B2C4"><w:t>note</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
        let footnotes = CT_Footnotes::from_xml(xml).unwrap();
        let output = String::from_utf8(footnotes.to_xml_footnotes().unwrap()).unwrap();
        assert_eq!(output.matches("xmlns:w=").count(), 1, "{output}");
        assert!(!output.contains('\n'), "{output}");
        assert!(
            output.contains(r#"<w:footnote w:id="1"><w:p w:rsidR="00A1B2C3"><w:r w:rsidRPr="00A1B2C4"><w:t>note</w:t></w:r></w:p></w:footnote>"#),
            "{output}"
        );
    }

    #[test]
    fn a_note_paragraph_identity_stays_bound_under_the_written_root() {
        // The written root declares only `w` and `r`, while Word declares
        // `w14` on the root of the notes part it writes.
        let xml = br#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:footnote w:id="1"><w:p w14:paraId="1A2B3C4D" w14:textId="5E6F7A8B"><w:r><w:t>note</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
        let footnotes = CT_Footnotes::from_xml(xml).unwrap();
        let output = footnotes.to_xml_footnotes().unwrap();
        let reparsed = CT_Footnotes::from_xml(&output)
            .unwrap_or_else(|error| panic!("{error}: {}", String::from_utf8_lossy(&output)));
        assert_eq!(reparsed.to_xml_footnotes().unwrap(), output);
    }

    #[test]
    fn parse_footnotes_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:footnote w:id="0">
                <w:p><w:r><w:t>separator</w:t></w:r></w:p>
            </w:footnote>
            <w:footnote w:id="1">
                <w:p><w:r><w:t>First footnote text.</w:t></w:r></w:p>
            </w:footnote>
            <w:footnote w:id="2">
                <w:p><w:r><w:t>Second footnote.</w:t></w:r></w:p>
                <w:p><w:r><w:t>With two paragraphs.</w:t></w:r></w:p>
            </w:footnote>
        </w:footnotes>"#;

        let footnotes = CT_Footnotes::from_xml(xml).unwrap();
        // The untyped id=0 entry is retained as a separator, so a round trip
        // preserves it, but it is not reachable as a note.
        assert_eq!(footnotes.footnotes.len(), 3);
        assert_eq!(footnotes.footnotes[0].note_type, NoteType::Separator);
        assert!(footnotes.get_by_id(0).is_none());

        let first = footnotes.get_by_id(1).unwrap();
        assert_eq!(first.paragraphs.len(), 1);
        assert_eq!(first.paragraphs[0].text(), "First footnote text.");
        assert_eq!(footnotes.get_by_id(2).unwrap().paragraphs.len(), 2);
    }

    #[test]
    fn parse_endnotes_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:endnote w:id="0">
                <w:p><w:r><w:t>separator</w:t></w:r></w:p>
            </w:endnote>
            <w:endnote w:id="1">
                <w:p><w:r><w:t>An endnote.</w:t></w:r></w:p>
            </w:endnote>
        </w:endnotes>"#;

        let endnotes = CT_Footnotes::from_xml(xml).unwrap();
        assert_eq!(endnotes.footnotes.len(), 2);
        assert_eq!(endnotes.footnotes[0].note_type, NoteType::Separator);
        assert_eq!(
            endnotes.get_by_id(1).unwrap().paragraphs[0].text(),
            "An endnote."
        );
    }

    #[test]
    fn aliased_footnote_paragraph_properties_keep_root_scope() {
        let xml = format!(
            r#"<q:footnotes xmlns:q="{W_NS}" xmlns:ext="urn:producer"><ext:footnote ext:id="9"><ext:p/></ext:footnote><q:footnote q:id="1"><ext:p><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:p><q:p><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr><q:r><q:t>Note</q:t></q:r></q:p></q:footnote></q:footnotes>"#
        );
        let footnotes = CT_Footnotes::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(footnotes.footnotes.len(), 1);
        assert_eq!(footnotes.footnotes[0].paragraphs.len(), 1);
        let paragraph = &footnotes.footnotes[0].paragraphs[0];
        assert_eq!(paragraph.text(), "Note");
        assert_eq!(
            paragraph.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn default_namespace_footnote_properties_keep_root_scope() {
        let xml = format!(
            r#"<footnotes xmlns="{W_NS}" xmlns:w="{W_NS}" xmlns:ext="urn:producer"><ext:footnote ext:id="9"><ext:p/></ext:footnote><footnote w:id="1"><ext:p><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:p><p><pPr><ext:jc ext:val="right"/><jc w:val="center"/></pPr><r><t>Note</t></r></p></footnote></footnotes>"#
        );
        let footnotes = CT_Footnotes::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(footnotes.footnotes.len(), 1);
        assert_eq!(footnotes.footnotes[0].paragraphs.len(), 1);
        let paragraph = &footnotes.footnotes[0].paragraphs[0];
        assert_eq!(paragraph.text(), "Note");
        assert_eq!(
            paragraph.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn get_footnote_by_id() {
        let footnotes = CT_Footnotes {
            footnotes: vec![
                CT_Footnote {
                    id: 1,
                    note_type: NoteType::Normal,
                    paragraphs: vec![],
                },
                CT_Footnote {
                    id: 2,
                    note_type: NoteType::Normal,
                    paragraphs: vec![],
                },
            ],
        };
        assert!(footnotes.get_by_id(1).is_some());
        assert!(footnotes.get_by_id(2).is_some());
        assert!(footnotes.get_by_id(3).is_none());
    }

    #[test]
    fn round_trip_footnotes() {
        let mut fn1_para = CT_P::new();
        fn1_para.add_run("First footnote.");

        let mut fn2_para = CT_P::new();
        fn2_para.add_run("Second footnote.");

        let footnotes = CT_Footnotes {
            footnotes: vec![
                CT_Footnote {
                    id: 1,
                    note_type: NoteType::Normal,
                    paragraphs: vec![fn1_para],
                },
                CT_Footnote {
                    id: 2,
                    note_type: NoteType::Normal,
                    paragraphs: vec![fn2_para],
                },
            ],
        };

        let xml = footnotes.to_xml_footnotes().unwrap();
        let parsed = CT_Footnotes::from_xml(&xml).unwrap();
        assert_eq!(parsed.footnotes.len(), 2);
        assert_eq!(parsed.footnotes[0].id, 1);
        assert_eq!(parsed.footnotes[0].paragraphs[0].text(), "First footnote.");
        assert_eq!(parsed.footnotes[1].id, 2);
        assert_eq!(parsed.footnotes[1].paragraphs[0].text(), "Second footnote.");
    }

    #[test]
    fn footnote_round_trip_keeps_an_inherited_raw_relationship_attribute_bound() {
        let xml = format!(
            r#"<w:footnotes xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:footnote w:id="1"><w:p><x:raw xmlns:x="urn:producer" r:id="rId9"/></w:p></w:footnote></w:footnotes>"#
        );
        let parsed = CT_Footnotes::from_xml(xml.as_bytes()).unwrap();
        let serialized = parsed.to_xml_footnotes().unwrap();
        let serialized = std::str::from_utf8(&serialized).unwrap();
        assert!(
            serialized.contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "{serialized}"
        );
        assert!(serialized.contains(r#"r:id="rId9""#), "{serialized}");
        CT_Footnotes::from_xml(serialized.as_bytes()).unwrap();
    }

    #[test]
    fn endnote_round_trip_keeps_an_inherited_raw_relationship_attribute_bound() {
        let xml = format!(
            r#"<w:endnotes xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:endnote w:id="1"><w:p><x:raw xmlns:x="urn:producer" r:id="rId8"/></w:p></w:endnote></w:endnotes>"#
        );
        let parsed = CT_Footnotes::from_xml(xml.as_bytes()).unwrap();
        let serialized = parsed.to_xml_endnotes().unwrap();
        let serialized = std::str::from_utf8(&serialized).unwrap();
        assert!(
            serialized.contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "{serialized}"
        );
        assert!(serialized.contains(r#"r:id="rId8""#), "{serialized}");
        CT_Footnotes::from_xml(serialized.as_bytes()).unwrap();
    }

    // F-X013b, note types and separator preservation.

    const WITH_SEPARATORS: &str = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>A real note.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;

    #[test]
    fn a_separator_definition_survives_open_and_save() {
        let parsed = CT_Footnotes::from_xml(WITH_SEPARATORS.as_bytes()).unwrap();
        assert_eq!(parsed.footnotes.len(), 3, "separators must be retained");

        let round_tripped = CT_Footnotes::from_xml(&parsed.to_xml_footnotes().unwrap()).unwrap();
        assert_eq!(round_tripped.footnotes.len(), 3);
        assert_eq!(round_tripped.footnotes[0].note_type, NoteType::Separator);
        assert_eq!(round_tripped.footnotes[0].id, 0);
        assert_eq!(
            round_tripped.footnotes[1].note_type,
            NoteType::ContinuationSeparator
        );
        assert_eq!(round_tripped.footnotes[1].id, 1);
        assert_eq!(round_tripped.footnotes[2].note_type, NoteType::Normal);
        assert_eq!(round_tripped, parsed, "a second trip must be a fixed point");
    }

    #[test]
    fn get_by_id_does_not_return_a_separator() {
        let parsed = CT_Footnotes::from_xml(WITH_SEPARATORS.as_bytes()).unwrap();

        // Id 1 is the continuation separator here, not note number one.
        assert!(parsed.get_by_id(0).is_none(), "separator is not a note");
        assert!(
            parsed.get_by_id(1).is_none(),
            "continuation separator is not a note"
        );
        assert_eq!(
            parsed.get_by_id(2).unwrap().paragraphs[0].text(),
            "A real note."
        );
        assert!(parsed.has_continuation_separator());
    }

    #[test]
    fn note_types_are_read_through_a_foreign_prefix() {
        let xml = r#"<?xml version="1.0"?>
<x:footnotes xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <x:footnote x:type="separator" x:id="0"><x:p/></x:footnote>
  <x:footnote x:id="7"><x:p><x:r><x:t>Prefixed.</x:t></x:r></x:p></x:footnote>
</x:footnotes>"#;
        let parsed = CT_Footnotes::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(parsed.footnotes.len(), 2);
        assert_eq!(parsed.footnotes[0].note_type, NoteType::Separator);
        assert!(parsed.get_by_id(0).is_none());
        assert_eq!(
            parsed.get_by_id(7).unwrap().paragraphs[0].text(),
            "Prefixed."
        );
    }

    #[test]
    fn an_unknown_note_type_reads_as_a_normal_note() {
        let xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="somethingNew" w:id="3"><w:p><w:r><w:t>Note.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
        let parsed = CT_Footnotes::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(parsed.footnotes[0].note_type, NoteType::Normal);
        assert!(parsed.get_by_id(3).is_some());
    }
}
