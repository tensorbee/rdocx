//! Whole-document text views and validation, shared by the CLI and the
//! bindings so that `rdocx text`, `rdocx convert --to md|html` and
//! `rdocx validate` give the same result as the Python methods.

use std::path::Path;

use oxml_opc::OpcPackage;
use oxml_opc::relationship::rel_types;
use quick_xml::XmlVersion;
use quick_xml::escape::escape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::namespace::{MC_NS, W_NS};

use crate::{Document, Result, StoryId, StoryItemKind, StoryKind};

/// The accepted-view text of one story that the body view leaves out.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryText {
    /// The story the items belong to.
    pub story: StoryId,
    /// Each direct paragraph and block content control of the story.
    pub items: Vec<StoryTextItem>,
}

/// One direct paragraph or block content control of a [`StoryText`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryTextItem {
    /// The index path of the item inside its story.
    pub index_path: Vec<usize>,
    /// [`StoryItemKind::Paragraph`] or [`StoryItemKind::ContentControl`].
    pub kind: StoryItemKind,
    /// The accepted-view text of the item.
    pub text: String,
}

/// A text, Markdown or HTML view of the body followed by every other story.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryTextExport {
    /// The complete view.
    pub text: String,
    /// Why the stories other than the body are left out, when a story part
    /// cannot be read. The view then holds the body only.
    pub omitted_stories: Option<String>,
}

/// The findings of `rdocx validate`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ValidationReport {
    /// Structural errors: a consumer may refuse or repair the file.
    pub errors: Vec<String>,
    /// Advisory findings, such as empty paragraphs or a missing title.
    pub warnings: Vec<String>,
}

impl ValidationReport {
    /// Whether no structural error was found. Warnings do not count.
    pub fn is_valid(&self) -> bool {
        self.errors.is_empty()
    }
}

/// The relationships from the main document to its other story parts.
const STORY_RELATIONSHIPS: [&str; 5] = [
    rel_types::HEADER,
    rel_types::FOOTER,
    rel_types::FOOTNOTES,
    rel_types::ENDNOTES,
    rel_types::COMMENTS,
];

impl Document {
    /// Read every story but the main body and its table cells, which the body
    /// view already covers, in [`Document::stories`] order.
    ///
    /// Only the direct items of a story are read, so the text of an inline
    /// control or a field is not repeated after its paragraph. A table is left
    /// out because each of its cells is a table-cell story of its own.
    ///
    /// A story part that cannot be read, such as a truncated header or a
    /// header relationship to a missing part, returns the reason, naming the
    /// first part that is not well-formed XML when there is one. `validate`
    /// reports the same part as an error.
    pub fn other_story_texts(&self) -> std::result::Result<Vec<StoryText>, String> {
        self.read_other_story_texts().map_err(|error| {
            self.malformed_story_part()
                .unwrap_or_else(|| error.to_string())
        })
    }

    fn read_other_story_texts(&self) -> Result<Vec<StoryText>> {
        let mut main_part = None;
        let mut stories: Vec<StoryText> = Vec::new();
        for item in self.story_item_snapshots()? {
            let story = item.location().story();
            if story.kind() == StoryKind::Body {
                main_part = Some(story.part_name().to_owned());
                continue;
            }
            if story.kind() == StoryKind::TableCell
                && main_part.as_deref() == Some(story.part_name())
            {
                continue;
            }
            if stories.last().is_none_or(|last| &last.story != story) {
                stories.push(StoryText {
                    story: story.clone(),
                    items: Vec::new(),
                });
            }
            let kind = item.location().item_kind();
            if !matches!(
                kind,
                StoryItemKind::Paragraph | StoryItemKind::ContentControl
            ) {
                continue;
            }
            if item.is_direct_child()
                && let Some(last) = stories.last_mut()
            {
                last.items.push(StoryTextItem {
                    index_path: item.location().index_path().to_vec(),
                    kind,
                    text: item.text().unwrap_or_default().to_owned(),
                });
            }
        }
        Ok(stories)
    }

    /// Name the first story part of the main document that is not
    /// well-formed XML, with the reason.
    fn malformed_story_part(&self) -> Option<String> {
        let package = &self.package;
        let doc_part = package.main_document_part()?;
        let rels = package.get_part_rels(&doc_part)?;
        rels.items
            .iter()
            .filter(|rel| {
                rel.target_mode.as_deref() != Some("External")
                    && STORY_RELATIONSHIPS.contains(&rel.rel_type.as_str())
            })
            .find_map(|rel| {
                let part_name = OpcPackage::resolve_rel_target(&doc_part, &rel.target);
                let detail = xml_style_references(package.get_part(&part_name)?).err()?;
                Some(format!("part {part_name} is not well-formed XML: {detail}"))
            })
    }

    /// The plain text of every story, as `rdocx text` prints it.
    ///
    /// The body comes first, as [`Document::text`] reads it. Each other
    /// story part with text follows under a `--- kind (part name) ---` line,
    /// one line per paragraph: text boxes, headers, footers, footnotes,
    /// endnotes, and comments.
    pub fn text_with_stories(&self) -> StoryTextExport {
        let mut text = self.text();
        let (stories, omitted_stories) = split_stories(self.other_story_texts());
        for (kind, part_name, texts) in story_parts(&stories) {
            text.push_str(&format!("--- {} ({part_name}) ---\n", kind.name()));
            for line in texts {
                text.push_str(line);
                text.push('\n');
            }
        }
        StoryTextExport {
            text,
            omitted_stories,
        }
    }

    /// The Markdown of the body followed by every other story but comments,
    /// as `rdocx convert --to md` writes it.
    ///
    /// After a `---` rule, each story part with text follows under a bold
    /// `**kind** (part name)` label, one paragraph per line of text.
    pub fn to_markdown_with_stories(&self) -> StoryTextExport {
        let mut text = self.to_markdown();
        let (stories, omitted_stories) = split_stories(self.other_story_texts());
        push_markdown_stories(&mut text, &stories);
        StoryTextExport {
            text,
            omitted_stories,
        }
    }

    /// The HTML document of the body followed by every other story but
    /// comments, as `rdocx convert --to html` writes it.
    ///
    /// After an `<hr>`, each story part with text is one `<section>` before
    /// the end of the HTML body.
    pub fn to_html_with_stories(&self) -> StoryTextExport {
        let mut text = self.to_html();
        let (stories, omitted_stories) = split_stories(self.other_story_texts());
        push_html_stories(&mut text, &stories);
        StoryTextExport {
            text,
            omitted_stories,
        }
    }

    /// Validate the package at `path` as `rdocx validate` does.
    ///
    /// An error is returned only when the file is not a readable package,
    /// or when the document does not open and no structural error explains
    /// why.
    pub fn validate_file<P: AsRef<Path>>(path: P) -> Result<ValidationReport> {
        validate_package(OpcPackage::open(path)?)
    }

    /// Validate a serialized package as `rdocx validate` does.
    pub fn validate_bytes(bytes: &[u8]) -> Result<ValidationReport> {
        validate_package(OpcPackage::from_reader(std::io::Cursor::new(bytes))?)
    }
}

fn split_stories(
    stories: std::result::Result<Vec<StoryText>, String>,
) -> (Vec<StoryText>, Option<String>) {
    match stories {
        Ok(stories) => (stories, None),
        Err(reason) => (Vec::new(), Some(reason)),
    }
}

/// Group the texts of the stories by package part, each part labelled with
/// the kind of its first story, such as a header part holding table cells.
/// A part without any text, such as an empty header variant, is left out.
fn story_parts(stories: &[StoryText]) -> Vec<(StoryKind, &str, Vec<&str>)> {
    let mut parts: Vec<(StoryKind, &str, Vec<&str>)> = Vec::new();
    for story in stories {
        let part_name = story.story.part_name();
        if parts.last().is_none_or(|(_, last, _)| *last != part_name) {
            parts.push((story.story.kind(), part_name, Vec::new()));
        }
        if let Some((_, _, texts)) = parts.last_mut() {
            texts.extend(story.items.iter().map(|item| item.text.as_str()));
        }
    }
    parts.retain(|(_, _, texts)| texts.iter().any(|text| !text.trim().is_empty()));
    parts
}

/// The story parts a converted document carries after its body. Comments are
/// review annotations rather than document content, so they are left out.
fn converted_story_parts(stories: &[StoryText]) -> Vec<(StoryKind, &str, Vec<&str>)> {
    story_parts(stories)
        .into_iter()
        .filter(|(kind, _, _)| *kind != StoryKind::Comment)
        .collect()
}

/// Append the stories after the Markdown body, each part under a bold label.
fn push_markdown_stories(markdown: &mut String, stories: &[StoryText]) {
    let parts = converted_story_parts(stories);
    if parts.is_empty() {
        return;
    }
    if !markdown.is_empty() {
        if !markdown.ends_with("\n\n") {
            markdown.push_str(if markdown.ends_with('\n') {
                "\n"
            } else {
                "\n\n"
            });
        }
        markdown.push_str("---\n\n");
    }
    for (kind, part_name, texts) in parts {
        markdown.push_str(&format!("**{}** ({part_name})\n\n", kind.name()));
        for text in texts.iter().map(|text| text.trim()) {
            if !text.is_empty() {
                markdown.push_str(text);
                markdown.push_str("\n\n");
            }
        }
    }
}

/// Insert the stories before the end of the HTML body, one section per part.
fn push_html_stories(html: &mut String, stories: &[StoryText]) {
    let parts = converted_story_parts(stories);
    if parts.is_empty() {
        return;
    }
    let mut sections = String::from("<hr>\n");
    for (kind, part_name, texts) in parts {
        sections.push_str(&format!(
            "<section>\n<p><strong>{}</strong> ({})</p>\n",
            kind.name(),
            escape(part_name)
        ));
        for text in texts.iter().map(|text| text.trim()) {
            if !text.is_empty() {
                sections.push_str(&format!("<p>{}</p>\n", escape(text)));
            }
        }
        sections.push_str("</section>\n");
    }
    match html.rfind("</body>") {
        Some(end) => html.insert_str(end, &sections),
        None => html.push_str(&sections),
    }
}

/// Validate a DOCX package's structure.
///
/// A relationship to a missing part, a part without a content type, an
/// undeclared Markup Compatibility prefix, a related part that is not
/// well-formed XML, a paragraph, character, or table style id that no style
/// defines, and broken comment ownership are structural errors. Advisory
/// findings (empty paragraphs, heading level gaps, missing metadata) are
/// warnings.
fn validate_package(package: OpcPackage) -> Result<ValidationReport> {
    let mut errors: Vec<String> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();

    // --- Structural errors ---

    // Every relationship must point at a part that is actually in the package.
    if let Some(doc_part) = package.main_document_part() {
        if package.get_part(&doc_part).is_none() {
            errors.push(format!("main document part {doc_part} is missing"));
        }
        if let Some(rels) = package.get_part_rels(&doc_part) {
            for rel in &rels.items {
                // External targets live outside the package by definition.
                if rel.target_mode.as_deref() == Some("External") {
                    continue;
                }
                let target = OpcPackage::resolve_rel_target(&doc_part, &rel.target);
                if package.get_part(&target).is_none() {
                    errors.push(format!(
                        "relationship {} points at missing part {target}",
                        rel.id
                    ));
                }
            }
        }
    } else {
        errors.push("package declares no main document relationship".to_string());
    }

    // Every part needs a declared content type.
    let mut part_names = package.parts.keys().collect::<Vec<_>>();
    part_names.sort();
    for part_name in part_names {
        if package.content_types.content_type_for(part_name).is_none() {
            errors.push(format!("part {part_name} has no declared content type"));
        }
    }

    // Markup Compatibility requires every prefix that `mc:Ignorable` or
    // `mc:MustUnderstand` lists to be declared, and a consumer may reject a
    // part that breaks the rule. A part that does not parse as XML is
    // outside this check.
    let mut xml_parts = package
        .parts
        .iter()
        .filter(|(part_name, _)| {
            package
                .content_types
                .content_type_for(part_name)
                .is_some_and(|content_type| content_type.ends_with("xml"))
        })
        .collect::<Vec<_>>();
    xml_parts.sort();
    for (part_name, xml) in xml_parts {
        let Ok(findings) = rdocx_oxml::namespace::undeclared_compatibility_prefixes(xml) else {
            continue;
        };
        for (attribute, prefix) in findings {
            errors.push(format!(
                "part {part_name} lists undeclared prefix `{prefix}` in mc:{attribute}"
            ));
        }
    }

    // Every XML part the main document relates to must be well formed. The
    // style ids that its story parts name are checked once the document opens.
    let mut style_references = Vec::new();
    if let Some(doc_part) = package.main_document_part() {
        let mut parts = vec![(doc_part.clone(), true)];
        for rel in package
            .get_part_rels(&doc_part)
            .iter()
            .flat_map(|rels| &rels.items)
        {
            if rel.target_mode.as_deref() == Some("External") {
                continue;
            }
            let target = OpcPackage::resolve_rel_target(&doc_part, &rel.target);
            let story = STORY_RELATIONSHIPS.contains(&rel.rel_type.as_str());
            if !parts.iter().any(|(known, _)| *known == target) {
                parts.push((target, story));
            }
        }
        for (part_name, story) in parts {
            let is_xml = package
                .content_types
                .content_type_for(&part_name)
                .is_some_and(|content_type| {
                    content_type.ends_with("+xml") || content_type.ends_with("/xml")
                });
            let Some(xml) = package.get_part(&part_name).filter(|_| is_xml) else {
                continue;
            };
            match xml_style_references(xml) {
                Err(detail) => {
                    errors.push(format!("part {part_name} is not well-formed XML: {detail}"));
                }
                Ok(references) if story => style_references.extend(
                    references
                        .into_iter()
                        .map(|(kind, style_id)| (kind, style_id, part_name.clone())),
                ),
                Ok(_) => {}
            }
        }
    }

    // A malformed part can keep the document from opening. The errors found
    // so far explain why, so they are reported together with the failure.
    let doc = match Document::from_package(package) {
        Ok(doc) => doc,
        Err(error) if !errors.is_empty() => {
            errors.push(format!("the document does not open: {error}"));
            return Ok(ValidationReport { errors, warnings });
        }
        Err(error) => return Err(error),
    };

    // Word falls back to the default style for a style id that no style
    // defines, so the style that the document names is silently lost.
    for (kind, style_id, part_name) in style_references {
        let issue = format!("{kind} style {style_id:?} used in {part_name} is not defined");
        if doc.style(&style_id).is_none() && !errors.contains(&issue) {
            errors.push(issue);
        }
    }

    if let Err(error) = doc.validate_comment_ownership() {
        errors.push(format!("invalid comment ownership: {error}"));
    }

    // --- Advisory findings ---

    if doc.content_count() == 0 {
        warnings.push("Document has no content (no paragraphs or tables)".to_string());
    }

    let empty_count = doc
        .paragraphs()
        .iter()
        .filter(|p| p.text().trim().is_empty())
        .count();
    if empty_count > 0 {
        warnings.push(format!("{empty_count} empty paragraph(s) found"));
    }

    let mut prev_level: Option<u32> = None;
    for (level, _) in doc.headings() {
        if let Some(prev) = prev_level
            && level > prev + 1
        {
            warnings.push(format!(
                "Heading level gap: Heading{prev} -> Heading{level} (skipped level(s))"
            ));
        }
        prev_level = Some(level);
    }

    if doc.title().is_none() {
        warnings.push("Missing document title".to_string());
    }
    if doc.author().is_none() {
        warnings.push("Missing document author".to_string());
    }

    Ok(ValidationReport { errors, warnings })
}

/// Read `xml` as one well-formed element tree and return the paragraph,
/// character, and table style ids it names. The ids inside a tracked property
/// change are left out, because they record the formatting before the change,
/// and so are those inside `mc:Fallback`, which Word does not read. An empty
/// id names no style.
fn xml_style_references(xml: &[u8]) -> std::result::Result<Vec<(&'static str, String)>, String> {
    let mut reader = NsReader::from_reader(xml);
    // One entry per open element: whether its style ids are left out.
    let mut open = Vec::new();
    let mut roots = 0usize;
    let mut references = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
            Ok(read) => read,
            Err(error) => return Err(format!("{error} at byte {}", reader.error_position())),
        };
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let is_compatibility =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MC_NS.as_bytes());
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if open.is_empty() {
                    roots += 1;
                    if roots > 1 {
                        return Err("the part has more than one root element".to_owned());
                    }
                }
                let local_name = element.local_name();
                let kind = match local_name.as_ref() {
                    b"pStyle" => Some("paragraph"),
                    b"rStyle" => Some("character"),
                    b"tblStyle" => Some("table"),
                    _ => None,
                };
                if let Some(kind) = kind.filter(|_| is_word && !open.contains(&true))
                    && let Some(style_id) = word_value(&reader, element)?
                    && !style_id.is_empty()
                {
                    references.push((kind, style_id));
                }
                if matches!(event, Event::Start(_)) {
                    open.push(
                        (is_word && local_name.as_ref().ends_with(b"PrChange"))
                            || (is_compatibility && local_name.as_ref() == b"Fallback"),
                    );
                }
            }
            Event::End(_) => {
                open.pop();
            }
            Event::Text(text) if open.is_empty() && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err("the part has text outside its root element".to_owned());
            }
            Event::CData(_) if open.is_empty() => {
                return Err("the part has CDATA outside its root element".to_owned());
            }
            Event::Eof if roots == 0 => return Err("the part has no root element".to_owned()),
            Event::Eof if !open.is_empty() => {
                return Err(format!(
                    "the part ends inside {} unclosed element(s)",
                    open.len()
                ));
            }
            Event::Eof => return Ok(references),
            _ => {}
        }
        buffer.clear();
    }
}

/// Return the `w:val` attribute of a WordprocessingML element.
fn word_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> std::result::Result<Option<String>, String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
            && local_name.as_ref() == b"val"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| error.to_string())?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}
