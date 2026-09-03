//! Deterministic EPUB 3 export for the native Word facade.

use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::{Seek, SeekFrom, Write};
use std::path::Path;

use oxml_opc::relationship::rel_types;
use rdocx_oxml::document::{BodyContent, CT_Document};
use rdocx_oxml::drawing::{CT_Drawing, CT_Inline};
use rdocx_oxml::numbering::{CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering, ST_NumberFormat};
use rdocx_oxml::styles::{CT_Style, CT_Styles, StyleType};
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R, Field, HyperlinkSpan, RunContent};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::{Document, Error, Result};

const MAX_EPUB_BYTES: usize = 64 * 1024 * 1024;
const MAX_SOURCE_TEXT_BYTES: usize = 8 * 1024 * 1024;
const MAX_MEDIA_BYTES: usize = 16 * 1024 * 1024;
const MAX_BODY_ITEMS: usize = 100_000;
const MAX_MEDIA_ITEMS: usize = 4_096;
const MAX_DIAGNOSTICS: usize = 10_000;
const MAX_NESTING_DEPTH: usize = 64;
const MAX_PROJECTED_NODES: usize = 100_000;
const MAX_IMAGE_OCCURRENCES: usize = 4_096;
const MAX_RELATIONSHIPS: usize = 16_384;
const MAX_STYLE_ITEMS: usize = 4_096;
const MAX_NUMBERING_ITEMS: usize = 4_096;
const MAX_PROJECTION_KEY_BYTES: usize = 1024 * 1024;

const CONTAINER_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="EPUB/package.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>
"#;

const STYLESHEET: &str = r#"body{font-family:serif;line-height:1.4;margin:5%;}img{height:auto;max-width:100%;}table{border-collapse:collapse;}td,th{padding:.25em;}nav ol{list-style:none;padding-left:1.25em;}ul.no-marker{list-style:none;}"#;

/// One stable report of source content that EPUB cannot preserve exactly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EpubDiagnostic {
    pub path: String,
    pub message: String,
}

/// Serialized EPUB bytes together with every lossy-conversion diagnostic.
pub struct EpubWriteResult {
    pub bytes: Vec<u8>,
    pub diagnostics: Vec<EpubDiagnostic>,
}

impl Document {
    /// Serialize this document as a bounded deterministic EPUB 3 publication.
    pub fn to_epub_bytes(&self) -> Result<EpubWriteResult> {
        EpubWriter::new(self).write()
    }

    /// Serialize and atomically save EPUB, returning lossy-conversion diagnostics.
    pub fn save_epub<P: AsRef<Path>>(&self, path: P) -> Result<Vec<EpubDiagnostic>> {
        let result = self.to_epub_bytes()?;
        crate::document::write_atomic_file(
            path.as_ref(),
            &result.bytes,
            "invalid EPUB file name",
            "could not allocate EPUB-save staging file",
        )?;
        Ok(result.diagnostics)
    }
}

#[derive(Clone)]
struct Heading {
    body_index: usize,
    level: u32,
    text: String,
    anchor: String,
    href: String,
}

struct NavNode {
    heading_index: usize,
    children: Vec<NavNode>,
}

struct SpineItem {
    id: String,
    href: String,
    title: String,
    start: usize,
    end: usize,
    xhtml: String,
}

struct MediaItem {
    relationship_id: String,
    href: String,
    media_type: String,
    data: Vec<u8>,
}

struct EpubWriter<'a> {
    document: &'a Document,
    diagnostics: Vec<EpubDiagnostic>,
}

impl<'a> EpubWriter<'a> {
    fn new(document: &'a Document) -> Self {
        Self {
            document,
            diagnostics: Vec::new(),
        }
    }

    fn write(mut self) -> Result<EpubWriteResult> {
        self.check_source_limits()?;
        self.check_relationship_limits()?;

        let styles = render_styles(&self.document.styles)?;
        let numbering = render_numbering(self.document.numbering.as_ref())?;
        let (media, html_images) = self.media_items()?;
        let mut input = rdocx_html::HtmlInput {
            document: CT_Document::new(),
            styles,
            numbering,
            images: html_images,
            hyperlink_urls: self.epub_hyperlinks(),
        };
        self.collect_diagnostics()?;
        let mut headings = self.headings();
        let mut spine = self.spine_items(&headings);

        let mut spine_index = 0;
        for heading in &mut headings {
            while spine_index + 1 < spine.len() && heading.body_index >= spine[spine_index].end {
                spine_index += 1;
            }
            let item = &spine[spine_index];
            if (item.start..item.end).contains(&heading.body_index) {
                heading.href = format!("{}#{}", item.href, heading.anchor);
            }
        }

        let heading_anchors = headings
            .iter()
            .map(|heading| (heading.body_index, heading.anchor.clone()))
            .collect::<HashMap<_, _>>();
        let mut list_counters = HashMap::new();

        for item in &mut spine {
            let fragment = emit_spine_fragment(
                &self.document.document.body.content,
                &mut input,
                &media,
                &heading_anchors,
                &mut list_counters,
                item.start,
                item.end,
            );
            let fragment = fragment?;
            item.xhtml = xhtml_document(&item.title, &fragment);
            if item.xhtml.len() > MAX_EPUB_BYTES {
                return Err(epub_error("generated XHTML exceeds the EPUB output limit"));
            }
        }

        let nav_tree = build_nav_tree(&headings);
        let title = self
            .document
            .core_properties
            .as_ref()
            .and_then(|properties| properties.title.as_deref())
            .filter(|title| !title.trim().is_empty())
            .unwrap_or("Untitled document");
        let author = self
            .document
            .core_properties
            .as_ref()
            .and_then(|properties| properties.creator.as_deref())
            .filter(|author| !author.trim().is_empty())
            .unwrap_or("Unknown author");
        let nav = navigation_document(title, &headings, &nav_tree, &spine);
        let identifier = publication_identifier(title, author, &spine, &media);
        let package = package_document(title, author, &identifier, &spine, &media);
        ensure_xml_10("EPUB/package.opf", &package)?;
        ensure_xml_10("EPUB/nav.xhtml", &nav)?;
        for item in &spine {
            ensure_xml_10(&format!("EPUB/{}", item.href), &item.xhtml)?;
        }
        let bytes = write_archive(&package, &nav, &spine, &media)?;

        Ok(EpubWriteResult {
            bytes,
            diagnostics: self.diagnostics,
        })
    }

    fn check_source_limits(&self) -> Result<()> {
        let mut item_count = 0_usize;
        let mut text_bytes = 0_usize;
        let mut projected_nodes = 0_usize;
        let mut image_occurrences = 0_usize;
        if let Some(properties) = &self.document.core_properties {
            for (location, value) in [
                ("document title", properties.title.as_deref()),
                ("document author", properties.creator.as_deref()),
            ] {
                if let Some(value) = value {
                    add_source_bytes(&mut text_bytes, value.len())?;
                    ensure_xml_value(location, value)?;
                }
            }
        }
        if let Some(background) = &self.document.document.background_xml {
            add_source_bytes(&mut text_bytes, background.len())?;
        }
        for content in &self.document.document.body.content {
            measure_body_content(
                content,
                0,
                &mut item_count,
                &mut text_bytes,
                &mut projected_nodes,
                &mut image_occurrences,
            )?;
        }
        if image_occurrences > MAX_IMAGE_OCCURRENCES {
            return Err(epub_error(
                "document has too many image occurrences for EPUB export",
            ));
        }
        let projected_bytes = text_bytes
            .checked_mul(6)
            .and_then(|bytes| {
                projected_nodes
                    .checked_mul(256)
                    .and_then(|markup| bytes.checked_add(markup))
            })
            .ok_or_else(|| epub_error("projected XHTML size overflow during EPUB export"))?;
        if projected_bytes > MAX_EPUB_BYTES / 2 {
            return Err(epub_error(
                "projected XHTML exceeds the EPUB intermediate limit",
            ));
        }
        Ok(())
    }

    fn check_relationship_limits(&self) -> Result<()> {
        let Some(relationships) = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name)
        else {
            return Ok(());
        };
        if relationships.items.len() > MAX_RELATIONSHIPS {
            return Err(epub_error(
                "document has too many relationships for EPUB export",
            ));
        }
        let mut bytes = 0_usize;
        for relationship in &relationships.items {
            for value in [
                relationship.id.as_str(),
                relationship.rel_type.as_str(),
                relationship.target.as_str(),
                relationship.target_mode.as_deref().unwrap_or(""),
            ] {
                bytes = bytes
                    .checked_add(value.len())
                    .ok_or_else(|| epub_error("relationship size overflow during EPUB export"))?;
                if bytes > MAX_PROJECTION_KEY_BYTES {
                    return Err(epub_error(
                        "document relationships exceed the EPUB projection limit",
                    ));
                }
                ensure_xml_value("relationship value", value)?;
            }
        }
        Ok(())
    }

    fn headings(&self) -> Vec<Heading> {
        let mut headings = Vec::new();
        for (body_index, content) in self.document.document.body.content.iter().enumerate() {
            let BodyContent::Paragraph(paragraph) = content else {
                continue;
            };
            let Some(level) = heading_level(paragraph) else {
                continue;
            };
            let ordinal = headings.len() + 1;
            headings.push(Heading {
                body_index,
                level,
                text: projected_paragraph_text(paragraph),
                anchor: format!("heading-{ordinal:04}"),
                href: String::new(),
            });
        }
        headings
    }

    fn spine_items(&self, headings: &[Heading]) -> Vec<SpineItem> {
        let roots = root_heading_indexes(headings);
        if roots.is_empty() {
            return vec![SpineItem {
                id: "document".to_owned(),
                href: "document.xhtml".to_owned(),
                title: "Document".to_owned(),
                start: 0,
                end: self.document.document.body.content.len(),
                xhtml: String::new(),
            }];
        }

        let mut items = Vec::new();
        let first_body_index = headings[roots[0]].body_index;
        if first_body_index > 0 {
            items.push(SpineItem {
                id: "front".to_owned(),
                href: "front.xhtml".to_owned(),
                title: "Front matter".to_owned(),
                start: 0,
                end: first_body_index,
                xhtml: String::new(),
            });
        }
        for (root_ordinal, heading_index) in roots.iter().copied().enumerate() {
            let start = headings[heading_index].body_index;
            let end = roots
                .get(root_ordinal + 1)
                .map(|next| headings[*next].body_index)
                .unwrap_or(self.document.document.body.content.len());
            let number = root_ordinal + 1;
            items.push(SpineItem {
                id: format!("chapter-{number:03}"),
                href: format!("chapter-{number:03}.xhtml"),
                title: headings[heading_index].text.clone(),
                start,
                end,
                xhtml: String::new(),
            });
        }
        items
    }

    fn media_items(&mut self) -> Result<(Vec<MediaItem>, HashMap<String, rdocx_html::ImageData>)> {
        let mut relationship_ids = referenced_drawing_ids(&self.document.document.body.content);
        relationship_ids.sort_unstable();
        relationship_ids.dedup();
        if relationship_ids.len() > MAX_MEDIA_ITEMS {
            return Err(epub_error("document has too many images for EPUB export"));
        }
        let relationships = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name);
        let total = relationship_ids.iter().try_fold(0_usize, |total, id| {
            let Some(relationship) = relationships.and_then(|items| items.get_by_id(id)) else {
                return Some(total);
            };
            if relationship.rel_type != rel_types::IMAGE
                || relationship.target_mode.as_deref() == Some("External")
            {
                return Some(total);
            }
            let part_name = oxml_opc::OpcPackage::resolve_rel_target(
                &self.document.doc_part_name,
                &relationship.target,
            );
            let Some(data) = self.document.package.get_part(&part_name) else {
                return Some(total);
            };
            if validated_epub_image(data).is_none() {
                return Some(total);
            }
            total.checked_add(data.len())
        });
        if total.is_none_or(|total| total > MAX_MEDIA_BYTES) {
            return Err(epub_error("document images exceed the EPUB media limit"));
        }

        let mut media: Vec<MediaItem> = Vec::new();
        let mut html_images = HashMap::new();
        for (ordinal, relationship_id) in relationship_ids.into_iter().enumerate() {
            let Some(relationship) =
                relationships.and_then(|items| items.get_by_id(relationship_id))
            else {
                continue;
            };
            if relationship.rel_type != rel_types::IMAGE
                || relationship.target_mode.as_deref() == Some("External")
            {
                continue;
            }
            let part_name = oxml_opc::OpcPackage::resolve_rel_target(
                &self.document.doc_part_name,
                &relationship.target,
            );
            let Some(data) = self.document.package.get_part(&part_name) else {
                continue;
            };
            let Some(format) = validated_epub_image(data) else {
                continue;
            };
            let content_type = format.content_type().to_owned();
            let extension = format.extension();
            html_images.insert(
                relationship.id.clone(),
                rdocx_html::ImageData {
                    data: (ordinal as u64).to_le_bytes().to_vec(),
                    content_type: content_type.clone(),
                },
            );
            if let Some(existing) = media.iter().find(|existing| {
                existing.media_type == content_type && existing.data.as_slice() == data
            }) {
                media.push(MediaItem {
                    relationship_id: relationship.id.clone(),
                    href: existing.href.clone(),
                    media_type: existing.media_type.clone(),
                    data: existing.data.clone(),
                });
                continue;
            }
            let number = media
                .iter()
                .filter(|item| item.href.starts_with("images/image-"))
                .count()
                + 1;
            media.push(MediaItem {
                relationship_id: relationship.id.clone(),
                href: format!("images/image-{number:03}.{extension}"),
                media_type: content_type,
                data: data.to_vec(),
            });
        }
        Ok((media, html_images))
    }

    fn epub_hyperlinks(&self) -> HashMap<String, String> {
        let mut relationship_ids = referenced_hyperlink_ids(&self.document.document.body.content);
        relationship_ids.sort_unstable();
        relationship_ids.dedup();
        let mut hyperlinks = HashMap::new();
        let relationships = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name);
        for relationship_id in relationship_ids {
            let Some(relationship) =
                relationships.and_then(|items| items.get_by_id(relationship_id))
            else {
                continue;
            };
            if relationship.rel_type == rel_types::HYPERLINK
                && relationship.target_mode.as_deref() == Some("External")
                && let Some(target) = safe_absolute_url(&relationship.target)
            {
                hyperlinks.insert(relationship.id.clone(), target.to_owned());
            }
        }
        hyperlinks
    }

    fn collect_diagnostics(&mut self) -> Result<()> {
        if self.document.document.background_xml.is_some() {
            self.diagnose(
                "document/background".to_owned(),
                "document background was dropped during EPUB export".to_owned(),
            )?;
        }
        if let Some(properties) = &self.document.core_properties {
            for (present, name) in [
                (properties.subject.is_some(), "subject"),
                (properties.description.is_some(), "description"),
                (properties.keywords.is_some(), "keywords"),
                (properties.last_modified_by.is_some(), "last-modified-by"),
                (properties.created.is_some(), "created"),
                (properties.modified.is_some(), "modified"),
            ] {
                if present {
                    self.diagnose(
                        format!("metadata/{name}"),
                        format!("document {name} metadata was dropped during EPUB export"),
                    )?;
                }
            }
        }
        if let Some(properties) = &self.document.custom_properties {
            for (index, _) in properties.properties.iter().enumerate() {
                self.diagnose(
                    format!("metadata/custom-property[{index}]"),
                    "custom document property was dropped during EPUB export".to_owned(),
                )?;
            }
        }
        for (index, content) in self.document.document.body.content.iter().enumerate() {
            self.scan_body_content(content, &format!("body[{index}]"), 0)?;
        }
        if self.document.document.body.sect_pr.is_some() {
            self.diagnose(
                "body/properties/section".to_owned(),
                "final section properties were dropped during EPUB export".to_owned(),
            )?;
        }
        Ok(())
    }

    fn scan_body_content(&mut self, content: &BodyContent, path: &str, depth: usize) -> Result<()> {
        match content {
            BodyContent::Paragraph(paragraph) => self.scan_paragraph(paragraph, path),
            BodyContent::Table(table) => self.scan_table(table, path, depth + 1),
            BodyContent::ContentControl(_) => self.diagnose(
                path.to_owned(),
                "body content control was dropped during EPUB export".to_owned(),
            ),
            BodyContent::RawXml(_) => self.diagnose(
                path.to_owned(),
                "unmodelled body XML was dropped during EPUB export".to_owned(),
            ),
        }
    }

    fn scan_table(&mut self, table: &CT_Tbl, path: &str, depth: usize) -> Result<()> {
        if depth > MAX_NESTING_DEPTH {
            return Err(epub_error("table nesting exceeds the EPUB depth limit"));
        }
        if table.grid.is_some() {
            self.diagnose(
                format!("{path}/grid"),
                "table grid widths were dropped during EPUB export".to_owned(),
            )?;
        }
        if let Some(properties) = &table.properties {
            self.scan_table_properties(properties, path)?;
        }
        for (raw_index, _) in table.extra_xml.iter().enumerate() {
            self.diagnose(
                format!("{path}/xml[{raw_index}]"),
                "unmodelled table XML was dropped during EPUB export".to_owned(),
            )?;
        }
        for (control_index, _) in table.content_controls.iter().enumerate() {
            self.diagnose(
                format!("{path}/content-control[{control_index}]"),
                "table row content control was dropped during EPUB export".to_owned(),
            )?;
        }
        for (row_index, row) in table.rows.iter().enumerate() {
            let row_path = format!("{path}/row[{row_index}]");
            if let Some(properties) = &row.properties {
                self.scan_row_properties(properties, &row_path)?;
            }
            for (raw_index, _) in row.extra_xml.iter().enumerate() {
                self.diagnose(
                    format!("{row_path}/xml[{raw_index}]"),
                    "unmodelled table-row XML was dropped during EPUB export".to_owned(),
                )?;
            }
            for (control_index, _) in row.content_controls.iter().enumerate() {
                self.diagnose(
                    format!("{row_path}/content-control[{control_index}]"),
                    "table-cell content control was dropped during EPUB export".to_owned(),
                )?;
            }
            for (cell_index, cell) in row.cells.iter().enumerate() {
                let cell_path = format!("{row_path}/cell[{cell_index}]");
                if let Some(properties) = &cell.properties {
                    self.scan_cell_properties(properties, &cell_path)?;
                }
                for (raw_index, _) in cell.extra_xml.iter().enumerate() {
                    self.diagnose(
                        format!("{cell_path}/xml[{raw_index}]"),
                        "unmodelled table-cell XML was dropped during EPUB export".to_owned(),
                    )?;
                }
                for (content_index, content) in cell.content.iter().enumerate() {
                    let child_path = format!("{cell_path}/content[{content_index}]");
                    match content {
                        CellContent::Paragraph(paragraph) => {
                            if detect_list(paragraph, self.document.numbering.as_ref()).is_some() {
                                self.diagnose(
                                    format!("{child_path}/properties/numbering"),
                                    "table-cell list semantics were flattened during EPUB export"
                                        .to_owned(),
                                )?;
                            }
                            self.scan_paragraph(paragraph, &child_path)?
                        }
                        CellContent::Table(nested) => {
                            self.scan_table(nested, &child_path, depth + 1)?
                        }
                        CellContent::ContentControl(_) => self.diagnose(
                            child_path,
                            "table-cell content control was dropped during EPUB export".to_owned(),
                        )?,
                    }
                }
            }
        }
        Ok(())
    }

    fn scan_table_properties(
        &mut self,
        properties: &rdocx_oxml::table::CT_TblPr,
        path: &str,
    ) -> Result<()> {
        for (present, name, message) in [
            (
                properties.style_id.is_some(),
                "style",
                "table style was dropped during EPUB export",
            ),
            (
                properties.width.is_some(),
                "width",
                "table width was dropped during EPUB export",
            ),
            (
                properties.cell_margin.is_some(),
                "cell-margin",
                "table cell margins were dropped during EPUB export",
            ),
            (
                properties.layout.is_some(),
                "layout",
                "table layout mode was dropped during EPUB export",
            ),
            (
                properties.indent.is_some(),
                "indent",
                "table indent was dropped during EPUB export",
            ),
            (
                properties.shading.is_some(),
                "shading",
                "table shading was dropped during EPUB export",
            ),
            (
                properties.look.is_some(),
                "look",
                "table conditional look was dropped during EPUB export",
            ),
            (
                properties.change.is_some(),
                "revision",
                "table property revision was dropped during EPUB export",
            ),
        ] {
            if present {
                self.diagnose(format!("{path}/properties/{name}"), message.to_owned())?;
            }
        }
        if properties.borders.is_some() {
            self.diagnose(
                format!("{path}/properties/borders"),
                "table border details were simplified during EPUB export".to_owned(),
            )?;
        }
        for (index, _) in properties.revision_xml.iter().enumerate() {
            self.diagnose(
                format!("{path}/properties/xml[{index}]"),
                "unmodelled table property XML was dropped during EPUB export".to_owned(),
            )?;
        }
        Ok(())
    }

    fn scan_row_properties(
        &mut self,
        properties: &rdocx_oxml::table::CT_TrPr,
        path: &str,
    ) -> Result<()> {
        for (present, name) in [
            (properties.height.is_some(), "height"),
            (properties.height_rule.is_some(), "height-rule"),
            (properties.header.is_some(), "repeat-header"),
            (properties.jc.is_some(), "alignment"),
            (properties.cant_split.is_some(), "cant-split"),
            (properties.cnf_style.is_some(), "conditional-style"),
        ] {
            if present {
                self.diagnose(
                    format!("{path}/properties/{name}"),
                    format!("table-row {name} was dropped during EPUB export"),
                )?;
            }
        }
        for (index, _) in properties.revision_markers.iter().enumerate() {
            self.diagnose(
                format!("{path}/properties/revision[{index}]"),
                "table-row revision marker was dropped during EPUB export".to_owned(),
            )?;
        }
        for (index, _) in properties.revision_xml.iter().enumerate() {
            self.diagnose(
                format!("{path}/properties/xml[{index}]"),
                "unmodelled table-row property XML was dropped during EPUB export".to_owned(),
            )?;
        }
        Ok(())
    }

    fn scan_cell_properties(
        &mut self,
        properties: &rdocx_oxml::table::CT_TcPr,
        path: &str,
    ) -> Result<()> {
        for (present, name) in [
            (properties.width.is_some(), "width"),
            (properties.no_wrap.is_some(), "no-wrap"),
            (properties.text_direction.is_some(), "text-direction"),
            (properties.cnf_style.is_some(), "conditional-style"),
        ] {
            if present {
                self.diagnose(
                    format!("{path}/properties/{name}"),
                    format!("table-cell {name} was dropped during EPUB export"),
                )?;
            }
        }
        if properties.borders.is_some() {
            self.diagnose(
                format!("{path}/properties/borders"),
                "table-cell border details were simplified during EPUB export".to_owned(),
            )?;
        }
        if properties
            .shading
            .as_ref()
            .is_some_and(shading_is_simplified)
        {
            self.diagnose(
                format!("{path}/properties/shading"),
                "table-cell shading pattern, foreground, or invalid colour was simplified during EPUB export"
                    .to_owned(),
            )?;
        }
        for (index, _) in properties.extra_xml.iter().enumerate() {
            self.diagnose(
                format!("{path}/properties/xml[{index}]"),
                "unmodelled table-cell property XML was dropped during EPUB export".to_owned(),
            )?;
        }
        Ok(())
    }

    fn scan_paragraph(&mut self, paragraph: &CT_P, path: &str) -> Result<()> {
        if paragraph
            .properties
            .as_ref()
            .and_then(|properties| properties.style_id.as_deref())
            .is_none()
            && default_paragraph_style_has_visible_effects(&self.document.styles, paragraph)
        {
            self.diagnose(
                format!("{path}/properties/default-style"),
                "default paragraph or run formatting was dropped during EPUB export".to_owned(),
            )?;
        }
        if let Some(properties) = &paragraph.properties {
            if let Some(style_id) = properties.style_id.as_deref() {
                let style = self.document.styles.get_by_id(style_id);
                if style.is_none_or(style_has_unprojected_visuals) {
                    self.diagnose(
                        format!("{path}/properties/style"),
                        "paragraph style formatting was dropped during EPUB export".to_owned(),
                    )?;
                }
            }
            if projected_heading_level(paragraph, &self.document.styles)
                .is_some_and(|level| level > 6)
            {
                let lowered = properties.outline_lvl.is_some_and(|level| level >= 6)
                    || heading_level(paragraph).is_some_and(|level| level > 6);
                self.diagnose(
                    format!("{path}/properties/heading-level"),
                    if lowered {
                        "heading level above 6 was reduced to level 6 during EPUB export"
                    } else {
                        "style-derived heading level above 6 was flattened to a paragraph during EPUB export"
                    }
                    .to_owned(),
                )?;
            }
            if properties
                .shading
                .as_ref()
                .is_some_and(shading_is_simplified)
            {
                self.diagnose(
                    format!("{path}/properties/shading"),
                    "paragraph shading pattern, foreground, or invalid colour was simplified during EPUB export"
                        .to_owned(),
                )?;
            }
            if properties.num_id.is_some_and(|id| id != 0)
                && detect_list(paragraph, self.document.numbering.as_ref()).is_none()
                && list_definition(paragraph, self.document.numbering.as_ref()).is_none_or(
                    |level| !matches!(level.num_fmt.as_ref(), Some(ST_NumberFormat::Other(_))),
                )
            {
                self.diagnose(
                    format!("{path}/properties/numbering"),
                    "unresolved list definition was flattened during EPUB export".to_owned(),
                )?;
            }
            if let Some(level) = list_definition(paragraph, self.document.numbering.as_ref()) {
                let producer_defined =
                    matches!(level.num_fmt.as_ref(), Some(ST_NumberFormat::Other(_)));
                if producer_defined {
                    self.diagnose(
                        format!("{path}/properties/numbering/format"),
                        "producer-defined numbering format was emitted without a marker during EPUB export"
                            .to_owned(),
                    )?;
                    if level.start.is_some_and(|start| start != 1) {
                        self.diagnose(
                            format!("{path}/properties/numbering/start"),
                            "producer-defined list start value was dropped during EPUB export"
                                .to_owned(),
                        )?;
                    }
                } else if level.num_fmt == Some(ST_NumberFormat::Ordinal) {
                    self.diagnose(
                        format!("{path}/properties/numbering/format"),
                        "ordinal list markers were reduced to decimal during EPUB export"
                            .to_owned(),
                    )?;
                }
                if let Some(marker) = &level.lvl_text {
                    let standard = format!("%{}.", level.ilvl + 1);
                    if producer_defined
                        || level.num_fmt == Some(ST_NumberFormat::Bullet)
                        || marker != &standard
                    {
                        self.diagnose(
                            format!("{path}/properties/numbering/marker"),
                            if producer_defined {
                                "list marker text was dropped during EPUB export"
                            } else {
                                "custom list marker text was replaced by EPUB list semantics"
                            }
                            .to_owned(),
                        )?;
                    }
                }
                if level.ppr.is_some() || level.ppr_raw.is_some() {
                    self.diagnose(
                        format!("{path}/properties/numbering/paragraph-style"),
                        "list-level paragraph styling was dropped during EPUB export".to_owned(),
                    )?;
                }
                if level.rpr.is_some() || level.rpr_raw.is_some() {
                    self.diagnose(
                        format!("{path}/properties/numbering/marker-style"),
                        "list marker run styling was dropped during EPUB export".to_owned(),
                    )?;
                }
                if level.lvl_jc.is_some() {
                    self.diagnose(
                        format!("{path}/properties/numbering/alignment"),
                        "list marker alignment was dropped during EPUB export".to_owned(),
                    )?;
                }
                if level.suffix.is_some() {
                    self.diagnose(
                        format!("{path}/properties/numbering/suffix"),
                        if producer_defined {
                            "list marker suffix was dropped during EPUB export"
                        } else {
                            "list marker suffix spacing was normalized during EPUB export"
                        }
                        .to_owned(),
                    )?;
                }
                if !level.extra_xml.is_empty() || !level.extra_attributes.is_empty() {
                    self.diagnose(
                        format!("{path}/properties/numbering/xml"),
                        "unmodelled list-level XML was dropped during EPUB export".to_owned(),
                    )?;
                }
            }
            for (present, name, message) in [
                (
                    properties.before_autospacing.is_some(),
                    "before-auto-spacing",
                    "paragraph automatic spacing was dropped during EPUB export",
                ),
                (
                    properties.after_autospacing.is_some(),
                    "after-auto-spacing",
                    "paragraph automatic spacing was dropped during EPUB export",
                ),
                (
                    properties.ind_hanging.is_some(),
                    "hanging-indent",
                    "paragraph hanging indent was dropped during EPUB export",
                ),
                (
                    properties.keep_next.is_some(),
                    "keep-next",
                    "paragraph keep-next was dropped during EPUB export",
                ),
                (
                    properties.keep_lines.is_some(),
                    "keep-lines",
                    "paragraph keep-lines was dropped during EPUB export",
                ),
                (
                    properties.page_break_before.is_some(),
                    "page-break-before",
                    "paragraph page-break-before was dropped during EPUB export",
                ),
                (
                    properties.widow_control.is_some(),
                    "widow-control",
                    "paragraph widow control was dropped during EPUB export",
                ),
                (
                    properties.suppress_auto_hyphens.is_some(),
                    "suppress-hyphens",
                    "paragraph hyphenation control was dropped during EPUB export",
                ),
                (
                    properties.borders.is_some(),
                    "borders",
                    "paragraph borders were dropped during EPUB export",
                ),
                (
                    properties.tabs.is_some(),
                    "tabs",
                    "paragraph tab stops were dropped during EPUB export",
                ),
                (
                    properties.rpr.is_some(),
                    "mark-properties",
                    "paragraph mark properties were dropped during EPUB export",
                ),
                (
                    properties.sect_pr.is_some(),
                    "section",
                    "paragraph section properties were dropped during EPUB export",
                ),
                (
                    properties.numbering_revision.is_some(),
                    "numbering-revision",
                    "paragraph numbering revision was dropped during EPUB export",
                ),
                (
                    properties.change.is_some(),
                    "revision",
                    "paragraph property revision was dropped during EPUB export",
                ),
            ] {
                if present {
                    self.diagnose(format!("{path}/properties/{name}"), message.to_owned())?;
                }
            }
            for (index, _) in properties.numbering_revision_xml.iter().enumerate() {
                self.diagnose(
                    format!("{path}/properties/numbering-xml[{index}]"),
                    "unmodelled paragraph numbering XML was dropped during EPUB export".to_owned(),
                )?;
            }
            for (index, _) in properties.revision_xml.iter().enumerate() {
                self.diagnose(
                    format!("{path}/properties/xml[{index}]"),
                    "unmodelled paragraph property XML was dropped during EPUB export".to_owned(),
                )?;
            }
        }
        // The typed source coordinate retains the parser's in-scope namespace
        // decision, including paragraph-local shadows of document bindings.
        let mut typed_raw_markers = HashMap::<(usize, usize, u8), usize>::new();
        for marker in &paragraph.bookmark_markers {
            let kind = if marker.is_start() { 2 } else { 3 };
            *typed_raw_markers
                .entry((marker.run_index(), marker.raw_before(), kind))
                .or_default() += 1;
        }
        for (run_index, raw_before, revision) in &paragraph.revisions {
            let kind = match revision.kind() {
                rdocx_oxml::revision::RevisionKind::Insertion => 10,
                rdocx_oxml::revision::RevisionKind::Deletion => 11,
                rdocx_oxml::revision::RevisionKind::MoveFrom => 12,
                rdocx_oxml::revision::RevisionKind::MoveTo => 13,
                _ => continue,
            };
            *typed_raw_markers
                .entry((*run_index, *raw_before, kind))
                .or_default() += 1;
        }
        let mut consumed_raw = vec![false; paragraph.extra_xml.len()];
        let mut raw_ordinals = HashMap::<usize, usize>::new();
        for (index, (run_index, raw)) in paragraph.extra_xml.iter().enumerate() {
            let raw_ordinal = raw_ordinals.entry(*run_index).or_default();
            let ordinal = *raw_ordinal;
            *raw_ordinal += 1;
            let Some(marker) = raw_marker_kind(raw) else {
                continue;
            };
            if let Some(remaining) = typed_raw_markers.get_mut(&(*run_index, ordinal, marker.kind))
                && *remaining > 0
            {
                *remaining -= 1;
                consumed_raw[index] = true;
            }
        }
        for (extra_index, consumed) in consumed_raw.into_iter().enumerate() {
            if !consumed {
                self.diagnose(
                    format!("{path}/xml[{extra_index}]"),
                    "unmodelled paragraph XML was dropped during EPUB export".to_owned(),
                )?;
            }
        }
        for (control_index, _) in paragraph.content_controls.iter().enumerate() {
            self.diagnose(
                format!("{path}/content-control[{control_index}]"),
                "run content control was dropped during EPUB export".to_owned(),
            )?;
        }
        for (revision_index, _) in paragraph.revisions.iter().enumerate() {
            self.diagnose(
                format!("{path}/revision[{revision_index}]"),
                "paragraph revision wrapper was flattened during EPUB export".to_owned(),
            )?;
        }
        for (marker_index, _) in paragraph.comment_ranges.iter().enumerate() {
            self.diagnose(
                format!("{path}/comment-range[{marker_index}]"),
                "comment range marker was dropped during EPUB export".to_owned(),
            )?;
        }
        for (marker_index, _) in paragraph.bookmark_markers.iter().enumerate() {
            self.diagnose(
                format!("{path}/bookmark[{marker_index}]"),
                "bookmark marker was dropped during EPUB export".to_owned(),
            )?;
        }
        for (hyperlink_index, hyperlink) in paragraph.hyperlinks.iter().enumerate() {
            let hyperlink_path = format!("{path}/hyperlink[{hyperlink_index}]");
            if hyperlink.anchor.is_some() {
                self.diagnose(
                    hyperlink_path.clone(),
                    "internal hyperlink anchor was dropped during EPUB export".to_owned(),
                )?;
            } else if let Some(reason) = self.hyperlink_loss_reason(hyperlink.rel_id.as_deref()) {
                self.diagnose(hyperlink_path.clone(), reason.to_owned())?;
            }
            if !hyperlink.extra_attributes.is_empty() || hyperlink.preserved_raw_before.is_some() {
                self.diagnose(
                    format!("{hyperlink_path}/owner-xml"),
                    "unmodelled hyperlink owner XML was dropped during EPUB export".to_owned(),
                )?;
            }
            for (raw_index, _) in hyperlink.extra_xml.iter().enumerate() {
                self.diagnose(
                    format!("{hyperlink_path}/xml[{raw_index}]"),
                    "unmodelled hyperlink child XML was dropped during EPUB export".to_owned(),
                )?;
            }
        }
        for (run_index, run) in paragraph.runs.iter().enumerate() {
            if let Some(properties) = &run.properties {
                self.scan_run_properties(properties, &format!("{path}/run[{run_index}]"))?;
            }
            for (raw_index, _) in run.extra_xml.iter().enumerate() {
                self.diagnose(
                    format!("{path}/run[{run_index}]/xml[{raw_index}]"),
                    "unmodelled run XML was dropped during EPUB export".to_owned(),
                )?;
            }
            for (drawing_index, _) in run.alt_drawings.iter().enumerate() {
                self.diagnose(
                    format!("{path}/run[{run_index}]/alternate-drawing[{drawing_index}]"),
                    "alternate drawing payload was dropped during EPUB export".to_owned(),
                )?;
            }
            for (content_index, content) in run.content.iter().enumerate() {
                let content_path = format!("{path}/run[{run_index}]/content[{content_index}]");
                match content {
                    RunContent::Text(text) if text.preserve_space => self.diagnose(
                        format!("{content_path}/space"),
                        "preserved Word text spacing was normalized during EPUB export".to_owned(),
                    )?,
                    RunContent::DeletedText(text) if text.preserve_space => {
                        self.diagnose(
                            format!("{content_path}/space"),
                            "preserved Word text spacing was normalized during EPUB export"
                                .to_owned(),
                        )?;
                        self.diagnose(
                            content_path,
                            "deleted-text revision semantics were flattened during EPUB export"
                                .to_owned(),
                        )?;
                    }
                    RunContent::Break(rdocx_oxml::text::BreakType::Column) => self.diagnose(
                        content_path,
                        "column break was simplified to a line break during EPUB export".to_owned(),
                    )?,
                    RunContent::DeletedText(_) => self.diagnose(
                        content_path,
                        "deleted-text revision semantics were flattened during EPUB export"
                            .to_owned(),
                    )?,
                    RunContent::Field(_) => self.diagnose(
                        content_path,
                        "field semantics were flattened to cached display during EPUB export"
                            .to_owned(),
                    )?,
                    RunContent::FootnoteRef { .. } => self.diagnose(
                        content_path,
                        "footnote reference was dropped during EPUB export".to_owned(),
                    )?,
                    RunContent::EndnoteRef { .. } => self.diagnose(
                        content_path,
                        "endnote reference was dropped during EPUB export".to_owned(),
                    )?,
                    RunContent::CommentReference { .. } => self.diagnose(
                        content_path,
                        "comment reference was dropped during EPUB export".to_owned(),
                    )?,
                    RunContent::Drawing(drawing) => {
                        let relationship_id = drawing
                            .inline
                            .as_ref()
                            .map(|image| image.embed_id.as_str())
                            .or_else(|| {
                                drawing.anchor.as_ref().map(|image| image.embed_id.as_str())
                            });
                        if let Some(reason) = self.image_loss_reason(relationship_id) {
                            self.diagnose(content_path.clone(), reason.to_owned())?;
                        } else if drawing.anchor.is_some() {
                            self.diagnose(
                                content_path.clone(),
                                "floating image placement was converted to inline flow".to_owned(),
                            )?;
                        }
                        let source = drawing
                            .inline
                            .as_ref()
                            .map(|image| {
                                (
                                    image.name.as_deref(),
                                    image.raw_xml.is_some(),
                                    image.extent_cx.0,
                                    image.extent_cy.0,
                                )
                            })
                            .or_else(|| {
                                drawing.anchor.as_ref().map(|image| {
                                    (
                                        image.name.as_deref(),
                                        image.raw_xml.is_some(),
                                        image.extent_cx.0,
                                        image.extent_cy.0,
                                    )
                                })
                            });
                        if let Some((name, has_raw_xml, width, height)) = source {
                            if name.is_some() {
                                self.diagnose(
                                    format!("{content_path}/name"),
                                    "drawing name was dropped during EPUB export".to_owned(),
                                )?;
                            }
                            if width != 0 || height != 0 {
                                self.diagnose(
                                    format!("{content_path}/extent"),
                                    "drawing extent was simplified to responsive EPUB sizing"
                                        .to_owned(),
                                )?;
                            }
                            if has_raw_xml {
                                self.diagnose(
                                    format!("{content_path}/xml"),
                                    "preserved drawing XML was dropped during EPUB export"
                                        .to_owned(),
                                )?;
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        Ok(())
    }

    fn scan_run_properties(
        &mut self,
        properties: &rdocx_oxml::properties::CT_RPr,
        path: &str,
    ) -> Result<()> {
        for (present, name, message) in [
            (
                properties.style_id.is_some(),
                "style",
                "run style was dropped during EPUB export",
            ),
            (
                properties.font_hansi.is_some() && properties.font_hansi != properties.font_ascii,
                "font-hansi",
                "alternate run font was dropped during EPUB export",
            ),
            (
                properties.font_east_asia.is_some(),
                "font-east-asia",
                "East Asian run font was dropped during EPUB export",
            ),
            (
                properties.font_cs.is_some(),
                "font-complex-script",
                "complex-script run font was dropped during EPUB export",
            ),
            (
                properties.font_ascii_theme.is_some() || properties.font_hansi_theme.is_some(),
                "theme-font",
                "theme run font was dropped during EPUB export",
            ),
            (
                properties.bold_cs.is_some() && properties.bold_cs != properties.bold,
                "bold-complex-script",
                "complex-script bold was dropped during EPUB export",
            ),
            (
                properties.italic_cs.is_some() && properties.italic_cs != properties.italic,
                "italic-complex-script",
                "complex-script italic was dropped during EPUB export",
            ),
            (
                properties.sz_cs.is_some() && properties.sz_cs != properties.sz,
                "size-complex-script",
                "complex-script font size was dropped during EPUB export",
            ),
            (
                properties.color_theme.is_some(),
                "theme-colour",
                "theme run colour was dropped during EPUB export",
            ),
            (
                properties.highlight.is_some(),
                "highlight",
                "keyword highlight colour was dropped during EPUB export",
            ),
            (
                properties.width_scale.is_some(),
                "width-scale",
                "run width scale was dropped during EPUB export",
            ),
            (
                properties.position.is_some(),
                "position",
                "run text position was dropped during EPUB export",
            ),
            (
                properties.vanish.is_some(),
                "hidden",
                "hidden-text semantics were dropped during EPUB export",
            ),
            (
                properties.change.is_some(),
                "revision",
                "run property revision was dropped during EPUB export",
            ),
        ] {
            if present {
                self.diagnose(format!("{path}/properties/{name}"), message.to_owned())?;
            }
        }
        if properties.dstrike.is_some() {
            self.diagnose(
                format!("{path}/properties/double-strike"),
                "double strikethrough was simplified during EPUB export".to_owned(),
            )?;
        }
        if properties.underline.is_some_and(|underline| {
            !matches!(
                underline,
                rdocx_oxml::shared::ST_Underline::None | rdocx_oxml::shared::ST_Underline::Single
            )
        }) {
            self.diagnose(
                format!("{path}/properties/underline"),
                "non-basic underline style was simplified to a single underline during EPUB export"
                    .to_owned(),
            )?;
        }
        if properties
            .shading
            .as_ref()
            .is_some_and(shading_is_simplified)
        {
            self.diagnose(
                format!("{path}/properties/shading"),
                "run shading pattern, foreground, or invalid colour was simplified during EPUB export"
                    .to_owned(),
            )?;
        }
        for (index, _) in properties.revision_markers.iter().enumerate() {
            self.diagnose(
                format!("{path}/properties/revision-marker[{index}]"),
                "run revision marker was dropped during EPUB export".to_owned(),
            )?;
        }
        for (index, _) in properties.revision_xml.iter().enumerate() {
            self.diagnose(
                format!("{path}/properties/xml[{index}]"),
                "unmodelled run property XML was dropped during EPUB export".to_owned(),
            )?;
        }
        Ok(())
    }

    fn image_loss_reason(&self, relationship_id: Option<&str>) -> Option<&'static str> {
        let Some(relationship_id) = relationship_id else {
            return Some("drawing without an image relationship was dropped during EPUB export");
        };
        let Some(relationship) = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name)
            .and_then(|relationships| relationships.get_by_id(relationship_id))
        else {
            return Some("unresolved image relationship was dropped during EPUB export");
        };
        if relationship.rel_type != rel_types::IMAGE
            || relationship.target_mode.as_deref() == Some("External")
        {
            return Some("non-package image relationship was dropped during EPUB export");
        }
        let part_name = oxml_opc::OpcPackage::resolve_rel_target(
            &self.document.doc_part_name,
            &relationship.target,
        );
        let Some(data) = self.document.package.get_part(&part_name) else {
            return Some("unresolved image part was dropped during EPUB export");
        };
        if validated_epub_image(data).is_none() {
            return Some("non-core EPUB image type was dropped during EPUB export");
        }
        None
    }

    fn hyperlink_loss_reason(&self, relationship_id: Option<&str>) -> Option<&'static str> {
        let Some(relationship_id) = relationship_id else {
            return Some(
                "hyperlink without an external relationship was dropped during EPUB export",
            );
        };
        let Some(relationship) = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name)
            .and_then(|relationships| relationships.get_by_id(relationship_id))
        else {
            return Some("unresolved hyperlink relationship was dropped during EPUB export");
        };
        if relationship.rel_type != rel_types::HYPERLINK
            || relationship.target_mode.as_deref() != Some("External")
        {
            return Some("non-external hyperlink relationship was dropped during EPUB export");
        }
        if safe_absolute_url(&relationship.target).is_none() {
            return Some(
                "empty, relative, or unsafe hyperlink target was dropped during EPUB export",
            );
        }
        None
    }

    fn diagnose(&mut self, path: String, message: String) -> Result<()> {
        if self.diagnostics.len() >= MAX_DIAGNOSTICS {
            return Err(epub_error("EPUB diagnostics exceed the configured limit"));
        }
        self.diagnostics.push(EpubDiagnostic { path, message });
        Ok(())
    }
}

fn render_styles(styles: &CT_Styles) -> Result<CT_Styles> {
    if styles.styles.len() > MAX_STYLE_ITEMS {
        return Err(epub_error("document has too many styles for EPUB export"));
    }
    let mut key_bytes = 0_usize;
    for style in &styles.styles {
        key_bytes = key_bytes
            .checked_add(style.style_id.len())
            .ok_or_else(|| epub_error("style key size overflow during EPUB export"))?;
        if key_bytes > MAX_PROJECTION_KEY_BYTES {
            return Err(epub_error(
                "document style keys exceed the EPUB projection limit",
            ));
        }
        ensure_xml_value("style identifier", &style.style_id)?;
    }

    let mut projected = CT_Styles::new();
    projected.styles.reserve(styles.styles.len());
    for style in &styles.styles {
        let ppr = style.ppr.as_ref().and_then(|source| {
            source
                .outline_lvl
                .map(|outline_lvl| rdocx_oxml::properties::CT_PPr {
                    outline_lvl: Some(outline_lvl),
                    ..Default::default()
                })
        });
        projected.styles.push(CT_Style {
            style_id: style.style_id.clone(),
            style_type: style.style_type,
            name: None,
            based_on: None,
            next_style: None,
            is_default: style.is_default,
            ppr,
            rpr: None,
            table_properties: None,
            table_properties_original: None,
            table_properties_xml: None,
            conditional_table_styles: Vec::new(),
            extra_xml: Vec::new(),
        });
    }
    Ok(projected)
}

fn render_numbering(numbering: Option<&CT_Numbering>) -> Result<Option<CT_Numbering>> {
    let Some(numbering) = numbering else {
        return Ok(None);
    };
    let item_count = numbering
        .nums
        .len()
        .checked_add(numbering.abstract_nums.len())
        .and_then(|count| {
            numbering
                .abstract_nums
                .iter()
                .try_fold(count, |count, item| count.checked_add(item.levels.len()))
        })
        .ok_or_else(|| epub_error("numbering item count overflow during EPUB export"))?;
    if item_count > MAX_NUMBERING_ITEMS {
        return Err(epub_error(
            "document has too many numbering items for EPUB export",
        ));
    }

    let nums = numbering
        .nums
        .iter()
        .map(|item| CT_Num {
            num_id: item.num_id,
            abstract_num_id: item.abstract_num_id,
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        })
        .collect();
    let abstract_nums = numbering
        .abstract_nums
        .iter()
        .map(|item| CT_AbstractNum {
            abstract_num_id: item.abstract_num_id,
            levels: item
                .levels
                .iter()
                .map(|level| CT_Lvl {
                    ilvl: level.ilvl,
                    start: level.start,
                    num_fmt: level.num_fmt.clone(),
                    p_style: None,
                    p_style_raw: None,
                    suffix: level.suffix,
                    lvl_text: None,
                    lvl_jc: level.lvl_jc,
                    ppr: None,
                    rpr: None,
                    extra_xml: Vec::new(),
                    extra_attributes: Vec::new(),
                    ppr_raw: None,
                    rpr_raw: None,
                })
                .collect(),
            nsid: None,
            nsid_raw: None,
            multi_level_type: None,
            tmpl: None,
            tmpl_raw: None,
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        })
        .collect();
    Ok(Some(CT_Numbering {
        abstract_nums,
        nums,
        root_attributes: Vec::new(),
        extra_xml: Vec::new(),
    }))
}

fn style_has_unprojected_visuals(style: &CT_Style) -> bool {
    let projected_ppr = style.ppr.as_ref().and_then(|properties| {
        properties
            .outline_lvl
            .map(|outline_lvl| rdocx_oxml::properties::CT_PPr {
                outline_lvl: Some(outline_lvl),
                ..Default::default()
            })
    });
    style.ppr != projected_ppr
        || style.rpr.is_some()
        || style.based_on.is_some()
        || style.table_properties.is_some()
        || style.table_properties_original.is_some()
        || style.table_properties_xml.is_some()
        || !style.conditional_table_styles.is_empty()
        || !style.extra_xml.is_empty()
}

fn default_paragraph_style_has_visible_effects(styles: &CT_Styles, paragraph: &CT_P) -> bool {
    let has_text = paragraph.runs.iter().any(|run| {
        run.content.iter().any(|content| match content {
            RunContent::Text(text) | RunContent::DeletedText(text) => !text.text.is_empty(),
            RunContent::Field(field) => !field.cached_result.is_empty(),
            _ => false,
        })
    });
    if styles.doc_defaults.as_ref().is_some_and(|defaults| {
        defaults
            .ppr
            .as_ref()
            .is_some_and(paragraph_properties_have_visible_effects)
            || has_text
                && defaults
                    .rpr
                    .as_ref()
                    .is_some_and(run_properties_have_visible_effects)
    }) {
        return true;
    }
    let Some(mut style) = styles.get_default(StyleType::Paragraph) else {
        return false;
    };
    for _ in 0..=MAX_NESTING_DEPTH {
        if style
            .ppr
            .as_ref()
            .is_some_and(paragraph_properties_have_visible_effects)
            || has_text
                && style
                    .rpr
                    .as_ref()
                    .is_some_and(run_properties_have_visible_effects)
        {
            return true;
        }
        let Some(base) = style
            .based_on
            .as_deref()
            .and_then(|style_id| styles.get_by_id(style_id))
        else {
            return false;
        };
        style = base;
    }
    true
}

fn paragraph_properties_have_visible_effects(properties: &rdocx_oxml::properties::CT_PPr) -> bool {
    properties.style_id.is_some()
        || properties.jc.is_some()
        || properties.space_before.is_some()
        || properties.space_after.is_some()
        || properties.line_spacing.is_some()
        || properties.line_rule.is_some()
        || properties.before_autospacing.is_some()
        || properties.after_autospacing.is_some()
        || properties.ind_left.is_some()
        || properties.ind_right.is_some()
        || properties.ind_first_line.is_some()
        || properties.ind_hanging.is_some()
        || properties.keep_next.is_some()
        || properties.keep_lines.is_some()
        || properties.page_break_before.is_some()
        || properties.widow_control.is_some()
        || properties.suppress_auto_hyphens.is_some()
        || properties.outline_lvl.is_some()
        || properties.borders.is_some()
        || properties.tabs.is_some()
        || properties.shading.is_some()
        || properties
            .rpr
            .as_ref()
            .is_some_and(run_properties_have_visible_effects)
        || properties.num_ilvl.is_some()
        || properties.num_id.is_some()
        || properties.sect_pr.is_some()
}

fn run_properties_have_visible_effects(properties: &rdocx_oxml::properties::CT_RPr) -> bool {
    properties.style_id.is_some()
        || properties.font_ascii.is_some()
        || properties.font_hansi.is_some()
        || properties.font_east_asia.is_some()
        || properties.font_cs.is_some()
        || properties.font_ascii_theme.is_some()
        || properties.font_hansi_theme.is_some()
        || properties.bold.is_some()
        || properties.bold_cs.is_some()
        || properties.italic.is_some()
        || properties.italic_cs.is_some()
        || properties.underline.is_some()
        || properties.strike.is_some()
        || properties.dstrike.is_some()
        || properties.sz.is_some()
        || properties.sz_cs.is_some()
        || properties.color.is_some()
        || properties.color_theme.is_some()
        || properties.highlight.is_some()
        || properties.caps.is_some()
        || properties.small_caps.is_some()
        || properties.vert_align.is_some()
        || properties.spacing.is_some()
        || properties.width_scale.is_some()
        || properties.position.is_some()
        || properties.shading.is_some()
        || properties.vanish.is_some()
}

fn list_definition<'a>(
    paragraph: &CT_P,
    numbering: Option<&'a CT_Numbering>,
) -> Option<&'a CT_Lvl> {
    let properties = paragraph.properties.as_ref()?;
    let num_id = properties.num_id?;
    if num_id == 0 {
        return None;
    }
    let level = properties.num_ilvl.unwrap_or(0);
    let numbering = numbering?;
    let abstract_id = numbering
        .nums
        .iter()
        .find(|item| item.num_id == num_id)?
        .abstract_num_id;
    numbering
        .abstract_nums
        .iter()
        .find(|item| item.abstract_num_id == abstract_id)?
        .levels
        .iter()
        .find(|item| item.ilvl == level)
}

fn projected_paragraph_text(paragraph: &CT_P) -> String {
    let mut text = String::new();
    for run in &paragraph.runs {
        for content in &run.content {
            match content {
                RunContent::Text(value) | RunContent::DeletedText(value) => {
                    text.push_str(&value.text)
                }
                RunContent::Tab => text.push('\t'),
                RunContent::Break(_) => text.push('\n'),
                RunContent::Field(field) => {
                    if let Some(value) = field.projected_text() {
                        text.push_str(value);
                    }
                }
                RunContent::Drawing(_)
                | RunContent::FootnoteRef { .. }
                | RunContent::EndnoteRef { .. }
                | RunContent::CommentReference { .. } => {}
            }
        }
    }
    text
}

fn referenced_drawing_ids(content: &[BodyContent]) -> Vec<&str> {
    fn paragraph<'a>(paragraph: &'a CT_P, ids: &mut Vec<&'a str>) {
        for run in &paragraph.runs {
            for content in &run.content {
                let RunContent::Drawing(drawing) = content else {
                    continue;
                };
                if let Some(id) = drawing
                    .inline
                    .as_ref()
                    .map(|image| image.embed_id.as_str())
                    .or_else(|| drawing.anchor.as_ref().map(|image| image.embed_id.as_str()))
                {
                    ids.push(id);
                }
            }
        }
    }
    fn table<'a>(current: &'a CT_Tbl, ids: &mut Vec<&'a str>) {
        for row in &current.rows {
            for cell in &row.cells {
                for content in &cell.content {
                    match content {
                        CellContent::Paragraph(item) => paragraph(item, ids),
                        CellContent::Table(item) => table(item, ids),
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
    }
    let mut ids = Vec::new();
    for item in content {
        match item {
            BodyContent::Paragraph(item) => paragraph(item, &mut ids),
            BodyContent::Table(item) => table(item, &mut ids),
            BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {}
        }
    }
    ids
}

fn referenced_hyperlink_ids(content: &[BodyContent]) -> Vec<&str> {
    fn paragraph<'a>(paragraph: &'a CT_P, ids: &mut Vec<&'a str>) {
        ids.extend(
            paragraph
                .hyperlinks
                .iter()
                .filter_map(|hyperlink| hyperlink.rel_id.as_deref()),
        );
    }
    fn table<'a>(current: &'a CT_Tbl, ids: &mut Vec<&'a str>) {
        for row in &current.rows {
            for cell in &row.cells {
                for content in &cell.content {
                    match content {
                        CellContent::Paragraph(item) => paragraph(item, ids),
                        CellContent::Table(item) => table(item, ids),
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
    }
    let mut ids = Vec::new();
    for item in content {
        match item {
            BodyContent::Paragraph(item) => paragraph(item, &mut ids),
            BodyContent::Table(item) => table(item, &mut ids),
            BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {}
        }
    }
    ids
}

fn heading_level(paragraph: &CT_P) -> Option<u32> {
    let style = paragraph.properties.as_ref()?.style_id.as_deref()?;
    let level = style.strip_prefix("Heading")?.parse::<u32>().ok()?;
    (1..=9).contains(&level).then_some(level)
}

fn projected_heading_level(paragraph: &CT_P, styles: &CT_Styles) -> Option<u32> {
    let properties = paragraph.properties.as_ref()?;
    if let Some(level) = properties.outline_lvl {
        return level.checked_add(1);
    }
    let style_id = properties.style_id.as_deref()?;
    if let Some(level) = style_id
        .strip_prefix("Heading")
        .and_then(|level| level.parse::<u32>().ok())
        .filter(|level| (1..=9).contains(level))
    {
        return Some(level);
    }
    styles
        .get_by_id(style_id)
        .and_then(|style| style.ppr.as_ref())
        .and_then(|properties| properties.outline_lvl)
        .and_then(|level| level.checked_add(1))
}

fn projects_as_heading(paragraph: &CT_P, styles: &CT_Styles) -> bool {
    let Some(properties) = &paragraph.properties else {
        return false;
    };
    if properties.outline_lvl.is_some_and(|level| level <= 8) {
        return true;
    }
    let Some(style_id) = properties.style_id.as_deref() else {
        return false;
    };
    if style_id
        .strip_prefix("Heading")
        .and_then(|level| level.parse::<u32>().ok())
        .is_some_and(|level| (1..=9).contains(&level))
    {
        return true;
    }
    styles
        .get_by_id(style_id)
        .and_then(|style| style.ppr.as_ref())
        .and_then(|properties| properties.outline_lvl)
        .is_some_and(|level| level <= 8)
}

fn safe_absolute_url(url: &str) -> Option<&str> {
    const SAFE_SCHEMES: &[&str] = &["http", "https", "mailto", "tel", "ftp", "ftps"];
    if url.is_empty()
        || url != url.trim()
        || !url.is_ascii()
        || url.bytes().filter(|byte| *byte == b'#').count() > 1
    {
        return None;
    }
    let separator = url.find(':')?;
    let scheme = &url[..separator];
    if scheme.is_empty()
        || !scheme.as_bytes()[0].is_ascii_alphabetic()
        || !scheme
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
        || !SAFE_SCHEMES
            .iter()
            .any(|allowed| scheme.eq_ignore_ascii_case(allowed))
    {
        return None;
    }
    let remainder = &url[separator + 1..];
    if remainder.is_empty() {
        return None;
    }
    let bytes = url.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'%' {
            if !bytes
                .get(index + 1..index + 3)
                .is_some_and(|hex| hex.iter().all(u8::is_ascii_hexdigit))
            {
                return None;
            }
            index += 3;
            continue;
        }
        if !(byte.is_ascii_alphanumeric()
            || matches!(
                byte,
                b'-' | b'.'
                    | b'_'
                    | b'~'
                    | b':'
                    | b'/'
                    | b'?'
                    | b'#'
                    | b'['
                    | b']'
                    | b'@'
                    | b'!'
                    | b'$'
                    | b'&'
                    | b'\''
                    | b'('
                    | b')'
                    | b'*'
                    | b'+'
                    | b','
                    | b';'
                    | b'='
            ))
        {
            return None;
        }
        index += 1;
    }
    if matches!(
        scheme.to_ascii_lowercase().as_str(),
        "http" | "https" | "ftp" | "ftps"
    ) {
        let authority_and_path = remainder.strip_prefix("//")?;
        let authority = authority_and_path
            .split(['/', '?', '#'])
            .next()
            .unwrap_or_default();
        if authority_and_path[authority.len()..].contains(['[', ']']) {
            return None;
        }
        if authority.is_empty()
            || authority.starts_with('.')
            || authority.ends_with('.')
            || authority.contains("..")
            || authority.bytes().filter(|byte| *byte == b'@').count() > 1
        {
            return None;
        }
        let (userinfo, host_port) = authority
            .rsplit_once('@')
            .map_or((None, authority), |(userinfo, host)| (Some(userinfo), host));
        if userinfo.is_some_and(|userinfo| {
            userinfo.is_empty()
                || !userinfo.bytes().all(|byte| {
                    byte.is_ascii_alphanumeric()
                        || matches!(
                            byte,
                            b'-' | b'.'
                                | b'_'
                                | b'~'
                                | b'!'
                                | b'$'
                                | b'&'
                                | b'\''
                                | b'('
                                | b')'
                                | b'*'
                                | b'+'
                                | b','
                                | b';'
                                | b'='
                                | b':'
                                | b'%'
                        )
                })
        }) {
            return None;
        }
        if let Some(host) = host_port.strip_prefix('[') {
            let close = host.find(']')?;
            let literal = &host[..close];
            if !valid_ip_literal(literal) || host[close + 1..].contains(']') {
                return None;
            }
            let suffix = &host[close + 1..];
            if !suffix.is_empty()
                && !suffix.strip_prefix(':').is_some_and(|port| {
                    !port.is_empty() && port.bytes().all(|b| b.is_ascii_digit())
                })
            {
                return None;
            }
        } else {
            if host_port.contains(['[', ']']) || host_port.matches(':').count() > 1 {
                return None;
            }
            let (host, port) = host_port
                .rsplit_once(':')
                .map_or((host_port, None), |(host, port)| (host, Some(port)));
            if host.is_empty()
                || port.is_some_and(|port| {
                    port.is_empty() || !port.bytes().all(|byte| byte.is_ascii_digit())
                })
            {
                return None;
            }
        }
    } else if remainder.contains(['[', ']']) {
        return None;
    }
    Some(url)
}

fn valid_ip_literal(literal: &str) -> bool {
    if literal.parse::<std::net::Ipv6Addr>().is_ok() {
        return true;
    }
    let Some(versioned) = literal
        .strip_prefix('v')
        .or_else(|| literal.strip_prefix('V'))
    else {
        return false;
    };
    let Some((version, address)) = versioned.split_once('.') else {
        return false;
    };
    !version.is_empty()
        && version.bytes().all(|byte| byte.is_ascii_hexdigit())
        && !address.is_empty()
        && address.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || matches!(
                    byte,
                    b'-' | b'.'
                        | b'_'
                        | b'~'
                        | b'!'
                        | b'$'
                        | b'&'
                        | b'\''
                        | b'('
                        | b')'
                        | b'*'
                        | b'+'
                        | b','
                        | b';'
                        | b'='
                        | b':'
                )
        })
}

fn ensure_xml_value(location: &str, value: &str) -> Result<()> {
    if value.chars().all(is_xml_10_character) {
        return Ok(());
    }
    Err(epub_error(format!(
        "{location} contains a character forbidden by XML 1.0"
    )))
}

fn ensure_xml_10(location: &str, xml: &str) -> Result<()> {
    ensure_xml_value(location, xml)
}

fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{0009}' | '\u{000A}' | '\u{000D}')
        || ('\u{0020}'..='\u{D7FF}').contains(&character)
        || ('\u{E000}'..='\u{FFFD}').contains(&character)
        || ('\u{10000}'..='\u{10FFFF}').contains(&character)
}

#[derive(Clone, Copy)]
struct RawMarkerKind {
    kind: u8,
}

fn raw_marker_kind(raw: &[u8]) -> Option<RawMarkerKind> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let name = loop {
        match reader.read_event_into(&mut buffer).ok()? {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element) => {
                break element.name().as_ref().to_vec();
            }
            quick_xml::events::Event::Eof => return None,
            _ => buffer.clear(),
        }
    };
    let local = name
        .iter()
        .rposition(|byte| *byte == b':')
        .map_or(name.as_slice(), |separator| &name[separator + 1..]);
    let kind = match local {
        b"commentRangeStart" => Some(0),
        b"commentRangeEnd" => Some(1),
        b"bookmarkStart" => Some(2),
        b"bookmarkEnd" => Some(3),
        b"ins" => Some(10),
        b"del" => Some(11),
        b"moveFrom" => Some(12),
        b"moveTo" => Some(13),
        _ => None,
    }?;
    Some(RawMarkerKind { kind })
}

fn root_heading_indexes(headings: &[Heading]) -> Vec<usize> {
    let mut roots = Vec::new();
    let mut stack = Vec::<u32>::new();
    for (index, heading) in headings.iter().enumerate() {
        while stack.last().is_some_and(|level| *level >= heading.level) {
            stack.pop();
        }
        if stack.is_empty() {
            roots.push(index);
        }
        stack.push(heading.level);
    }
    roots
}

fn build_nav_tree(headings: &[Heading]) -> Vec<NavNode> {
    fn parse(headings: &[Heading], next: &mut usize, parent_level: u32) -> Vec<NavNode> {
        let mut nodes = Vec::new();
        while *next < headings.len() && headings[*next].level > parent_level {
            let index = *next;
            let level = headings[index].level;
            *next += 1;
            let children = parse(headings, next, level);
            nodes.push(NavNode {
                heading_index: index,
                children,
            });
        }
        nodes
    }
    let mut next = 0;
    parse(headings, &mut next, 0)
}

fn emit_spine_fragment(
    content: &[BodyContent],
    input: &mut rdocx_html::HtmlInput,
    media: &[MediaItem],
    heading_anchors: &HashMap<usize, String>,
    list_counters: &mut HashMap<(u32, u32), u32>,
    start: usize,
    end: usize,
) -> Result<String> {
    let mut output = String::new();
    let mut index = start;
    while index < end {
        if matches!(
            content[index],
            BodyContent::ContentControl(_) | BodyContent::RawXml(_)
        ) {
            index += 1;
            continue;
        }
        if let BodyContent::Paragraph(paragraph) = &content[index]
            && let Some(list) = detect_list(paragraph, input.numbering.as_ref())
        {
            emit_list_level(
                &mut output,
                content,
                input,
                media,
                heading_anchors,
                list_counters,
                &mut index,
                end,
                list,
                0,
            )?;
            continue;
        }
        let anchor = heading_anchors.get(&index).map(String::as_str);
        let fragment = render_source_block(input, &content[index], media, anchor, false)?;
        push_bounded_xhtml(&mut output, &fragment)?;
        index += 1;
    }
    Ok(output)
}

#[allow(clippy::too_many_arguments)]
fn emit_list_level(
    output: &mut String,
    content: &[BodyContent],
    input: &mut rdocx_html::HtmlInput,
    media: &[MediaItem],
    heading_anchors: &HashMap<usize, String>,
    list_counters: &mut HashMap<(u32, u32), u32>,
    index: &mut usize,
    end: usize,
    list: ListInfo,
    depth: usize,
) -> Result<()> {
    if depth >= MAX_NESTING_DEPTH {
        return Err(epub_error("list nesting exceeds the EPUB depth limit"));
    }
    let tag = list.kind.tag();
    let counter_key = (list.num_id, list.level);
    let effective_start = list_counters
        .get(&counter_key)
        .copied()
        .unwrap_or(list.start);
    let mut opening = format!("<{tag}");
    if list.kind == ListKind::None {
        opening.push_str(" class=\"no-marker\"");
    }
    if list.kind == ListKind::Ordered && effective_start != 1 {
        write!(opening, " start=\"{effective_start}\"")
            .map_err(|_| epub_error("failed to format EPUB list start"))?;
    }
    if let Some(marker_style) = list.marker_style {
        write!(opening, " style=\"list-style-type:{marker_style}\"")
            .map_err(|_| epub_error("failed to format EPUB list marker style"))?;
    }
    opening.push_str(">\n");
    push_bounded_xhtml(output, &opening)?;
    let mut item_open = false;

    while *index < end {
        let BodyContent::Paragraph(paragraph) = &content[*index] else {
            break;
        };
        let Some(item) = detect_list(paragraph, input.numbering.as_ref()) else {
            break;
        };
        if item.level < list.level
            || (item.level == list.level && (item.num_id != list.num_id || item.kind != list.kind))
        {
            break;
        }
        if item.level > list.level {
            if !item_open || item.num_id != list.num_id {
                break;
            }
            emit_list_level(
                output,
                content,
                input,
                media,
                heading_anchors,
                list_counters,
                index,
                end,
                item,
                depth + 1,
            )?;
            continue;
        }

        if item_open {
            push_bounded_xhtml(output, "</li>\n")?;
        }
        list_counters.retain(|(num_id, level), _| *num_id != list.num_id || *level <= list.level);
        let anchor = heading_anchors.get(index).map(String::as_str);
        let preserve_heading = projects_as_heading(paragraph, &input.styles);
        push_bounded_xhtml(output, "<li>")?;
        let fragment = render_source_block(
            input,
            &content[*index],
            media,
            if preserve_heading { anchor } else { None },
            preserve_heading,
        )?;
        if preserve_heading {
            push_bounded_xhtml(output, &fragment)?;
        } else {
            push_bounded_xhtml(output, list_item_inner(&fragment)?)?;
        }
        item_open = true;
        if list.kind == ListKind::Ordered {
            let next = list_counters.entry(counter_key).or_insert(effective_start);
            *next = next
                .checked_add(1)
                .ok_or_else(|| epub_error("list counter overflow during EPUB export"))?;
        }
        *index += 1;
    }

    if item_open {
        push_bounded_xhtml(output, "</li>\n")?;
    }
    push_bounded_xhtml(output, &format!("</{tag}>\n"))?;
    Ok(())
}

fn render_source_block(
    input: &mut rdocx_html::HtmlInput,
    content: &BodyContent,
    media: &[MediaItem],
    anchor: Option<&str>,
    suppress_numbering: bool,
) -> Result<String> {
    let image_occurrences = supported_image_occurrences(content, input);
    let mut content = render_body_projection(content)
        .ok_or_else(|| epub_error("unsupported source block reached EPUB projection"))?;
    if let BodyContent::Paragraph(paragraph) = &mut content
        && heading_level(paragraph).is_some_and(|level| level > 6)
    {
        paragraph.properties.get_or_insert_default().style_id = Some("Heading6".to_owned());
    }
    if suppress_numbering && let BodyContent::Paragraph(paragraph) = &mut content {
        let properties = paragraph.properties.get_or_insert_default();
        properties.num_id = None;
        properties.num_ilvl = None;
    }
    input.document.body.content = vec![content];
    input.document.body.sect_pr = None;
    let fragment = rdocx_html::to_html_fragment(
        input,
        &rdocx_html::HtmlOptions {
            inline_images: true,
        },
    );
    let fragment = replace_media_sources(&fragment, input, media, &image_occurrences)?;
    let fragment = lift_page_breaks(&fragment)?;
    let mut fragment = normalize_xhtml(fragment);
    if let Some(anchor) = anchor {
        inject_first_source_anchor(&mut fragment, anchor)?;
    }
    if fragment.len() > MAX_EPUB_BYTES / 2 {
        return Err(epub_error(
            "generated source block exceeds the EPUB intermediate limit",
        ));
    }
    Ok(fragment)
}

fn render_body_projection(content: &BodyContent) -> Option<BodyContent> {
    match content {
        BodyContent::Paragraph(paragraph) => Some(BodyContent::Paragraph(
            render_paragraph_projection(paragraph),
        )),
        BodyContent::Table(table) => Some(BodyContent::Table(render_table_projection(table))),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => None,
    }
}

fn render_paragraph_projection(paragraph: &CT_P) -> CT_P {
    let mut properties = paragraph
        .properties
        .as_ref()
        .map(render_paragraph_properties);
    if properties
        .as_ref()
        .and_then(|properties| properties.outline_lvl)
        .is_some_and(|level| level >= 6)
    {
        let properties = properties.get_or_insert_default();
        properties.outline_lvl = Some(5);
        properties.style_id = Some("Heading6".to_owned());
    }
    CT_P {
        properties,
        runs: paragraph.runs.iter().map(render_run_projection).collect(),
        hyperlinks: paragraph
            .hyperlinks
            .iter()
            .map(|hyperlink| HyperlinkSpan {
                rel_id: hyperlink.rel_id.clone(),
                anchor: hyperlink.anchor.clone(),
                tooltip: hyperlink.tooltip.clone(),
                doc_location: hyperlink.doc_location.clone(),
                run_start: hyperlink.run_start,
                run_end: hyperlink.run_end,
                extra_attributes: Vec::new(),
                extra_xml: Vec::new(),
                preserved_raw_before: None,
            })
            .collect(),
        comment_ranges: Vec::new(),
        bookmark_markers: Vec::new(),
        extra_xml: Vec::new(),
        content_controls: Vec::new(),
        revisions: Vec::new(),
        equations: Vec::new(),
    }
}

fn render_paragraph_properties(
    source: &rdocx_oxml::properties::CT_PPr,
) -> rdocx_oxml::properties::CT_PPr {
    rdocx_oxml::properties::CT_PPr {
        style_id: source.style_id.clone(),
        jc: source.jc,
        space_before: source.space_before,
        space_after: source.space_after,
        line_spacing: source.line_spacing,
        line_rule: source.line_rule.clone(),
        ind_left: source.ind_left,
        ind_right: source.ind_right,
        ind_first_line: source.ind_first_line,
        outline_lvl: source.outline_lvl,
        shading: source.shading.as_ref().map(render_shading_projection),
        num_ilvl: source.num_ilvl,
        num_id: source.num_id,
        ..Default::default()
    }
}

fn render_run_projection(run: &CT_R) -> CT_R {
    CT_R {
        properties: run.properties.as_ref().map(render_run_properties),
        content: run.content.iter().map(render_run_content).collect(),
        extra_xml: Vec::new(),
        extra_xml_positions: Vec::new(),
        alt_drawings: Vec::new(),
    }
}

fn render_run_properties(
    source: &rdocx_oxml::properties::CT_RPr,
) -> rdocx_oxml::properties::CT_RPr {
    rdocx_oxml::properties::CT_RPr {
        font_ascii: source.font_ascii.clone(),
        bold: source.bold,
        italic: source.italic,
        underline: source.underline.and_then(|underline| match underline {
            rdocx_oxml::shared::ST_Underline::None => None,
            _ => Some(rdocx_oxml::shared::ST_Underline::Single),
        }),
        strike: source.strike,
        dstrike: source.dstrike,
        sz: source.sz,
        color: source.color.clone(),
        caps: source.caps,
        small_caps: source.small_caps,
        vert_align: source.vert_align.clone(),
        spacing: source.spacing,
        shading: source.shading.as_ref().map(render_shading_projection),
        ..Default::default()
    }
}

fn render_shading_projection(
    source: &rdocx_oxml::properties::CT_Shd,
) -> rdocx_oxml::properties::CT_Shd {
    rdocx_oxml::properties::CT_Shd {
        val: "clear".to_owned(),
        color: None,
        fill: source.fill.clone(),
    }
}

fn render_run_content(content: &RunContent) -> RunContent {
    match content {
        RunContent::Text(text) => RunContent::Text(text.clone()),
        RunContent::DeletedText(text) => RunContent::DeletedText(text.clone()),
        RunContent::Tab => RunContent::Tab,
        RunContent::Break(kind) => RunContent::Break(*kind),
        RunContent::Drawing(drawing) => RunContent::Drawing(render_drawing_projection(drawing)),
        RunContent::Field(field) => {
            let mut projected = Field::new("", &field.cached_result);
            projected.dirty = field.dirty;
            RunContent::Field(projected)
        }
        RunContent::FootnoteRef { id } => RunContent::FootnoteRef { id: *id },
        RunContent::EndnoteRef { id } => RunContent::EndnoteRef { id: *id },
        RunContent::CommentReference { id, .. } => RunContent::CommentReference {
            id: *id,
            raw_before: 0,
        },
    }
}

fn render_drawing_projection(drawing: &CT_Drawing) -> CT_Drawing {
    let source = drawing.inline.as_ref().map(|inline| {
        (
            inline.extent_cx,
            inline.extent_cy,
            inline.embed_id.as_str(),
            inline.chart_rel_id.as_ref(),
            inline.description.as_ref(),
        )
    });
    let source = source.or_else(|| {
        drawing.anchor.as_ref().map(|anchor| {
            (
                anchor.extent_cx,
                anchor.extent_cy,
                anchor.embed_id.as_str(),
                anchor.chart_rel_id.as_ref(),
                anchor.description.as_ref(),
            )
        })
    });
    let Some((extent_cx, extent_cy, embed_id, chart_rel_id, description)) = source else {
        return CT_Drawing {
            inline: None,
            anchor: None,
        };
    };
    CT_Drawing::inline(CT_Inline {
        extent_cx,
        extent_cy,
        embed_id: embed_id.to_owned(),
        link_id: None,
        chart_rel_id: chart_rel_id.cloned(),
        description: description.cloned(),
        name: None,
        raw_xml: None,
    })
}

fn render_table_projection(table: &CT_Tbl) -> CT_Tbl {
    CT_Tbl {
        properties: table
            .properties
            .as_ref()
            .map(|source| rdocx_oxml::table::CT_TblPr {
                jc: source.jc,
                borders: source.borders.clone(),
                ..Default::default()
            }),
        grid: None,
        rows: table
            .rows
            .iter()
            .map(|row| CT_Row {
                table_property_exception: None,
                properties: None,
                cells: row
                    .cells
                    .iter()
                    .map(|cell| CT_Tc {
                        properties: cell.properties.as_ref().map(|source| {
                            rdocx_oxml::table::CT_TcPr {
                                grid_span: source.grid_span,
                                v_merge: source.v_merge,
                                borders: source.borders.clone(),
                                shading: source.shading.as_ref().map(render_shading_projection),
                                v_align: source.v_align,
                                ..Default::default()
                            }
                        }),
                        content: cell
                            .content
                            .iter()
                            .filter_map(|content| match content {
                                CellContent::Paragraph(paragraph) => Some(CellContent::Paragraph(
                                    render_paragraph_projection(paragraph),
                                )),
                                CellContent::Table(table) => {
                                    Some(CellContent::Table(render_table_projection(table)))
                                }
                                CellContent::ContentControl(_) => None,
                            })
                            .collect(),
                        extra_xml: Vec::new(),
                    })
                    .collect(),
                extra_xml: Vec::new(),
                content_controls: Vec::new(),
            })
            .collect(),
        extra_xml: Vec::new(),
        content_controls: Vec::new(),
    }
}

fn inject_first_source_anchor(fragment: &mut String, anchor: &str) -> Result<()> {
    let Some(start) = fragment.find('<') else {
        return Err(epub_error("heading projection produced no source element"));
    };
    let Some(relative_end) = fragment[start..].find([' ', '>']) else {
        return Err(epub_error(
            "heading projection produced a malformed source element",
        ));
    };
    let insertion = start + relative_end;
    fragment.insert_str(insertion, &format!(" id=\"{anchor}\""));
    Ok(())
}

fn list_item_inner(fragment: &str) -> Result<&str> {
    let start = fragment
        .find("<li>")
        .map(|start| start + "<li>".len())
        .ok_or_else(|| epub_error("list paragraph projection produced no list item"))?;
    let end = fragment[start..]
        .find("</li>")
        .map(|end| start + end)
        .ok_or_else(|| epub_error("list paragraph projection left its list item open"))?;
    Ok(&fragment[start..end])
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ListKind {
    Ordered,
    Unordered,
    None,
}

impl ListKind {
    fn tag(self) -> &'static str {
        if self == Self::Ordered { "ol" } else { "ul" }
    }
}

#[derive(Clone, Copy)]
struct ListInfo {
    num_id: u32,
    level: u32,
    kind: ListKind,
    start: u32,
    marker_style: Option<&'static str>,
}

fn detect_list(paragraph: &CT_P, numbering: Option<&CT_Numbering>) -> Option<ListInfo> {
    let properties = paragraph.properties.as_ref()?;
    let num_id = properties.num_id?;
    let level = properties.num_ilvl.unwrap_or(0);
    if num_id == 0 {
        return None;
    }
    let numbering = numbering?;
    let abstract_id = numbering
        .nums
        .iter()
        .find(|numbering| numbering.num_id == num_id)?
        .abstract_num_id;
    let abstract_numbering = numbering
        .abstract_nums
        .iter()
        .find(|numbering| numbering.abstract_num_id == abstract_id)?;
    let definition = abstract_numbering
        .levels
        .iter()
        .find(|definition| definition.ilvl == level)?;
    if matches!(definition.num_fmt.as_ref(), Some(ST_NumberFormat::Other(_))) {
        return None;
    }
    let kind = match definition.num_fmt {
        Some(ST_NumberFormat::Bullet) => ListKind::Unordered,
        Some(ST_NumberFormat::None) => ListKind::None,
        _ => ListKind::Ordered,
    };
    let marker_style = match definition.num_fmt {
        Some(ST_NumberFormat::UpperRoman) => Some("upper-roman"),
        Some(ST_NumberFormat::LowerRoman) => Some("lower-roman"),
        Some(ST_NumberFormat::UpperLetter) => Some("upper-alpha"),
        Some(ST_NumberFormat::LowerLetter) => Some("lower-alpha"),
        _ => None,
    };
    Some(ListInfo {
        num_id,
        level,
        kind,
        start: definition.start.unwrap_or(1),
        marker_style,
    })
}

fn push_bounded_xhtml(output: &mut String, value: &str) -> Result<()> {
    let next = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| epub_error("XHTML size overflow during EPUB export"))?;
    if next > MAX_EPUB_BYTES / 2 {
        return Err(epub_error(
            "generated XHTML exceeds the EPUB intermediate limit",
        ));
    }
    output.push_str(value);
    Ok(())
}

struct ImageOccurrence {
    relationship_id: String,
    description: String,
}

fn supported_image_occurrences(
    content: &BodyContent,
    input: &rdocx_html::HtmlInput,
) -> Vec<ImageOccurrence> {
    fn paragraph(
        paragraph: &CT_P,
        input: &rdocx_html::HtmlInput,
        occurrences: &mut Vec<ImageOccurrence>,
    ) {
        for run in &paragraph.runs {
            for content in &run.content {
                let RunContent::Drawing(drawing) = content else {
                    continue;
                };
                let source = drawing
                    .inline
                    .as_ref()
                    .map(|image| (image.embed_id.as_str(), image.description.as_deref()))
                    .or_else(|| {
                        drawing
                            .anchor
                            .as_ref()
                            .map(|image| (image.embed_id.as_str(), image.description.as_deref()))
                    });
                let Some((relationship_id, description)) = source else {
                    continue;
                };
                if input.images.contains_key(relationship_id) {
                    occurrences.push(ImageOccurrence {
                        relationship_id: relationship_id.to_owned(),
                        description: description.unwrap_or_default().to_owned(),
                    });
                }
            }
        }
    }
    fn table(
        current: &CT_Tbl,
        input: &rdocx_html::HtmlInput,
        occurrences: &mut Vec<ImageOccurrence>,
    ) {
        for row in &current.rows {
            for cell in &row.cells {
                for content in &cell.content {
                    match content {
                        CellContent::Paragraph(item) => paragraph(item, input, occurrences),
                        CellContent::Table(item) => table(item, input, occurrences),
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
    }

    let mut occurrences = Vec::new();
    match content {
        BodyContent::Paragraph(item) => paragraph(item, input, &mut occurrences),
        BodyContent::Table(item) => table(item, input, &mut occurrences),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {}
    }
    occurrences
}

fn replace_media_sources(
    fragment: &str,
    input: &rdocx_html::HtmlInput,
    media: &[MediaItem],
    occurrences: &[ImageOccurrence],
) -> Result<String> {
    let mut output = String::with_capacity(fragment.len());
    let mut remaining = fragment;
    let mut occurrence_index = 0;
    while let Some(start) = remaining.find("<img ") {
        output.push_str(&remaining[..start]);
        let tag = &remaining[start..];
        let end = tag
            .find('>')
            .ok_or_else(|| epub_error("image projection produced an unterminated tag"))?;
        let tag = &tag[..=end];
        let value_start = tag
            .find(" src=\"")
            .map(|offset| offset + " src=\"".len())
            .ok_or_else(|| epub_error("image projection produced no source attribute"))?;
        let value_end = tag[value_start..]
            .find('"')
            .map(|offset| value_start + offset)
            .ok_or_else(|| epub_error("image projection produced an open source attribute"))?;
        let source = &tag[value_start..value_end];
        let relationship_id = input.images.iter().find_map(|(relationship_id, image)| {
            let marker = format!(
                "data:{};base64,{}",
                image.content_type,
                base64_encode(&image.data)
            );
            (marker == source).then_some(relationship_id)
        });
        let relationship_id = relationship_id
            .ok_or_else(|| epub_error("image projection produced an unknown source marker"))?;
        let occurrence = occurrences
            .get(occurrence_index)
            .filter(|occurrence| occurrence.relationship_id == *relationship_id)
            .ok_or_else(|| epub_error("image projection lost source occurrence correlation"))?;
        occurrence_index += 1;
        let item = media
            .iter()
            .find(|item| item.relationship_id == *relationship_id)
            .ok_or_else(|| epub_error("image projection lost its packaged media item"))?;
        let attribute_start = value_start - "src=\"".len();
        output.push_str(&tag[..attribute_start]);
        write!(
            output,
            "alt=\"{}\" src=\"",
            escape_xml(&occurrence.description)
        )
        .map_err(|_| epub_error("failed to format EPUB image alternative text"))?;
        output.push_str(&item.href);
        output.push_str(&tag[value_end..]);
        remaining = &remaining[start + tag.len()..];
    }
    output.push_str(remaining);
    if occurrence_index != occurrences.len() {
        return Err(epub_error(
            "image projection omitted a supported source occurrence",
        ));
    }
    Ok(output)
}

#[derive(Clone)]
struct OpenXhtmlTag {
    name: String,
    opening: String,
}

fn lift_page_breaks(fragment: &str) -> Result<String> {
    let mut output = String::with_capacity(fragment.len());
    let mut stack = Vec::<OpenXhtmlTag>::new();
    let mut remaining = fragment;
    while let Some(start) = remaining.find('<') {
        output.push_str(&remaining[..start]);
        let tag_and_rest = &remaining[start..];
        let end = tag_and_rest
            .find('>')
            .ok_or_else(|| epub_error("HTML projection produced an unterminated tag"))?;
        let tag = &tag_and_rest[..=end];
        let name = xhtml_tag_name(tag)
            .ok_or_else(|| epub_error("HTML projection produced a malformed tag"))?;
        let closing = tag.as_bytes().get(1) == Some(&b'/');
        let void = tag.ends_with("/>") || matches!(name, "br" | "hr" | "img" | "input");

        if name == "hr" && !closing {
            let keep = stack
                .iter()
                .rposition(|entry| flow_container(&entry.name))
                .map_or(0, |index| index + 1);
            for entry in stack[keep..].iter().rev() {
                write!(output, "</{}>", entry.name)
                    .map_err(|_| epub_error("failed to close XHTML around a page break"))?;
            }
            output.push_str(tag);
            for entry in &stack[keep..] {
                output.push_str(&opening_without_id(&entry.opening));
            }
        } else {
            output.push_str(tag);
            if closing {
                if stack.last().is_none_or(|entry| entry.name != name) {
                    return Err(epub_error("HTML projection produced mismatched XHTML tags"));
                }
                stack.pop();
            } else if !void {
                stack.push(OpenXhtmlTag {
                    name: name.to_owned(),
                    opening: tag.to_owned(),
                });
            }
        }
        remaining = &tag_and_rest[end + 1..];
    }
    output.push_str(remaining);
    if !stack.is_empty() {
        return Err(epub_error("HTML projection left XHTML tags open"));
    }
    Ok(output)
}

fn xhtml_tag_name(tag: &str) -> Option<&str> {
    let content = tag.strip_prefix('<')?;
    let content = content.strip_prefix('/').unwrap_or(content);
    let end = content
        .find(|character: char| character.is_ascii_whitespace() || matches!(character, '/' | '>'))
        .unwrap_or(content.len());
    (end > 0).then_some(&content[..end])
}

fn flow_container(name: &str) -> bool {
    matches!(
        name,
        "article"
            | "aside"
            | "blockquote"
            | "caption"
            | "dd"
            | "div"
            | "dt"
            | "figcaption"
            | "figure"
            | "footer"
            | "header"
            | "li"
            | "main"
            | "nav"
            | "section"
            | "td"
            | "th"
    )
}

fn opening_without_id(opening: &str) -> String {
    let Some(start) = opening.find(" id=\"") else {
        return opening.to_owned();
    };
    let value_start = start + " id=\"".len();
    let Some(end) = opening[value_start..].find('"') else {
        return opening.to_owned();
    };
    let mut reopened = String::with_capacity(opening.len());
    reopened.push_str(&opening[..start]);
    reopened.push_str(&opening[value_start + end + 1..]);
    reopened
}

fn normalize_xhtml(fragment: String) -> String {
    fragment
        .replace("&emsp;", "&#8195;")
        .replace("<br>", "<br/>")
        .replace("<hr>", "<hr/>")
        .replace(" style=\"max-width:100%\">", " style=\"max-width:100%\"/>")
}

fn xhtml_document(title: &str, fragment: &str) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<html xmlns=\"http://www.w3.org/1999/xhtml\" xml:lang=\"en\" lang=\"en\">\n<head><title>{}</title><link rel=\"stylesheet\" type=\"text/css\" href=\"styles.css\"/></head>\n<body>\n{fragment}</body>\n</html>\n",
        escape_xml(title)
    )
}

fn navigation_document(
    title: &str,
    headings: &[Heading],
    tree: &[NavNode],
    spine: &[SpineItem],
) -> String {
    let mut list = String::new();
    if tree.is_empty() {
        let href = spine
            .first()
            .map(|item| item.href.as_str())
            .unwrap_or("document.xhtml");
        let _ = write!(
            list,
            "<ol><li><a href=\"{}\">{}</a></li></ol>",
            escape_xml(href),
            escape_xml(title)
        );
    } else {
        write_nav_nodes(&mut list, headings, tree);
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\" xml:lang=\"en\" lang=\"en\">\n<head><title>{}</title><link rel=\"stylesheet\" type=\"text/css\" href=\"styles.css\"/></head>\n<body><nav epub:type=\"toc\" id=\"toc\"><h1>Contents</h1>{list}</nav></body>\n</html>\n",
        escape_xml(title)
    )
}

fn write_nav_nodes(output: &mut String, headings: &[Heading], nodes: &[NavNode]) {
    output.push_str("<ol>");
    for node in nodes {
        let heading = &headings[node.heading_index];
        let _ = write!(
            output,
            "<li><a href=\"{}\">{}</a>",
            escape_xml(&heading.href),
            escape_xml(&heading.text)
        );
        if !node.children.is_empty() {
            write_nav_nodes(output, headings, &node.children);
        }
        output.push_str("</li>");
    }
    output.push_str("</ol>");
}

fn package_document(
    title: &str,
    author: &str,
    identifier: &str,
    spine: &[SpineItem],
    media: &[MediaItem],
) -> String {
    let mut manifest = String::from(
        "<item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>\n<item id=\"styles\" href=\"styles.css\" media-type=\"text/css\"/>\n",
    );
    let mut spine_xml = String::new();
    for item in spine {
        let _ = writeln!(
            manifest,
            "<item id=\"{}\" href=\"{}\" media-type=\"application/xhtml+xml\"/>",
            escape_xml(&item.id),
            escape_xml(&item.href)
        );
        let _ = writeln!(spine_xml, "<itemref idref=\"{}\"/>", escape_xml(&item.id));
    }
    let mut emitted_hrefs = Vec::<&str>::new();
    for item in media {
        if emitted_hrefs.contains(&item.href.as_str()) {
            continue;
        }
        emitted_hrefs.push(&item.href);
        let id = format!("image-{:03}", emitted_hrefs.len());
        let _ = writeln!(
            manifest,
            "<item id=\"{id}\" href=\"{}\" media-type=\"{}\"/>",
            escape_xml(&item.href),
            escape_xml(&item.media_type)
        );
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"publication-id\" xml:lang=\"en\">\n<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\">\n<dc:identifier id=\"publication-id\">{}</dc:identifier>\n<dc:title>{}</dc:title>\n<dc:language>en</dc:language>\n<dc:creator>{}</dc:creator>\n<meta property=\"dcterms:modified\">1980-01-01T00:00:00Z</meta>\n</metadata>\n<manifest>\n{manifest}</manifest>\n<spine>\n{spine_xml}</spine>\n</package>\n",
        escape_xml(identifier),
        escape_xml(title),
        escape_xml(author)
    )
}

fn publication_identifier(
    title: &str,
    author: &str,
    spine: &[SpineItem],
    media: &[MediaItem],
) -> String {
    let mut hash = 0xcbf29ce484222325_u64;
    for bytes in [title.as_bytes(), author.as_bytes()] {
        fnv1a(&mut hash, bytes);
    }
    for item in spine {
        fnv1a(&mut hash, item.xhtml.as_bytes());
    }
    let mut media = media.iter().collect::<Vec<_>>();
    media.sort_unstable_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    for item in media {
        fnv1a(&mut hash, item.media_type.as_bytes());
        fnv1a(&mut hash, &item.data);
    }
    format!("urn:rdocx:{hash:016x}")
}

fn fnv1a(hash: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x100000001b3);
    }
}

fn write_archive(
    package: &str,
    nav: &str,
    spine: &[SpineItem],
    media: &[MediaItem],
) -> Result<Vec<u8>> {
    let cursor = BoundedCursor::new(MAX_EPUB_BYTES);
    let mut archive = ZipWriter::new(cursor);
    let stored = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Stored)
        .last_modified_time(zip::DateTime::default());
    let deflated = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .last_modified_time(zip::DateTime::default());

    add_entry(&mut archive, "mimetype", b"application/epub+zip", stored)?;
    add_entry(
        &mut archive,
        "META-INF/container.xml",
        CONTAINER_XML.as_bytes(),
        deflated,
    )?;
    add_entry(
        &mut archive,
        "EPUB/package.opf",
        package.as_bytes(),
        deflated,
    )?;
    add_entry(&mut archive, "EPUB/nav.xhtml", nav.as_bytes(), deflated)?;
    add_entry(
        &mut archive,
        "EPUB/styles.css",
        STYLESHEET.as_bytes(),
        deflated,
    )?;
    for item in spine {
        add_entry(
            &mut archive,
            &format!("EPUB/{}", item.href),
            item.xhtml.as_bytes(),
            deflated,
        )?;
    }
    let mut emitted_hrefs = Vec::<&str>::new();
    for item in media {
        if emitted_hrefs.contains(&item.href.as_str()) {
            continue;
        }
        emitted_hrefs.push(&item.href);
        add_entry(
            &mut archive,
            &format!("EPUB/{}", item.href),
            &item.data,
            deflated,
        )?;
    }
    let cursor = archive.finish().map_err(zip_error)?;
    Ok(cursor.into_inner())
}

fn add_entry(
    archive: &mut ZipWriter<BoundedCursor>,
    name: &str,
    bytes: &[u8],
    options: SimpleFileOptions,
) -> Result<()> {
    archive.start_file(name, options).map_err(zip_error)?;
    archive.write_all(bytes)?;
    Ok(())
}

struct BoundedCursor {
    bytes: Vec<u8>,
    position: u64,
    limit: usize,
}

impl BoundedCursor {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            position: 0,
            limit,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for BoundedCursor {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let start = usize::try_from(self.position).map_err(|_| output_limit_error())?;
        let end = start
            .checked_add(buffer.len())
            .ok_or_else(output_limit_error)?;
        if end > self.limit {
            return Err(output_limit_error());
        }
        if end > self.bytes.len() {
            self.bytes.resize(end, 0);
        }
        self.bytes[start..end].copy_from_slice(buffer);
        self.position = end as u64;
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Seek for BoundedCursor {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        let base = match position {
            SeekFrom::Start(position) => i128::from(position),
            SeekFrom::End(offset) => self.bytes.len() as i128 + i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
        };
        if base < 0 || base > self.limit as i128 {
            return Err(output_limit_error());
        }
        self.position = base as u64;
        Ok(self.position)
    }
}

fn measure_body_content(
    content: &BodyContent,
    depth: usize,
    item_count: &mut usize,
    text_bytes: &mut usize,
    projected_nodes: &mut usize,
    image_occurrences: &mut usize,
) -> Result<()> {
    if depth > MAX_NESTING_DEPTH {
        return Err(epub_error("table nesting exceeds the EPUB depth limit"));
    }
    *item_count = item_count
        .checked_add(1)
        .ok_or_else(|| epub_error("document item count overflow during EPUB export"))?;
    if *item_count > MAX_BODY_ITEMS {
        return Err(epub_error(
            "document has too many body items for EPUB export",
        ));
    }
    add_projected_nodes(projected_nodes, 1)?;
    match content {
        BodyContent::Paragraph(paragraph) => {
            measure_paragraph(paragraph, text_bytes, projected_nodes, image_occurrences)?
        }
        BodyContent::Table(table) => measure_table(
            table,
            depth + 1,
            item_count,
            text_bytes,
            projected_nodes,
            image_occurrences,
        )?,
        BodyContent::ContentControl(_) => {}
        BodyContent::RawXml(raw) => add_source_bytes(text_bytes, raw.len())?,
    }
    Ok(())
}

fn measure_table(
    table: &CT_Tbl,
    depth: usize,
    item_count: &mut usize,
    text_bytes: &mut usize,
    projected_nodes: &mut usize,
    image_occurrences: &mut usize,
) -> Result<()> {
    if depth > MAX_NESTING_DEPTH {
        return Err(epub_error("table nesting exceeds the EPUB depth limit"));
    }
    for (_, raw) in &table.extra_xml {
        add_source_bytes(text_bytes, raw.len())?;
    }
    if let Some(properties) = &table.properties {
        if let Some(borders) = &properties.borders {
            measure_table_borders(borders, text_bytes)?;
        }
        for raw in &properties.revision_xml {
            add_source_bytes(text_bytes, raw.len())?;
        }
    }
    for row in &table.rows {
        add_projected_nodes(projected_nodes, 1)?;
        for (_, raw) in &row.extra_xml {
            add_source_bytes(text_bytes, raw.len())?;
        }
        if let Some(properties) = &row.properties {
            for raw in &properties.revision_xml {
                add_source_bytes(text_bytes, raw.len())?;
            }
        }
        for cell in &row.cells {
            add_projected_nodes(projected_nodes, 1)?;
            for (_, raw) in &cell.extra_xml {
                add_source_bytes(text_bytes, raw.len())?;
            }
            if let Some(properties) = &cell.properties {
                if let Some(shading) = &properties.shading {
                    measure_shading(shading, text_bytes)?;
                }
                if let Some(borders) = &properties.borders {
                    measure_table_borders(borders, text_bytes)?;
                }
                for (_, raw) in &properties.extra_xml {
                    add_source_bytes(text_bytes, raw.len())?;
                }
            }
            for content in &cell.content {
                *item_count = item_count
                    .checked_add(1)
                    .ok_or_else(|| epub_error("document item count overflow during EPUB export"))?;
                if *item_count > MAX_BODY_ITEMS {
                    return Err(epub_error(
                        "document has too many body items for EPUB export",
                    ));
                }
                add_projected_nodes(projected_nodes, 1)?;
                match content {
                    CellContent::Paragraph(paragraph) => measure_paragraph(
                        paragraph,
                        text_bytes,
                        projected_nodes,
                        image_occurrences,
                    )?,
                    CellContent::Table(nested) => measure_table(
                        nested,
                        depth + 1,
                        item_count,
                        text_bytes,
                        projected_nodes,
                        image_occurrences,
                    )?,
                    CellContent::ContentControl(_) => {}
                }
            }
        }
    }
    Ok(())
}

fn measure_paragraph(
    paragraph: &CT_P,
    text_bytes: &mut usize,
    projected_nodes: &mut usize,
    image_occurrences: &mut usize,
) -> Result<()> {
    if let Some(properties) = &paragraph.properties {
        for value in [
            properties.style_id.as_deref(),
            properties.line_rule.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            add_source_bytes(text_bytes, value.len())?;
        }
        if let Some(shading) = &properties.shading {
            measure_shading(shading, text_bytes)?;
        }
        for raw in &properties.numbering_revision_xml {
            add_source_bytes(text_bytes, raw.len())?;
        }
        for raw in &properties.revision_xml {
            add_source_bytes(text_bytes, raw.len())?;
        }
    }
    for hyperlink in &paragraph.hyperlinks {
        if hyperlink.run_start > hyperlink.run_end || hyperlink.run_end > paragraph.runs.len() {
            return Err(epub_error(
                "hyperlink span is outside its source paragraph during EPUB export",
            ));
        }
        let expansion = (hyperlink.run_end - hyperlink.run_start)
            .checked_add(1)
            .ok_or_else(|| epub_error("hyperlink span size overflow during EPUB export"))?;
        add_projected_nodes(projected_nodes, expansion)?;
        if let Some(value) = &hyperlink.rel_id {
            add_source_bytes(text_bytes, value.len())?;
        }
        if let Some(value) = &hyperlink.anchor {
            add_source_bytes(text_bytes, value.len())?;
        }
        for (name, value) in &hyperlink.extra_attributes {
            add_source_bytes(text_bytes, name.len())?;
            add_source_bytes(text_bytes, value.len())?;
        }
        for (_, _, raw) in &hyperlink.extra_xml {
            add_source_bytes(text_bytes, raw.len())?;
        }
    }
    for run in &paragraph.runs {
        add_projected_nodes(projected_nodes, 1 + run.content.len())?;
        if let Some(properties) = &run.properties {
            for value in [
                properties.font_ascii.as_deref(),
                properties.color.as_deref(),
                properties.vert_align.as_deref(),
            ]
            .into_iter()
            .flatten()
            {
                add_source_bytes(text_bytes, value.len())?;
            }
            if let Some(shading) = &properties.shading {
                measure_shading(shading, text_bytes)?;
            }
            for raw in &properties.revision_xml {
                add_source_bytes(text_bytes, raw.len())?;
            }
        }
        for content in &run.content {
            match content {
                RunContent::Text(text) | RunContent::DeletedText(text) => {
                    add_source_bytes(text_bytes, text.text.len())?;
                }
                RunContent::Tab | RunContent::Break(_) => add_source_bytes(text_bytes, 1)?,
                RunContent::Field(field) => measure_field(field, text_bytes, 0)?,
                RunContent::Drawing(drawing) => {
                    *image_occurrences = image_occurrences.checked_add(1).ok_or_else(|| {
                        epub_error("image occurrence count overflow during EPUB export")
                    })?;
                    if *image_occurrences > MAX_IMAGE_OCCURRENCES {
                        return Err(epub_error(
                            "document has too many image occurrences for EPUB export",
                        ));
                    }
                    let image = drawing.inline.as_ref().map(|image| {
                        (
                            image.embed_id.as_str(),
                            image.chart_rel_id.as_deref(),
                            image.description.as_deref(),
                            image.name.as_deref(),
                        )
                    });
                    let image = image.or_else(|| {
                        drawing.anchor.as_ref().map(|image| {
                            (
                                image.embed_id.as_str(),
                                image.chart_rel_id.as_deref(),
                                image.description.as_deref(),
                                image.name.as_deref(),
                            )
                        })
                    });
                    if let Some((embed_id, chart_rel_id, description, name)) = image {
                        add_source_bytes(text_bytes, embed_id.len())?;
                        for value in [chart_rel_id, description, name].into_iter().flatten() {
                            add_source_bytes(text_bytes, value.len())?;
                        }
                    }
                }
                RunContent::FootnoteRef { .. }
                | RunContent::EndnoteRef { .. }
                | RunContent::CommentReference { .. } => {}
            }
        }
        for raw in &run.extra_xml {
            add_source_bytes(text_bytes, raw.len())?;
        }
    }
    for (_, raw) in &paragraph.extra_xml {
        add_source_bytes(text_bytes, raw.len())?;
    }
    Ok(())
}

fn measure_shading(shading: &rdocx_oxml::properties::CT_Shd, text_bytes: &mut usize) -> Result<()> {
    add_source_bytes(text_bytes, shading.val.len())?;
    for value in [shading.color.as_deref(), shading.fill.as_deref()]
        .into_iter()
        .flatten()
    {
        add_source_bytes(text_bytes, value.len())?;
    }
    Ok(())
}

fn shading_is_simplified(shading: &rdocx_oxml::properties::CT_Shd) -> bool {
    !shading.val.eq_ignore_ascii_case("clear")
        || shading
            .color
            .as_deref()
            .is_some_and(|color| !color.eq_ignore_ascii_case("auto"))
        || shading
            .fill
            .as_deref()
            .is_some_and(|fill| !is_word_hex_color(fill))
}

fn is_word_hex_color(value: &str) -> bool {
    let digits = value.strip_prefix('#').unwrap_or(value);
    matches!(digits.len(), 3 | 6 | 8) && digits.bytes().all(|byte| byte.is_ascii_hexdigit())
}

fn measure_table_borders(
    borders: &rdocx_oxml::table::CT_TblBorders,
    text_bytes: &mut usize,
) -> Result<()> {
    for edge in [
        borders.top.as_ref(),
        borders.bottom.as_ref(),
        borders.left.as_ref(),
        borders.right.as_ref(),
        borders.inside_h.as_ref(),
        borders.inside_v.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        if let Some(color) = &edge.color {
            add_source_bytes(text_bytes, color.len())?;
        }
    }
    Ok(())
}

fn measure_field(
    field: &rdocx_oxml::text::Field,
    text_bytes: &mut usize,
    depth: usize,
) -> Result<()> {
    if depth >= MAX_NESTING_DEPTH {
        return Err(epub_error("field nesting exceeds the EPUB depth limit"));
    }
    add_source_bytes(text_bytes, field.cached_result.len())?;
    add_source_bytes(text_bytes, field.instruction.raw.len())?;
    add_source_bytes(text_bytes, field.instruction.name.len())?;
    for argument in &field.instruction.arguments {
        measure_field_argument(argument, text_bytes, depth + 1)?;
    }
    for switch in &field.instruction.switches {
        add_source_bytes(text_bytes, switch.name.len())?;
        if let Some(argument) = &switch.argument {
            measure_field_argument(argument, text_bytes, depth + 1)?;
        }
    }
    Ok(())
}

fn measure_field_argument(
    argument: &rdocx_oxml::text::FieldArgument,
    text_bytes: &mut usize,
    depth: usize,
) -> Result<()> {
    match argument {
        rdocx_oxml::text::FieldArgument::Text(text) => add_source_bytes(text_bytes, text.len()),
        rdocx_oxml::text::FieldArgument::Nested(field) => measure_field(field, text_bytes, depth),
    }
}

fn add_projected_nodes(total: &mut usize, additional: usize) -> Result<()> {
    *total = total
        .checked_add(additional)
        .ok_or_else(|| epub_error("projected XHTML node overflow during EPUB export"))?;
    if *total > MAX_PROJECTED_NODES {
        return Err(epub_error(
            "document has too many projected XHTML nodes for EPUB export",
        ));
    }
    Ok(())
}

fn add_source_bytes(total: &mut usize, additional: usize) -> Result<()> {
    *total = total
        .checked_add(additional)
        .ok_or_else(|| epub_error("document text size overflow during EPUB export"))?;
    if *total > MAX_SOURCE_TEXT_BYTES {
        return Err(epub_error("document text exceeds the EPUB source limit"));
    }
    Ok(())
}

fn validated_epub_image(data: &[u8]) -> Option<oxml_media::ImageFormat> {
    let format = oxml_media::ImageFormat::sniff(data)?;
    let probed = oxml_media::probe(data)?;
    if probed.format != format {
        return None;
    }
    match format {
        oxml_media::ImageFormat::Png if valid_png_structure(data) => Some(format),
        oxml_media::ImageFormat::Jpeg if valid_jpeg_structure(data) => Some(format),
        oxml_media::ImageFormat::Gif if valid_gif_structure(data) => Some(format),
        _ => None,
    }
}

fn valid_png_structure(data: &[u8]) -> bool {
    if !data.starts_with(b"\x89PNG\r\n\x1a\n") {
        return false;
    }
    let mut offset = 8_usize;
    let mut chunk_index = 0_usize;
    let mut bit_depth = None;
    let mut colour_type = None;
    let mut saw_plte = false;
    let mut saw_idat = false;
    let mut ended_idat = false;
    while let Some(header_end) = offset.checked_add(8) {
        let Some(header) = data.get(offset..header_end) else {
            return false;
        };
        let length = u32::from_be_bytes([header[0], header[1], header[2], header[3]]) as usize;
        let kind = &header[4..8];
        let Some(payload_end) = header_end.checked_add(length) else {
            return false;
        };
        let Some(chunk_end) = payload_end.checked_add(4) else {
            return false;
        };
        let Some(payload) = data.get(header_end..payload_end) else {
            return false;
        };
        let Some(crc_bytes) = data.get(payload_end..chunk_end) else {
            return false;
        };
        let expected_crc =
            u32::from_be_bytes([crc_bytes[0], crc_bytes[1], crc_bytes[2], crc_bytes[3]]);
        if !kind.iter().all(u8::is_ascii_alphabetic)
            || !kind[2].is_ascii_uppercase()
            || png_crc32(kind, payload) != expected_crc
        {
            return false;
        }
        match kind {
            b"IHDR" => {
                if chunk_index != 0
                    || length != 13
                    || payload[0..4] == [0, 0, 0, 0]
                    || payload[4..8] == [0, 0, 0, 0]
                    || !png_bit_depth_is_valid(payload[8], payload[9])
                    || payload[10] != 0
                    || payload[11] != 0
                    || payload[12] > 1
                {
                    return false;
                }
                bit_depth = Some(payload[8]);
                colour_type = Some(payload[9]);
            }
            b"PLTE" => {
                if chunk_index == 0
                    || saw_plte
                    || saw_idat
                    || length == 0
                    || length > 768
                    || !length.is_multiple_of(3)
                    || matches!(colour_type, Some(0 | 4))
                    || colour_type == Some(3)
                        && bit_depth.is_none_or(|depth| length / 3 > 1_usize << depth)
                {
                    return false;
                }
                saw_plte = true;
            }
            b"IDAT" => {
                if chunk_index == 0 || ended_idat || colour_type == Some(3) && !saw_plte {
                    return false;
                }
                saw_idat = true;
            }
            b"IEND" => {
                return length == 0 && saw_idat && chunk_end == data.len();
            }
            _ => {
                if kind[0].is_ascii_uppercase() {
                    return false;
                }
                if saw_idat {
                    ended_idat = true;
                }
            }
        }
        offset = chunk_end;
        chunk_index += 1;
    }
    false
}

fn png_bit_depth_is_valid(bit_depth: u8, colour_type: u8) -> bool {
    match colour_type {
        0 => matches!(bit_depth, 1 | 2 | 4 | 8 | 16),
        2 | 4 | 6 => matches!(bit_depth, 8 | 16),
        3 => matches!(bit_depth, 1 | 2 | 4 | 8),
        _ => false,
    }
}

fn png_crc32(kind: &[u8], payload: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for byte in kind.iter().chain(payload) {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = if crc & 1 == 0 {
                crc >> 1
            } else {
                (crc >> 1) ^ 0xedb8_8320
            };
        }
    }
    !crc
}

fn valid_jpeg_structure(data: &[u8]) -> bool {
    if !data.starts_with(b"\xff\xd8") {
        return false;
    }
    let mut offset = 2_usize;
    let mut in_scan = false;
    let mut saw_sof = false;
    let mut saw_sos = false;
    while offset < data.len() {
        if in_scan && data[offset] != 0xff {
            offset += 1;
            continue;
        }
        if data[offset] != 0xff {
            return false;
        }
        while data.get(offset) == Some(&0xff) {
            offset += 1;
        }
        let Some(&marker) = data.get(offset) else {
            return false;
        };
        offset += 1;
        if in_scan && (marker == 0 || (0xd0..=0xd7).contains(&marker)) {
            continue;
        }
        in_scan = false;
        if marker == 0xd9 {
            return saw_sof && saw_sos && offset == data.len();
        }
        if marker == 0xd8 {
            return false;
        }
        if marker == 0x01 || (0xd0..=0xd7).contains(&marker) {
            continue;
        }
        let Some(length_bytes) = data.get(offset..offset.saturating_add(2)) else {
            return false;
        };
        let length = usize::from(u16::from_be_bytes([length_bytes[0], length_bytes[1]]));
        if length < 2 {
            return false;
        }
        let Some(segment_end) = offset.checked_add(length) else {
            return false;
        };
        if segment_end > data.len() {
            return false;
        }
        if matches!(
            marker,
            0xc0 | 0xc1
                | 0xc2
                | 0xc3
                | 0xc5
                | 0xc6
                | 0xc7
                | 0xc9
                | 0xca
                | 0xcb
                | 0xcd
                | 0xce
                | 0xcf
        ) {
            saw_sof = true;
        }
        if marker == 0xda {
            if !saw_sof {
                return false;
            }
            saw_sos = true;
            in_scan = true;
        }
        offset = segment_end;
    }
    false
}

fn valid_gif_structure(data: &[u8]) -> bool {
    if !(data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a")) || data.len() < 13 {
        return false;
    }
    let packed = data[10];
    let global_table = if packed & 0x80 != 0 {
        3_usize << (usize::from(packed & 0x07) + 1)
    } else {
        0
    };
    let Some(mut offset) = 13_usize.checked_add(global_table) else {
        return false;
    };
    if offset > data.len() {
        return false;
    }
    let mut saw_image = false;
    loop {
        let Some(&kind) = data.get(offset) else {
            return false;
        };
        offset += 1;
        match kind {
            0x2c => {
                let Some(descriptor) = data.get(offset..offset.saturating_add(9)) else {
                    return false;
                };
                let width = u16::from_le_bytes([descriptor[4], descriptor[5]]);
                let height = u16::from_le_bytes([descriptor[6], descriptor[7]]);
                if width == 0 || height == 0 {
                    return false;
                }
                offset += 9;
                let local_table = if descriptor[8] & 0x80 != 0 {
                    3_usize << (usize::from(descriptor[8] & 0x07) + 1)
                } else {
                    0
                };
                let Some(after_table) = offset.checked_add(local_table) else {
                    return false;
                };
                if after_table >= data.len() {
                    return false;
                }
                let minimum_code_size = data[after_table];
                if !(2..=8).contains(&minimum_code_size) {
                    return false;
                }
                offset = after_table + 1;
                let Some((after_blocks, has_image_data)) = gif_sub_blocks_end(data, offset) else {
                    return false;
                };
                if !has_image_data {
                    return false;
                }
                offset = after_blocks;
                saw_image = true;
            }
            0x21 => {
                if data.get(offset).is_none() {
                    return false;
                }
                offset += 1;
                let Some((after_blocks, _)) = gif_sub_blocks_end(data, offset) else {
                    return false;
                };
                offset = after_blocks;
            }
            0x3b => return saw_image && offset == data.len(),
            _ => return false,
        }
    }
}

fn gif_sub_blocks_end(data: &[u8], mut offset: usize) -> Option<(usize, bool)> {
    let mut has_data = false;
    loop {
        let size = usize::from(*data.get(offset)?);
        offset = offset.checked_add(1)?;
        if size == 0 {
            return Some((offset, has_data));
        }
        has_data = true;
        offset = offset.checked_add(size)?;
        data.get(..offset)?;
    }
}

fn escape_xml(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn base64_encode(data: &[u8]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut output = String::with_capacity(data.len().div_ceil(3) * 4);
    for chunk in data.chunks(3) {
        let first = u32::from(chunk[0]);
        let second = chunk.get(1).copied().map(u32::from).unwrap_or(0);
        let third = chunk.get(2).copied().map(u32::from).unwrap_or(0);
        let bits = (first << 16) | (second << 8) | third;
        output.push(ALPHABET[((bits >> 18) & 63) as usize] as char);
        output.push(ALPHABET[((bits >> 12) & 63) as usize] as char);
        output.push(if chunk.len() > 1 {
            ALPHABET[((bits >> 6) & 63) as usize] as char
        } else {
            '='
        });
        output.push(if chunk.len() > 2 {
            ALPHABET[(bits & 63) as usize] as char
        } else {
            '='
        });
    }
    output
}

fn epub_error(message: impl Into<String>) -> Error {
    Error::Other(format!("EPUB export error: {}", message.into()))
}

fn zip_error(error: zip::result::ZipError) -> Error {
    epub_error(error.to_string())
}

fn output_limit_error() -> std::io::Error {
    std::io::Error::other("EPUB output exceeds the size limit")
}

#[cfg(test)]
mod tests {
    use std::io::Read;

    use oxml_core::custom_properties::CustomProperties;
    use rdocx_oxml::Twips;
    use rdocx_oxml::drawing::{CT_Drawing, CT_Inline};
    use rdocx_oxml::properties::{CT_PPr, CT_RPr, CT_Shd};
    use rdocx_oxml::shared::{ST_Jc, ST_Underline};
    use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_Tc, CT_TcPr};
    use rdocx_oxml::text::{BreakType, CT_R, CT_Text, Field, HyperlinkSpan};

    use super::*;
    use crate::{Length, ListLevel, ListNumberFormat, StyleBuilder};

    const PNG_1X1: &[u8] = &[
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6,
        0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 8, 29, 99, 96, 96, 96, 248, 15, 0,
        1, 4, 1, 0, 30, 115, 156, 64, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ];

    fn png_with_padding(padding: usize) -> Vec<u8> {
        let mut bytes = PNG_1X1[..33].to_vec();
        bytes.extend_from_slice(&(padding as u32).to_be_bytes());
        bytes.extend_from_slice(b"tEXt");
        bytes.resize(bytes.len() + padding, b'x');
        bytes.extend_from_slice(&png_crc32(b"tEXt", &vec![b'x'; padding]).to_be_bytes());
        bytes.extend_from_slice(&PNG_1X1[33..]);
        bytes
    }

    fn png_with_duplicate_ihdr() -> Vec<u8> {
        let mut bytes = PNG_1X1[..33].to_vec();
        bytes.extend_from_slice(&PNG_1X1[8..33]);
        bytes.extend_from_slice(&PNG_1X1[33..]);
        bytes
    }

    fn png_with_late_plte() -> Vec<u8> {
        let payload = [0_u8, 0, 0];
        let mut bytes = PNG_1X1[..PNG_1X1.len() - 12].to_vec();
        bytes.extend_from_slice(&(payload.len() as u32).to_be_bytes());
        bytes.extend_from_slice(b"PLTE");
        bytes.extend_from_slice(&payload);
        bytes.extend_from_slice(&png_crc32(b"PLTE", &payload).to_be_bytes());
        bytes.extend_from_slice(&PNG_1X1[PNG_1X1.len() - 12..]);
        bytes
    }

    fn png_with_oversized_indexed_palette() -> Vec<u8> {
        let palette = [0_u8, 0, 0, 127, 127, 127, 255, 255, 255];
        let mut bytes = PNG_1X1[..33].to_vec();
        bytes[24] = 1;
        bytes[25] = 3;
        let ihdr_crc = png_crc32(b"IHDR", &bytes[16..29]);
        bytes[29..33].copy_from_slice(&ihdr_crc.to_be_bytes());
        bytes.extend_from_slice(&(palette.len() as u32).to_be_bytes());
        bytes.extend_from_slice(b"PLTE");
        bytes.extend_from_slice(&palette);
        bytes.extend_from_slice(&png_crc32(b"PLTE", &palette).to_be_bytes());
        bytes.extend_from_slice(&PNG_1X1[33..]);
        bytes
    }

    fn png_with_chunk_type(kind: [u8; 4]) -> Vec<u8> {
        let mut bytes = PNG_1X1[..33].to_vec();
        bytes.extend_from_slice(&0_u32.to_be_bytes());
        bytes.extend_from_slice(&kind);
        bytes.extend_from_slice(&png_crc32(&kind, &[]).to_be_bytes());
        bytes.extend_from_slice(&PNG_1X1[33..]);
        bytes
    }

    fn jpeg_1x1() -> Vec<u8> {
        vec![
            0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00,
            0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06,
            0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c, 0x14, 0x0d,
            0x0c, 0x0b, 0x0b, 0x0c, 0x19, 0x12, 0x13, 0x0f, 0x14, 0x1d, 0x1a, 0x1f, 0x1e, 0x1d,
            0x1a, 0x1c, 0x1c, 0x20, 0x24, 0x2e, 0x27, 0x20, 0x22, 0x2c, 0x23, 0x1c, 0x1c, 0x28,
            0x37, 0x29, 0x2c, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1f, 0x27, 0x39, 0x3d, 0x38, 0x32,
            0x3c, 0x2e, 0x33, 0x34, 0x32, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01,
            0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x07, 0xff, 0xc4,
            0x00, 0x14, 0x10, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00,
            0x3f, 0x00, 0x7f, 0x7f, 0xff, 0xd9,
        ]
    }

    fn jpeg_with_repeated_soi() -> Vec<u8> {
        let mut bytes = jpeg_1x1();
        bytes.splice(2..2, [0xff, 0xd8]);
        bytes
    }

    fn progressive_jpeg_1x1() -> Vec<u8> {
        vec![
            0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00,
            0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06,
            0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c, 0x14, 0x0d,
            0x0c, 0x0b, 0x0b, 0x0c, 0x19, 0x12, 0x13, 0x0f, 0x14, 0x1d, 0x1a, 0x1f, 0x1e, 0x1d,
            0x1a, 0x1c, 0x1c, 0x20, 0x24, 0x2e, 0x27, 0x20, 0x22, 0x2c, 0x23, 0x1c, 0x1c, 0x28,
            0x37, 0x29, 0x2c, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1f, 0x27, 0x39, 0x3d, 0x38, 0x32,
            0x3c, 0x2e, 0x33, 0x34, 0x32, 0xff, 0xc2, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01,
            0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0xff, 0xda,
            0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x00, 0x01, 0x7f, 0xff, 0xc4, 0x00, 0x14, 0x10,
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x01, 0x05, 0x02, 0x7f,
            0xff, 0xc4, 0x00, 0x14, 0x10, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01,
            0x00, 0x06, 0x3f, 0x02, 0x7f, 0xff, 0xc4, 0x00, 0x14, 0x10, 0x01, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff,
            0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x01, 0x3f, 0x21, 0x7f, 0xff, 0xda, 0x00, 0x08,
            0x01, 0x01, 0x00, 0x00, 0x00, 0x10, 0xff, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x10, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x01, 0x3f, 0x10, 0x7f, 0xff,
            0xd9,
        ]
    }

    fn jpeg_with_sos_before_sof() -> Vec<u8> {
        let mut bytes = b"\xff\xd8\xff\xda\x00\x08\x01\x01\x00\x00\x3f\x00".to_vec();
        bytes.extend_from_slice(b"\xff\xc0\x00\x0b\x08\x00\x01\x00\x01\x01\x01\x11\x00");
        bytes.extend_from_slice(b"\xff\xd9");
        bytes
    }

    fn gif_1x1() -> Vec<u8> {
        b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff\x2c\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3b"
            .to_vec()
    }

    fn gif_with_minimum_code_size(size: u8) -> Vec<u8> {
        let mut bytes = gif_1x1();
        bytes[29] = size;
        bytes
    }

    fn gif_with_empty_image_data() -> Vec<u8> {
        let mut bytes = gif_1x1();
        bytes.truncate(30);
        bytes.extend_from_slice(b"\x00\x3b");
        bytes
    }

    fn gif_with_image_size(width: u16, height: u16) -> Vec<u8> {
        let mut bytes = gif_1x1();
        bytes[24..26].copy_from_slice(&width.to_le_bytes());
        bytes[26..28].copy_from_slice(&height.to_le_bytes());
        bytes
    }

    fn archive_entries(bytes: &[u8]) -> Vec<(String, Vec<u8>, CompressionMethod)> {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut entries = Vec::new();
        for index in 0..archive.len() {
            let mut file = archive.by_index(index).unwrap();
            let name = file.name().to_owned();
            let compression = file.compression();
            let mut contents = Vec::new();
            file.read_to_end(&mut contents).unwrap();
            entries.push((name, contents, compression));
        }
        entries
    }

    fn entry_text(entries: &[(String, Vec<u8>, CompressionMethod)], name: &str) -> String {
        String::from_utf8(
            entries
                .iter()
                .find(|entry| entry.0 == name)
                .unwrap()
                .1
                .clone(),
        )
        .unwrap()
    }

    #[test]
    fn epub_spine_and_navigation_follow_the_document_outline() {
        let mut document = Document::new();
        document.add_paragraph("preface");
        document.add_paragraph("First").set_style("Heading1");
        document.add_paragraph("Nested").set_style("Heading2");
        document.add_paragraph("first body");
        document.add_paragraph("Second").set_style("Heading1");
        document.add_paragraph("second body");

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let opf = entry_text(&entries, "EPUB/package.opf");
        let nav = entry_text(&entries, "EPUB/nav.xhtml");

        assert!(opf.contains("<itemref idref=\"front\"/>\n<itemref idref=\"chapter-001\"/>\n<itemref idref=\"chapter-002\"/>"));
        assert!(nav.contains("chapter-001.xhtml#heading-0001\">First</a><ol><li><a href=\"chapter-001.xhtml#heading-0002\">Nested"));
        assert!(nav.contains("chapter-002.xhtml#heading-0003\">Second"));
        assert!(entry_text(&entries, "EPUB/front.xhtml").contains("preface"));
        assert!(entry_text(&entries, "EPUB/chapter-001.xhtml").contains("first body"));
        assert!(!entry_text(&entries, "EPUB/chapter-001.xhtml").contains("second body"));
    }

    #[test]
    fn epub_archive_starts_with_uncompressed_mimetype_and_is_deterministic() {
        let mut document = Document::new();
        document.set_title("A & B");
        document.set_author("Ada");
        document.add_paragraph("stable");

        let first = document.to_epub_bytes().unwrap().bytes;
        let second = document.to_epub_bytes().unwrap().bytes;
        assert_eq!(first, second);
        assert_eq!(&first[..4], b"PK\x03\x04");
        assert_eq!(u16::from_le_bytes([first[8], first[9]]), 0);
        let name_len = u16::from_le_bytes([first[26], first[27]]) as usize;
        assert_eq!(&first[30..30 + name_len], b"mimetype");
        let entries = archive_entries(&first);
        assert_eq!(entries[0].0, "mimetype");
        assert_eq!(entries[0].1, b"application/epub+zip");
        assert_eq!(entries[0].2, CompressionMethod::Stored);
        let opf = entry_text(&entries, "EPUB/package.opf");
        assert!(opf.contains("<dc:title>A &amp; B</dc:title>"));
        assert!(opf.contains("<dc:creator>Ada</dc:creator>"));
        assert!(opf.contains("1980-01-01T00:00:00Z"));
    }

    #[test]
    fn epub_preserves_reflowable_text_lists_tables_links_and_images() {
        let mut document = Document::new();
        document.add_paragraph("Chapter").set_style("Heading1");
        document.add_bullet_list_item("bullet", 0);
        document.append_hyperlink("example", "https://example.com/?a=1&b=2");
        let mut table = document.add_table(1, 1);
        table.cell(0, 0).unwrap().set_text("cell");
        document.add_picture(PNG_1X1, "pixel.png", Length::pt(1.0), Length::pt(1.0));

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let chapter = entry_text(&entries, "EPUB/chapter-001.xhtml");
        assert!(chapter.contains("<ul>"));
        assert!(chapter.contains("<li>"));
        assert!(chapter.contains("bullet"));
        assert!(chapter.contains("<table>"));
        assert!(chapter.contains("<td>"));
        assert!(chapter.contains("cell"));
        assert!(chapter.contains("href=\"https://example.com/?a=1&amp;b=2\""));
        assert!(chapter.contains("<img alt=\"\""));
        assert!(chapter.contains("src=\"images/image-001.png\""));
        assert!(!chapter.contains("base64,"));
        assert_eq!(
            entries
                .iter()
                .find(|entry| entry.0 == "EPUB/images/image-001.png")
                .unwrap()
                .1,
            PNG_1X1
        );
    }

    #[test]
    fn epub_reports_lossy_content_without_dropping_supported_siblings() {
        let mut document = Document::new();
        document.add_paragraph("before");
        document
            .document
            .body
            .content
            .push(BodyContent::RawXml(b"<w:custom/>".to_vec()));
        document.add_paragraph("after");
        let source_before = document.document.to_xml().unwrap();

        let result = document.to_epub_bytes().unwrap();
        assert!(
            result
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path == "body[1]")
        );
        assert!(
            result
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path == "body/properties/section")
        );
        assert_eq!(document.document.to_xml().unwrap(), source_before);
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert!(body.contains("before"));
        assert!(body.contains("after"));

        let source_docx = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&source_docx).unwrap();
        assert_eq!(
            reopened.document.body.content[1],
            document.document.body.content[1]
        );
    }

    #[test]
    fn epub_save_replaces_an_existing_file_atomically() {
        let root = std::env::temp_dir().join(format!(
            "rdocx-epub-atomic-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        let _ = std::fs::remove_dir_all(&root);
        std::fs::create_dir_all(&root).unwrap();
        let path = root.join("book.epub");
        std::fs::write(&path, b"prior publication").unwrap();
        for attempt in 0..128_u8 {
            std::fs::create_dir(root.join(format!(
                ".book.epub.rdocx-{}-{attempt}.tmp",
                std::process::id()
            )))
            .unwrap();
        }

        let mut document = Document::new();
        document.add_paragraph("replacement");
        assert!(document.save_epub(&path).is_err());
        assert_eq!(std::fs::read(&path).unwrap(), b"prior publication");
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn epub_rejects_source_growth_before_building_the_archive() {
        let mut document = Document::new();
        document.add_paragraph(&"x".repeat(MAX_SOURCE_TEXT_BYTES + 1));

        let error = document.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("document text exceeds the EPUB source limit"));

        let mut cursor = BoundedCursor::new(4);
        assert!(cursor.write_all(b"12345").is_err());
        assert!(cursor.into_inner().is_empty());

        let mut described = Document::new();
        let relationship_id = described.embed_image(PNG_1X1, "pixel.png");
        let mut inline = CT_Inline::new(&relationship_id, 1, 1);
        inline.description = Some("x".repeat(MAX_SOURCE_TEXT_BYTES + 1));
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Drawing(CT_Drawing::inline(inline))];
        let mut paragraph = CT_P::new();
        paragraph.runs.push(run);
        described
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        let error = described.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("document text exceeds the EPUB source limit"));

        let mut formatted = Document::new();
        let mut run = CT_R::new("bounded");
        run.properties = Some(CT_RPr {
            font_ascii: Some("x".repeat(MAX_SOURCE_TEXT_BYTES + 1)),
            ..Default::default()
        });
        let mut paragraph = CT_P::new();
        paragraph.runs.push(run);
        formatted
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        let error = formatted.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("document text exceeds the EPUB source limit"));
    }

    #[test]
    fn epub_rejects_hyperlink_spans_outside_their_paragraph() {
        let mut document = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R::new("one run"));
        paragraph.hyperlinks.push(HyperlinkSpan {
            rel_id: Some("rId1".to_owned()),
            anchor: None,
            tooltip: None,
            doc_location: None,
            run_start: 0,
            run_end: usize::MAX,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let error = document.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("hyperlink span is outside"), "{error}");
    }

    #[test]
    fn epub_bounds_media_and_image_occurrences_before_html_projection() {
        let mut oversized_media = Document::new();
        let oversized = png_with_padding(MAX_MEDIA_BYTES + 1);
        let oversized_id = oversized_media.embed_image(&oversized, "oversized.bin");
        let mut oversized_run = CT_R::new("");
        oversized_run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
            &oversized_id,
            Length::pt(1.0).to_emu(),
            Length::pt(1.0).to_emu(),
        )))];
        let mut oversized_paragraph = CT_P::new();
        oversized_paragraph.runs.push(oversized_run);
        oversized_media
            .document
            .body
            .content
            .push(BodyContent::Paragraph(oversized_paragraph));
        let error = oversized_media.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("document images exceed the EPUB media limit"));

        let mut repeated_image = Document::new();
        let relationship_id = repeated_image.embed_image(PNG_1X1, "pixel.png");
        let mut paragraph = rdocx_oxml::text::CT_P::new();
        for _ in 0..=MAX_IMAGE_OCCURRENCES {
            let drawing = CT_Drawing::inline(CT_Inline::new(
                &relationship_id,
                Length::pt(1.0).to_emu(),
                Length::pt(1.0).to_emu(),
            ));
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Drawing(drawing)];
            paragraph.runs.push(run);
        }
        repeated_image
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        let error = repeated_image.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("too many image occurrences"));
    }

    #[test]
    fn epub_bounds_projection_trees_and_relationships_before_cloning() {
        let mut styles = Document::new();
        let template = styles.styles.styles[0].clone();
        styles.styles.styles = vec![template; MAX_STYLE_ITEMS + 1];
        let error = styles.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("too many styles"));

        let mut numbering = Document::new();
        let mut definitions = CT_Numbering::new();
        for id in 0..=MAX_NUMBERING_ITEMS as u32 {
            definitions.nums.push(CT_Num {
                num_id: id + 1,
                abstract_num_id: 0,
                extra_xml: Vec::new(),
                extra_attributes: Vec::new(),
            });
        }
        numbering.numbering = Some(definitions);
        let error = numbering.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("too many numbering items"));

        let mut relationships = Document::new();
        relationships.add_hyperlink_relationship(&format!(
            "https://example.com/{}",
            "x".repeat(MAX_PROJECTION_KEY_BYTES)
        ));
        let error = relationships.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("relationships exceed the EPUB projection limit"));

        let mut grid = Document::new();
        let mut table = CT_Tbl::new();
        table.grid = Some(CT_TblGrid {
            columns: (0..=MAX_PROJECTED_NODES)
                .map(|_| CT_TblGridCol { width: Twips(1) })
                .collect(),
            ..Default::default()
        });
        table.rows.push(CT_Row::new());
        grid.document.body.content.push(BodyContent::Table(table));
        let result = grid.to_epub_bytes().unwrap();
        assert!(
            result
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path == "body[0]/grid")
        );

        let mut controlled = Document::new();
        controlled.document = CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>"#,
        )
        .unwrap();
        assert!(render_body_projection(&controlled.document.body.content[0]).is_none());
        let result = controlled.to_epub_bytes().unwrap();
        assert!(
            result
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("content control"))
        );

        let mut drawing = CT_Inline::new("rId1", 1, 1);
        drawing.raw_xml = Some(vec![b'x'; MAX_SOURCE_TEXT_BYTES]);
        let projected = render_drawing_projection(&CT_Drawing::inline(drawing));
        assert!(projected.inline.unwrap().raw_xml.is_none());

        let mut field = Field::new("PAGE", "cached");
        field.dirty = Some(true);
        let projected = render_run_content(&RunContent::Field(field));
        let RunContent::Field(projected) = projected else {
            unreachable!();
        };
        assert!(projected.instruction.raw.is_empty());
        assert_eq!(projected.cached_result, "cached");
        assert_eq!(projected.dirty, Some(true));

        let parsed = CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pPrChange w:id="1" w:author="author"><w:pPr><w:spacing w:before="240"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="2" w:author="author"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>text</w:t></w:r></w:p></w:body></w:document>"#,
        )
        .unwrap();
        let BodyContent::Paragraph(source) = &parsed.body.content[0] else {
            unreachable!();
        };
        let projected = render_paragraph_projection(source);
        assert!(projected.properties.unwrap().change.is_none());
        let properties = projected.runs[0].properties.as_ref().unwrap();
        assert!(properties.change.is_none());
        assert!(properties.revision_markers.is_empty());
        assert!(properties.revision_xml.is_empty());
    }

    #[test]
    fn epub_heading_text_uses_only_bounded_direct_projected_runs() {
        let hidden = "controlled".repeat(MAX_SOURCE_TEXT_BYTES / 10 + 1);
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Visible heading</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>{hidden}</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>"#
        );
        let mut document = Document::new();
        document.document = CT_Document::from_xml(xml.as_bytes()).unwrap();

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let nav = entry_text(&entries, "EPUB/nav.xhtml");
        let chapter = entry_text(&entries, "EPUB/chapter-001.xhtml");
        assert!(nav.contains(">Visible heading</a>"), "{nav}");
        assert!(chapter.contains(">Visible heading</h1>"), "{chapter}");
        assert!(!nav.contains("controlledcontrolled"), "{nav}");
        assert!(!chapter.contains("controlledcontrolled"), "{chapter}");
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/content-control[0]"
                && diagnostic.message.contains("content control")
        }));
    }

    #[test]
    fn epub_reports_named_style_and_deep_heading_losses() {
        let mut document = Document::new();
        document.add_style(
            StyleBuilder::paragraph("Spaced", "Spaced").paragraph_properties(CT_PPr {
                space_before: Some(Twips(240)),
                ..Default::default()
            }),
        );
        document.add_paragraph("styled").set_style("Spaced");
        document.add_style(
            StyleBuilder::paragraph("DeepStyle", "Deep style").paragraph_properties(CT_PPr {
                outline_lvl: Some(6),
                ..Default::default()
            }),
        );
        document.add_paragraph("deep").set_style("Heading7");
        let direct = document.add_paragraph("direct deep");
        direct.inner.properties.get_or_insert_default().outline_lvl = Some(8);
        document.add_paragraph("style deep").set_style("DeepStyle");

        let result = document.to_epub_bytes().unwrap();
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/properties/style"
                && diagnostic.message.contains("style formatting")
        }));
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[1]/properties/heading-level"
                && diagnostic.message.contains("reduced to level 6")
        }));
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[2]/properties/heading-level"
                && diagnostic.message.contains("reduced to level 6")
        }));
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[3]/properties/heading-level"
                && diagnostic.message.contains("flattened to a paragraph")
        }));
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/chapter-001.xhtml");
        assert_eq!(body.matches("<h6").count(), 2, "{body}");
        assert!(body.contains("<p>style deep</p>"), "{body}");
    }

    #[test]
    fn epub_preserves_list_identity_and_no_number_levels() {
        let mut document = Document::new();
        let first = document.add_list_definition(&[ListLevel::decimal()]);
        let second = document.add_list_definition(&[ListLevel::decimal().start(3)]);
        let unmarked = document.add_list_definition(&[ListLevel::decimal()]);
        let abstract_id = document
            .numbering
            .as_ref()
            .unwrap()
            .nums
            .iter()
            .find(|item| item.num_id == unmarked)
            .unwrap()
            .abstract_num_id;
        document
            .numbering
            .as_mut()
            .unwrap()
            .abstract_nums
            .iter_mut()
            .find(|item| item.abstract_num_id == abstract_id)
            .unwrap()
            .levels[0]
            .num_fmt = Some(ST_NumberFormat::None);
        document.add_paragraph("first").set_numbering(first, 0);
        document.add_paragraph("restart").set_numbering(second, 0);
        document
            .add_paragraph("without marker")
            .set_numbering(unmarked, 0);

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert_eq!(body.matches("<ol").count(), 2, "{body}");
        assert!(body.contains("<ol start=\"3\">"), "{body}");
        assert!(body.contains("<ul class=\"no-marker\">"), "{body}");
        assert!(body.contains("<li>without marker</li>"), "{body}");
    }

    #[test]
    fn epub_does_not_invent_markers_for_producer_defined_numbering() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[ListLevel::decimal()]);
        let abstract_id = document
            .numbering
            .as_ref()
            .unwrap()
            .nums
            .iter()
            .find(|item| item.num_id == number)
            .unwrap()
            .abstract_num_id;
        document
            .numbering
            .as_mut()
            .unwrap()
            .abstract_nums
            .iter_mut()
            .find(|item| item.abstract_num_id == abstract_id)
            .unwrap()
            .levels[0] = {
            let mut level = CT_Lvl::new(0);
            level.num_fmt = Some(ST_NumberFormat::Other("chicago".to_owned()));
            level.start = Some(3);
            level.lvl_text = Some("custom".to_owned());
            level.suffix = Some(rdocx_oxml::numbering::ST_LvlSuffix::Space);
            level
        };
        document
            .add_paragraph("producer marker")
            .set_numbering(number, 0);

        let result = document.to_epub_bytes().unwrap();
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert!(!body.contains("<ol"), "{body}");
        assert!(!body.contains("<ul"), "{body}");
        assert!(body.contains("<p>producer marker</p>"), "{body}");
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/properties/numbering/format"
                && diagnostic.message
                    == "producer-defined numbering format was emitted without a marker during EPUB export"
        }));
        for (suffix, message) in [
            (
                "start",
                "producer-defined list start value was dropped during EPUB export",
            ),
            ("marker", "list marker text was dropped during EPUB export"),
            (
                "suffix",
                "list marker suffix was dropped during EPUB export",
            ),
        ] {
            assert!(result.diagnostics.iter().any(|diagnostic| {
                diagnostic.path == format!("body[0]/properties/numbering/{suffix}")
                    && diagnostic.message == message
            }));
        }
        assert!(result.diagnostics.iter().all(|diagnostic| {
            !diagnostic
                .message
                .contains("replaced by EPUB list semantics")
                && !diagnostic.message.contains("spacing was normalized")
        }));
    }

    #[test]
    fn epub_continues_one_numbering_instance_across_an_interruption() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[ListLevel::decimal().start(3)]);
        document.add_paragraph("three").set_numbering(number, 0);
        document.add_paragraph("four").set_numbering(number, 0);
        document.add_paragraph("interruption");
        document.add_paragraph("five").set_numbering(number, 0);

        let result = document.to_epub_bytes().unwrap();
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert_eq!(body.matches("<ol").count(), 2, "{body}");
        assert!(body.contains("<ol start=\"3\">"), "{body}");
        assert!(body.contains("<ol start=\"5\">"), "{body}");
    }

    #[test]
    fn epub_preserves_numbered_heading_elements_and_navigation_anchors() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[ListLevel::decimal()]);
        let mut heading = document.add_paragraph("Numbered chapter");
        heading.set_style("Heading1");
        heading.set_numbering(number, 0);
        document.add_paragraph("body");

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let chapter = entry_text(&entries, "EPUB/chapter-001.xhtml");
        assert!(
            chapter.contains("<li><h1 id=\"heading-0001\">Numbered chapter</h1>"),
            "{chapter}"
        );
        assert!(!chapter.contains("<li id=\"heading-0001\">"), "{chapter}");
        let nav = entry_text(&entries, "EPUB/nav.xhtml");
        assert!(nav.contains("chapter-001.xhtml#heading-0001"), "{nav}");
    }

    #[test]
    fn epub_restarts_nested_counters_when_the_parent_advances() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[
            ListLevel::decimal(),
            ListLevel::new(ListNumberFormat::LowerLetter),
        ]);
        document
            .add_paragraph("parent one")
            .set_numbering(number, 0);
        document.add_paragraph("child a").set_numbering(number, 1);
        document
            .add_paragraph("parent two")
            .set_numbering(number, 0);
        document
            .add_paragraph("child a again")
            .set_numbering(number, 1);

        let result = document.to_epub_bytes().unwrap();
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert_eq!(
            body.matches("style=\"list-style-type:lower-alpha\"")
                .count(),
            2
        );
        assert!(
            !body.contains("<ol start=\"2\" style=\"list-style-type:lower-alpha\""),
            "{body}"
        );
    }

    #[test]
    fn epub_reports_resolved_list_semantics_flattened_inside_a_table_cell() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[ListLevel::decimal()]);
        let mut table = CT_Tbl::new();
        let mut row = CT_Row::new();
        let mut cell = CT_Tc::new();
        let mut paragraph = CT_P::new();
        paragraph.add_run("cell list item");
        paragraph.properties = Some(CT_PPr {
            num_id: Some(number),
            num_ilvl: Some(0),
            ..Default::default()
        });
        cell.content.push(CellContent::Paragraph(paragraph));
        row.cells.push(cell);
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        let result = document.to_epub_bytes().unwrap();
        assert!(result.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .path
                .starts_with("body[0]/row[0]/cell[0]/content[")
                && diagnostic.path.ends_with("/properties/numbering")
                && diagnostic.message.contains("table-cell list semantics")
        }));
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert!(body.contains("<table>"), "{body}");
        assert!(body.contains("cell list item"), "{body}");
    }

    #[test]
    fn epub_preserves_standard_marker_formats_and_reports_custom_marker_losses() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[ListLevel::new(ListNumberFormat::UpperRoman)]);
        let abstract_id = document
            .numbering
            .as_ref()
            .unwrap()
            .nums
            .iter()
            .find(|item| item.num_id == number)
            .unwrap()
            .abstract_num_id;
        let level = &mut document
            .numbering
            .as_mut()
            .unwrap()
            .abstract_nums
            .iter_mut()
            .find(|item| item.abstract_num_id == abstract_id)
            .unwrap()
            .levels[0];
        level.lvl_text = Some("Article %1)".to_owned());
        level.rpr = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        level.lvl_jc = Some(ST_Jc::Right);
        document.add_paragraph("roman").set_numbering(number, 0);

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert!(
            body.contains("style=\"list-style-type:upper-roman\""),
            "{body}"
        );
        for path in [
            "body[0]/properties/numbering/marker",
            "body[0]/properties/numbering/marker-style",
            "body[0]/properties/numbering/alignment",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.path == path)
            );
        }
    }

    #[test]
    fn epub_lifts_page_breaks_out_of_paragraph_and_inline_formatting() {
        let mut document = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R {
            properties: Some(CT_RPr {
                bold: Some(true),
                ..Default::default()
            }),
            content: vec![
                RunContent::Text(CT_Text::new("before")),
                RunContent::Break(BreakType::Page),
                RunContent::Text(CT_Text::new("after")),
            ],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        let mut field_paragraph = CT_P::new();
        let mut field_run = CT_R::new("");
        field_run.content = vec![RunContent::Field(Field::new(
            "DISPLAY",
            "field-before\u{000c}field-after",
        ))];
        field_paragraph.runs.push(field_run);
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(field_paragraph));

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert!(!body.contains("<p><strong>before<hr"), "{body}");
        assert!(
            body.contains("<p><strong>before</strong></p><hr/><p><strong>after</strong></p>"),
            "{body}"
        );
        assert!(
            body.contains("<p>field-before</p><hr/><p>field-after</p>"),
            "{body}"
        );
    }

    #[test]
    fn epub_rejects_excessive_list_nesting_without_recursing_unboundedly() {
        let mut document = Document::new();
        let mut abstract_numbering = CT_AbstractNum::new(7);
        for level in 0..=MAX_NESTING_DEPTH as u32 {
            let mut definition = CT_Lvl::new(level);
            definition.num_fmt = Some(ST_NumberFormat::Decimal);
            abstract_numbering.levels.push(definition);
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("level {level}"));
            paragraph.properties = Some(rdocx_oxml::properties::CT_PPr {
                num_id: Some(9),
                num_ilvl: Some(level),
                ..Default::default()
            });
            document
                .document
                .body
                .content
                .push(BodyContent::Paragraph(paragraph));
        }
        document.numbering = Some(CT_Numbering {
            abstract_nums: vec![abstract_numbering],
            nums: vec![CT_Num {
                num_id: 9,
                abstract_num_id: 7,
                extra_xml: Vec::new(),
                extra_attributes: Vec::new(),
            }],
            root_attributes: Vec::new(),
            extra_xml: Vec::new(),
        });

        let error = document.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("list nesting exceeds"), "{error}");
    }

    #[test]
    fn epub_rewrites_only_exact_image_source_attributes() {
        let mut document = Document::new();
        document.add_paragraph("data:image/png;base64,AAAAAAAAAAA=");
        document.add_picture(PNG_1X1, "pixel.png", Length::pt(1.0), Length::pt(1.0));

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert!(
            body.contains("data:image/png;base64,AAAAAAAAAAA="),
            "{body}"
        );
        assert!(body.contains("src=\"images/image-001.png\""), "{body}");
    }

    #[test]
    fn epub_nested_lists_are_children_of_the_owning_list_item() {
        let mut document = Document::new();
        document.add_paragraph("Chapter").set_style("Heading1");
        document.add_bullet_list_item("parent", 0);
        document.add_bullet_list_item("child", 1);
        document.add_bullet_list_item("sibling", 0);

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let chapter = entry_text(&entries, "EPUB/chapter-001.xhtml");
        let parent = chapter.find("parent").unwrap();
        let nested_list = chapter[parent..].find("<ul>").unwrap();
        let parent_close = chapter[parent..].find("</li>").unwrap();
        assert!(nested_list < parent_close, "{chapter}");
        assert!(chapter.contains("child</li>\n</ul>\n</li>\n<li>sibling"));
    }

    #[test]
    fn epub_reports_run_xml_and_unrepresentable_hyperlinks() {
        let mut document = Document::new();
        let relative = document.add_hyperlink_relationship("chapter.html");
        let unsafe_target = document.add_hyperlink_relationship("javascript:alert(1)");
        let mut paragraph = document.add_paragraph("");
        paragraph.add_hyperlink("relative", &relative);
        paragraph.add_hyperlink("unsafe", &unsafe_target);
        let internal_start = paragraph.inner.runs.len();
        paragraph.add_run("internal");
        paragraph.inner.hyperlinks.push(HyperlinkSpan {
            rel_id: None,
            anchor: Some("bookmark".to_owned()),
            tooltip: None,
            doc_location: None,
            run_start: internal_start,
            run_end: internal_start + 1,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        paragraph.inner.runs[0]
            .extra_xml
            .push(b"<w:unsupported/>".to_vec());
        paragraph.inner.runs[0].extra_xml_positions.push(0);

        let result = document.to_epub_bytes().unwrap();
        let messages = result
            .diagnostics
            .iter()
            .map(|diagnostic| (diagnostic.path.as_str(), diagnostic.message.as_str()))
            .collect::<Vec<_>>();
        assert!(messages.iter().any(|(path, message)| {
            *path == "body[0]/run[0]/xml[0]" && message.contains("unmodelled run XML")
        }));
        assert_eq!(
            messages
                .iter()
                .filter(|(_, message)| message.contains("hyperlink"))
                .count(),
            3
        );
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert!(body.contains("relative"));
        assert!(body.contains("unsafe"));
        assert!(body.contains("internal"));
        assert!(!body.contains("<a href="));
    }

    #[test]
    fn epub_reports_typed_and_raw_paragraph_run_table_row_cell_and_field_losses() {
        let mut document = Document::new();
        document.document = CT_Document::from_xml(
            b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"urn:root-foreign\"><w:body><w:p xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:y=\"urn:foreign\"><w:bookmarkStart w:id=\"1\" w:name=\"place\"/><w:commentRangeStart w:id=\"2\"/><y:ins/><x:ins\n x:id=\"3\" x:author=\"author\"><x:r><x:t>revision</x:t></x:r></x:ins><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rPrChange w:id=\"4\" w:author=\"author\"><w:rPr/></w:rPrChange></w:rPr><w:t>run</w:t></w:r></w:p></w:body></w:document>",
        )
        .unwrap();
        let mut table = CT_Tbl::new();
        table.extra_xml.push((0, b"<w:tableRaw/>".to_vec()));
        let mut row = CT_Row::new();
        row.extra_xml.push((0, b"<w:rowRaw/>".to_vec()));
        let mut cell = CT_Tc::new();
        cell.extra_xml.push((0, b"<w:cellRaw/>".to_vec()));
        let mut cell_properties = CT_TcPr::default();
        cell_properties
            .extra_xml
            .push((0, b"<w:cellPropertyRaw/>".to_vec()));
        cell.properties = Some(cell_properties);
        row.cells.push(cell);
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        let mut field_paragraph = CT_P::new();
        let mut field_run = CT_R::new("");
        field_run.content = vec![RunContent::Field(Field::new("PAGE", "cached"))];
        field_paragraph.runs.push(field_run);
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(field_paragraph));

        let result = document.to_epub_bytes().unwrap();
        let messages = result
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>();
        for expected in [
            "bookmark marker",
            "comment range marker",
            "paragraph revision wrapper",
            "run property revision",
            "unmodelled table XML",
            "unmodelled table-row XML",
            "unmodelled table-cell XML",
            "unmodelled table-cell property XML",
            "field semantics",
        ] {
            assert!(
                messages.iter().any(|message| message.contains(expected)),
                "missing {expected}: {messages:?}"
            );
        }
        assert_eq!(
            messages
                .iter()
                .filter(|message| message.contains("paragraph revision wrapper"))
                .count(),
            1,
            "{messages:?}"
        );
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("unmodelled paragraph XML"))
                .count(),
            1
        );
    }

    #[test]
    fn epub_reports_each_dropped_document_metadata_field() {
        let mut document = Document::new();
        document.set_subject("subject");
        document.set_keywords("keywords");
        let properties = document.core_properties.get_or_insert_default();
        properties.description = Some("description".to_owned());
        properties.last_modified_by = Some("editor".to_owned());
        properties.created = Some("2026-08-24T00:00:00Z".to_owned());
        properties.modified = Some("2026-08-24T01:00:00Z".to_owned());
        document.custom_properties = Some(
            CustomProperties::from_xml(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="One"><vt:lpwstr>first</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Two"><vt:i4>2</vt:i4></property></Properties>"#,
            )
            .unwrap(),
        );

        let result = document.to_epub_bytes().unwrap();
        for path in [
            "metadata/subject",
            "metadata/description",
            "metadata/keywords",
            "metadata/last-modified-by",
            "metadata/created",
            "metadata/modified",
            "metadata/custom-property[0]",
            "metadata/custom-property[1]",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.path == path)
            );
        }
    }

    #[test]
    fn epub_reports_background_cell_shading_and_visible_default_style_losses() {
        let mut inert = Document::new();
        inert.styles.doc_defaults = None;
        inert.add_paragraph("inert default");
        let inert_result = inert.to_epub_bytes().unwrap();
        assert!(
            !inert_result
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path.ends_with("/properties/default-style"))
        );

        let mut document_defaults = Document::new();
        document_defaults.add_paragraph("visible document defaults");
        document_defaults.add_paragraph("second affected paragraph");
        let defaults_result = document_defaults.to_epub_bytes().unwrap();
        for path in [
            "body[0]/properties/default-style",
            "body[1]/properties/default-style",
        ] {
            assert!(defaults_result.diagnostics.iter().any(|diagnostic| {
                diagnostic.path == path && diagnostic.message.contains("default paragraph or run")
            }));
        }

        let mut preservation_only = Document::new();
        preservation_only.styles = rdocx_oxml::styles::CT_Styles::from_xml(
            br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rPrChange w:id="1" w:author="a"><w:rPr><w:b/></w:rPr></w:rPrChange><w:foreign/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:pPrChange w:id="2" w:author="a"><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrChange><w:foreign/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:pPrChange w:id="3" w:author="a"><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrChange><w:foreign/></w:pPr><w:rPr><w:rPrChange w:id="4" w:author="a"><w:rPr><w:i/></w:rPr></w:rPrChange><w:foreign/></w:rPr></w:style></w:styles>"#,
        )
        .unwrap();
        preservation_only.add_paragraph("preservation only");
        let preservation_result = preservation_only.to_epub_bytes().unwrap();
        assert!(
            !preservation_result
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path.ends_with("/properties/default-style"))
        );

        let mut document = Document::new();
        document.document.background_xml = Some(b"<w:background w:color=\"112233\"/>".to_vec());
        let default_style = document
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_type == StyleType::Paragraph && style.is_default)
            .unwrap();
        default_style.ppr = Some(CT_PPr {
            space_before: Some(Twips(240)),
            ..Default::default()
        });
        default_style.rpr = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        document.add_paragraph("default styled");
        document
            .add_paragraph("explicit heading")
            .set_style("Heading1");

        let mut table = CT_Tbl::new();
        let mut row = CT_Row::new();
        for shading in [
            CT_Shd {
                val: "horzStripe".to_owned(),
                color: None,
                fill: Some("FFFFFF".to_owned()),
            },
            CT_Shd {
                val: "clear".to_owned(),
                color: Some("00FF00".to_owned()),
                fill: None,
            },
            CT_Shd {
                val: "clear".to_owned(),
                color: None,
                fill: Some("invalid".to_owned()),
            },
        ] {
            let mut cell = CT_Tc::new();
            cell.properties = Some(CT_TcPr {
                shading: Some(shading),
                ..Default::default()
            });
            row.cells.push(cell);
        }
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        let result = document.to_epub_bytes().unwrap();
        for path in [
            "document/background",
            "body[0]/properties/default-style",
            "body[2]/row[0]/cell[0]/properties/shading",
            "body[2]/row[0]/cell[1]/properties/shading",
            "body[2]/row[0]/cell[2]/properties/shading",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.path == path),
                "missing {path}: {:?}",
                result.diagnostics
            );
        }
        assert!(
            !result
                .diagnostics
                .iter()
                .any(|diagnostic| { diagnostic.path == "body[1]/properties/default-style" })
        );
    }

    #[test]
    fn epub_packages_only_referenced_core_media_and_reports_each_unsupported_occurrence() {
        let mut document = Document::new();
        document.embed_image(PNG_1X1, "orphan.png");
        let unsupported = document.embed_image(b"II*\0unsupported", "scan.tiff");
        let mut paragraph = CT_P::new();
        for _ in 0..2 {
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
                &unsupported,
                Length::pt(1.0).to_emu(),
                Length::pt(1.0).to_emu(),
            )))];
            paragraph.runs.push(run);
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let result = document.to_epub_bytes().unwrap();
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("non-core EPUB image type"))
                .count(),
            2
        );
        let entries = archive_entries(&result.bytes);
        assert!(
            !entries
                .iter()
                .any(|(name, _, _)| name.starts_with("EPUB/images/"))
        );
        let opf = entry_text(&entries, "EPUB/package.opf");
        assert!(!opf.contains("media-type=\"image/"));
    }

    #[test]
    fn epub_packages_only_structurally_valid_byte_sniffed_core_images() {
        let mut document = Document::new();
        let valid = document.embed_image(PNG_1X1, "valid.bin");
        let forged = document.embed_image(b"not a PNG", "forged.png");
        let malformed = document.embed_image(b"\x89PNG\r\n\x1a\ntruncated", "malformed.png");
        let duplicate_ihdr = png_with_duplicate_ihdr();
        assert!(!valid_png_structure(&duplicate_ihdr));
        let duplicate = document.embed_image(&duplicate_ihdr, "duplicate-ihdr.png");
        let late_plte = png_with_late_plte();
        assert!(!valid_png_structure(&late_plte));
        let late_palette = document.embed_image(&late_plte, "late-palette.png");
        let oversized_palette = png_with_oversized_indexed_palette();
        assert!(!valid_png_structure(&oversized_palette));
        let indexed = document.embed_image(&oversized_palette, "oversized-palette.png");
        let safe_private_chunk = png_with_chunk_type(*b"teXt");
        assert!(valid_png_structure(&safe_private_chunk));
        let safe_private = document.embed_image(&safe_private_chunk, "safe-private.png");
        let non_letter_chunk = png_with_chunk_type(*b"t0XT");
        assert!(!valid_png_structure(&non_letter_chunk));
        let non_letter = document.embed_image(&non_letter_chunk, "non-letter.png");
        let lowercase_reserved_chunk = png_with_chunk_type(*b"text");
        assert!(!valid_png_structure(&lowercase_reserved_chunk));
        let lowercase_reserved =
            document.embed_image(&lowercase_reserved_chunk, "lowercase-reserved.png");
        let jpeg = jpeg_1x1();
        assert!(valid_jpeg_structure(&jpeg));
        let valid_jpeg = document.embed_image(&jpeg, "valid.jpeg");
        let progressive_jpeg = progressive_jpeg_1x1();
        assert!(valid_jpeg_structure(&progressive_jpeg));
        let valid_progressive_jpeg =
            document.embed_image(&progressive_jpeg, "valid-progressive.jpeg");
        let repeated_soi_jpeg = jpeg_with_repeated_soi();
        assert!(!valid_jpeg_structure(&repeated_soi_jpeg));
        let repeated_soi = document.embed_image(&repeated_soi_jpeg, "repeated-soi.jpeg");
        let early_sos_jpeg = jpeg_with_sos_before_sof();
        assert!(!valid_jpeg_structure(&early_sos_jpeg));
        let early_sos = document.embed_image(&early_sos_jpeg, "sos-before-sof.jpeg");
        let gif = gif_1x1();
        assert!(valid_gif_structure(&gif));
        let valid_gif = document.embed_image(&gif, "valid.gif");
        let zero_code_size_gif = gif_with_minimum_code_size(0);
        assert!(!valid_gif_structure(&zero_code_size_gif));
        assert!(!valid_gif_structure(&gif_with_minimum_code_size(9)));
        let zero_code_size = document.embed_image(&zero_code_size_gif, "zero-code-size.gif");
        let empty_data_gif = gif_with_empty_image_data();
        assert!(!valid_gif_structure(&empty_data_gif));
        let empty_data = document.embed_image(&empty_data_gif, "empty-data.gif");
        let zero_width_gif = gif_with_image_size(0, 1);
        assert!(!valid_gif_structure(&zero_width_gif));
        let zero_width = document.embed_image(&zero_width_gif, "zero-width.gif");
        let zero_height_gif = gif_with_image_size(1, 0);
        assert!(!valid_gif_structure(&zero_height_gif));
        let zero_height = document.embed_image(&zero_height_gif, "zero-height.gif");
        let active_svg = document.embed_image(
            br#"<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script><image href="https://example.test/tracker.png"/></svg>"#,
            "active.svg",
        );
        let mut paragraph = CT_P::new();
        for relationship_id in [
            valid,
            forged,
            malformed,
            duplicate,
            late_palette,
            indexed,
            safe_private,
            non_letter,
            lowercase_reserved,
            valid_jpeg,
            valid_progressive_jpeg,
            repeated_soi,
            early_sos,
            valid_gif,
            zero_code_size,
            empty_data,
            zero_width,
            zero_height,
            active_svg,
        ] {
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
                &relationship_id,
                1,
                1,
            )))];
            paragraph.runs.push(run);
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let result = document.to_epub_bytes().unwrap();
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("non-core EPUB image type"))
                .count(),
            14
        );
        let entries = archive_entries(&result.bytes);
        assert_eq!(
            entries
                .iter()
                .filter(|(name, _, _)| name.starts_with("EPUB/images/"))
                .count(),
            5
        );
        assert_eq!(
            entries
                .iter()
                .find(|(name, _, _)| name == "EPUB/images/image-001.png")
                .unwrap()
                .1,
            PNG_1X1
        );
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert_eq!(body.matches("<img ").count(), 5, "{body}");
        assert!(entries.iter().any(|(name, bytes, _)| {
            name.ends_with(".jpeg") && bytes.as_slice() == jpeg.as_slice()
        }));
        assert!(entries.iter().any(|(name, bytes, _)| {
            name.ends_with(".jpeg") && bytes.as_slice() == progressive_jpeg.as_slice()
        }));
        assert!(entries.iter().any(|(name, bytes, _)| {
            name.ends_with(".gif") && bytes.as_slice() == gif.as_slice()
        }));
        assert!(!body.contains("script"), "{body}");
        assert!(!body.contains("tracker.png"), "{body}");
    }

    #[test]
    fn epub_preserves_supported_image_alternative_descriptions() {
        let mut document = Document::new();
        let relationship_id = document.embed_image(PNG_1X1, "pixel.png");
        let mut inline = CT_Inline::new(
            &relationship_id,
            Length::pt(1.0).to_emu(),
            Length::pt(1.0).to_emu(),
        );
        inline.description = Some("A pixel & its <edge>".to_owned());
        inline.name = Some("Source pixel".to_owned());
        inline.raw_xml = Some(b"<wp:inline><a:extLst/></wp:inline>".to_vec());
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Drawing(CT_Drawing::inline(inline))];
        let mut paragraph = CT_P::new();
        paragraph.runs.push(run);
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let result = document.to_epub_bytes().unwrap();
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert!(
            body.contains("alt=\"A pixel &amp; its &lt;edge&gt;\""),
            "{body}"
        );
        for path in [
            "body[0]/run[0]/content[0]/name",
            "body[0]/run[0]/content[0]/extent",
            "body[0]/run[0]/content[0]/xml",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.path == path),
                "missing {path}: {:?}",
                result.diagnostics
            );
        }
    }

    #[test]
    fn epub_reports_alternate_drawings_preserved_spacing_and_column_breaks() {
        let mut document = Document::new();
        let mut run = CT_R::new("");
        run.properties = Some(CT_RPr {
            underline: Some(ST_Underline::None),
            ..Default::default()
        });
        run.content = vec![
            RunContent::Text(CT_Text::new(" spaced  text ")),
            RunContent::Break(BreakType::Column),
            RunContent::Text(CT_Text::new("after")),
        ];
        run.alt_drawings
            .push(CT_Drawing::inline(CT_Inline::new("rIdAlt", 1, 1)));
        let mut paragraph = CT_P::new();
        paragraph.runs.push(run);
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let result = document.to_epub_bytes().unwrap();
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert!(!body.contains("<u>"), "{body}");
        assert!(body.contains("<br/>"), "{body}");
        for path in [
            "body[0]/run[0]/alternate-drawing[0]",
            "body[0]/run[0]/content[0]/space",
            "body[0]/run[0]/content[1]",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.path == path),
                "missing {path}: {:?}",
                result.diagnostics
            );
        }
    }

    #[test]
    fn epub_reports_non_basic_underlines_shading_and_preserved_deleted_text_once() {
        let mut document = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.properties = Some(CT_PPr {
            shading: Some(CT_Shd {
                val: "horzStripe".to_owned(),
                color: Some("00FF00".to_owned()),
                fill: Some("FFFFFF".to_owned()),
            }),
            ..Default::default()
        });
        for underline in [
            ST_Underline::None,
            ST_Underline::Single,
            ST_Underline::Words,
            ST_Underline::Double,
            ST_Underline::Thick,
            ST_Underline::Dotted,
            ST_Underline::Dash,
            ST_Underline::DotDash,
            ST_Underline::DotDotDash,
            ST_Underline::Wave,
        ] {
            let mut run = CT_R::new("underlined ");
            run.properties = Some(CT_RPr {
                underline: Some(underline),
                ..Default::default()
            });
            paragraph.runs.push(run);
        }
        let mut foreground = CT_R::new("foreground");
        foreground.properties = Some(CT_RPr {
            shading: Some(CT_Shd {
                val: "clear".to_owned(),
                color: Some("FF0000".to_owned()),
                fill: None,
            }),
            ..Default::default()
        });
        paragraph.runs.push(foreground);
        let mut invalid = CT_R::new("invalid");
        invalid.properties = Some(CT_RPr {
            shading: Some(CT_Shd {
                val: "clear".to_owned(),
                color: None,
                fill: Some("not-a-colour".to_owned()),
            }),
            ..Default::default()
        });
        paragraph.runs.push(invalid);
        let mut deleted = CT_R::new("");
        deleted.content = vec![RunContent::DeletedText(CT_Text::new(" deleted "))];
        paragraph.runs.push(deleted);
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let result = document.to_epub_bytes().unwrap();
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("non-basic underline"))
                .count(),
            8
        );
        for path in [
            "body[0]/properties/shading",
            "body[0]/run[10]/properties/shading",
            "body[0]/run[11]/properties/shading",
            "body/properties/section",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.path == path),
                "missing {path}: {:?}",
                result.diagnostics
            );
        }
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| { diagnostic.path == "body[0]/run[12]/content[0]/space" })
                .count(),
            1
        );
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| {
                    diagnostic.path == "body[0]/run[12]/content[0]"
                        && diagnostic.message.contains("deleted-text revision")
                })
                .count(),
            1
        );
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert_eq!(body.matches("<u>").count(), 9, "{body}");
    }

    #[test]
    fn epub_rejects_forbidden_xml_1_0_characters_in_output_values() {
        for forbidden in [
            '\0', '\u{0001}', '\u{000B}', '\u{000C}', '\u{001F}', '\u{FFFE}', '\u{FFFF}',
        ] {
            let mut document = Document::new();
            document.add_paragraph(&format!("before{forbidden}after"));
            let error = document.to_epub_bytes().err().unwrap().to_string();
            assert!(
                error.contains("forbidden by XML 1.0"),
                "{forbidden:?}: {error}"
            );
        }

        let mut metadata = Document::new();
        metadata.set_title("bad\0title");
        let error = metadata.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("document title"));

        let mut relationship = Document::new();
        relationship.add_hyperlink_relationship("https://example.com/bad\0target");
        let error = relationship.to_epub_bytes().err().unwrap().to_string();
        assert!(error.contains("relationship value"));
    }

    #[test]
    fn epub_drops_malformed_absolute_hyperlink_uris_with_diagnostics() {
        let mut document = Document::new();
        for (text, target) in [
            ("bad percent", "https://example.com/%ZZ"),
            ("backslash", "https://example.com\\path"),
            ("missing host", "https:///path"),
            ("bad port", "https://example.com:abc/path"),
            ("bad literal", "https://[not-an-ip]/"),
            ("multiple userinfo", "https://one@two@example.com/"),
            ("userinfo bracket", "https://us[er]@example.com/"),
            ("multiple fragments", "https://example.com/#one#two"),
            ("path bracket", "https://example.com/a[b]"),
            ("query bracket", "https://example.com/?q=[bad]"),
            ("fragment bracket", "https://example.com/#bad[fragment]"),
            ("mail bracket", "mailto:user[bad]@example.com"),
        ] {
            let id = document.add_hyperlink_relationship(target);
            document.add_paragraph("").add_hyperlink(text, &id);
        }

        let result = document.to_epub_bytes().unwrap();
        assert_eq!(
            result
                .diagnostics
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("unsafe hyperlink target"))
                .count(),
            12
        );
        let entries = archive_entries(&result.bytes);
        let body = entry_text(&entries, "EPUB/document.xhtml");
        assert!(!body.contains("<a href="), "{body}");

        let mut valid = Document::new();
        for target in [
            "https://[2001:db8::1]/path",
            "https://[v1.example]/path",
            "https://user:pass@example.com/path",
        ] {
            let id = valid.add_hyperlink_relationship(target);
            valid.add_paragraph("").add_hyperlink("valid", &id);
        }
        let result = valid.to_epub_bytes().unwrap();
        let body = entry_text(&archive_entries(&result.bytes), "EPUB/document.xhtml");
        assert_eq!(body.matches("<a href=").count(), 3, "{body}");
    }

    #[test]
    fn epub_heading_anchors_follow_the_exact_source_paragraph() {
        let mut document = Document::new();
        document.add_paragraph("Root").set_style("Heading1");
        document.add_paragraph("style-derived");
        let BodyContent::Paragraph(style_derived) = &mut document.document.body.content[1] else {
            unreachable!();
        };
        style_derived.properties.get_or_insert_default().outline_lvl = Some(1);
        document.add_paragraph("Direct").set_style("Heading2");

        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let chapter = entry_text(&entries, "EPUB/chapter-001.xhtml");
        assert!(chapter.contains("<h2>style-derived</h2>"));
        assert!(chapter.contains("<h2 id=\"heading-0002\">Direct</h2>"));
        let nav = entry_text(&entries, "EPUB/nav.xhtml");
        assert!(nav.contains("chapter-001.xhtml#heading-0002\">Direct"));
    }

    fn command_sha256(path: &std::path::Path) -> String {
        let output = std::process::Command::new("shasum")
            .arg("-a")
            .arg("256")
            .arg(path)
            .output()
            .or_else(|_| std::process::Command::new("sha256sum").arg(path).output())
            .expect("a SHA-256 command is installed");
        assert!(output.status.success());
        String::from_utf8(output.stdout)
            .unwrap()
            .split_whitespace()
            .next()
            .unwrap()
            .to_owned()
    }

    #[test]
    #[ignore = "requires EPUBCHECK_JAR pointing at pinned EPUBCheck 5.3.0"]
    fn epubcheck_5_3_0_accepts_the_source_built_publication() {
        const EPUBCHECK_5_3_0_JAR_SHA256: &str =
            "f7f96617c929371821609b88c8484d6dc9f24fe916499863c46094c5fb778a65";
        let jar = std::env::var_os("EPUBCHECK_JAR").expect("EPUBCHECK_JAR is set");
        assert_eq!(
            command_sha256(std::path::Path::new(&jar)),
            EPUBCHECK_5_3_0_JAR_SHA256
        );
        let version = std::process::Command::new("java")
            .arg("-jar")
            .arg(&jar)
            .arg("--version")
            .output()
            .expect("run EPUBCheck 5.3.0");
        let version_text = format!(
            "{}{}",
            String::from_utf8_lossy(&version.stdout),
            String::from_utf8_lossy(&version.stderr)
        );
        assert!(version_text.contains("5.3.0"), "{version_text}");

        let root = std::env::temp_dir().join(format!("rdocx-epubcheck-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        std::fs::create_dir_all(&root).unwrap();
        let path = root.join("book.epub");
        let mut document = Document::new();
        document.document.background_xml = Some(b"<w:background w:color=\"F0F0F0\"/>".to_vec());
        document.styles.doc_defaults = Some(rdocx_oxml::styles::CT_DocDefaults {
            ppr: Some(CT_PPr {
                space_after: Some(Twips(120)),
                ..Default::default()
            }),
            rpr: Some(CT_RPr {
                italic: Some(true),
                ..Default::default()
            }),
        });
        let default_style = document
            .styles
            .styles
            .iter_mut()
            .find(|style| style.style_type == StyleType::Paragraph && style.is_default)
            .unwrap();
        default_style.ppr = Some(CT_PPr {
            space_after: Some(Twips(120)),
            ..Default::default()
        });
        default_style.rpr = Some(CT_RPr {
            italic: Some(true),
            ..Default::default()
        });
        document.add_paragraph("Front matter");
        let chapter_number = document.add_list_definition(&[ListLevel::decimal()]);
        let mut first_heading = document.add_paragraph("First");
        first_heading.set_style("Heading1");
        first_heading.set_numbering(chapter_number, 0);
        document.add_bullet_list_item("parent", 0);
        document.add_bullet_list_item("child", 1);
        let mut page_break = CT_R::new("before break");
        page_break.properties = Some(CT_RPr {
            underline: Some(ST_Underline::None),
            ..Default::default()
        });
        page_break.content.push(RunContent::Break(BreakType::Page));
        page_break
            .content
            .push(RunContent::Text(CT_Text::new("after break")));
        let mut page_break_paragraph = CT_P::new();
        page_break_paragraph.runs.push(page_break);
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(page_break_paragraph));
        document.add_paragraph("Nested").set_style("Heading2");
        document.add_picture(PNG_1X1, "pixel.png", Length::pt(1.0), Length::pt(1.0));
        let BodyContent::Paragraph(image_paragraph) =
            document.document.body.content.last_mut().unwrap()
        else {
            unreachable!();
        };
        let RunContent::Drawing(image) = &mut image_paragraph.runs[0].content[0] else {
            unreachable!();
        };
        image.inline.as_mut().unwrap().description = Some("Fixture pixel".to_owned());
        let continued = document.add_list_definition(&[ListLevel::decimal().start(3)]);
        document.add_paragraph("three").set_numbering(continued, 0);
        document.add_paragraph("interruption");
        document.add_paragraph("four").set_numbering(continued, 0);
        let mut table = CT_Tbl::new();
        let mut row = CT_Row::new();
        let mut cell = CT_Tc::new();
        cell.content.clear();
        let mut cell_list = CT_P::new();
        cell_list.add_run("table-cell list item");
        cell_list.properties = Some(CT_PPr {
            num_id: Some(continued),
            num_ilvl: Some(0),
            ..Default::default()
        });
        cell.content.push(CellContent::Paragraph(cell_list));
        cell.properties = Some(CT_TcPr {
            shading: Some(CT_Shd {
                val: "diagStripe".to_owned(),
                color: Some("00FF00".to_owned()),
                fill: Some("FFFFFF".to_owned()),
            }),
            ..Default::default()
        });
        row.cells.push(cell);
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        document.add_style(
            StyleBuilder::paragraph("OracleDeep", "Oracle deep").paragraph_properties(CT_PPr {
                outline_lvl: Some(6),
                ..Default::default()
            }),
        );
        document
            .add_paragraph("style-derived deep heading")
            .set_style("OracleDeep");
        let unsafe_svg = document.embed_image(
            br#"<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>"#,
            "active.svg",
        );
        let duplicate_png = png_with_duplicate_ihdr();
        let malformed_png = document.embed_image(&duplicate_png, "duplicate-ihdr.png");
        let oversized_palette = png_with_oversized_indexed_palette();
        let oversized_palette_png =
            document.embed_image(&oversized_palette, "oversized-palette.png");
        let invalid_chunk = png_with_chunk_type(*b"t0XT");
        let invalid_chunk_png = document.embed_image(&invalid_chunk, "invalid-chunk.png");
        let valid_jpeg_bytes = jpeg_1x1();
        let valid_jpeg = document.embed_image(&valid_jpeg_bytes, "valid.jpeg");
        let progressive_jpeg_bytes = progressive_jpeg_1x1();
        let progressive_jpeg =
            document.embed_image(&progressive_jpeg_bytes, "valid-progressive.jpeg");
        let repeated_soi_bytes = jpeg_with_repeated_soi();
        let repeated_soi = document.embed_image(&repeated_soi_bytes, "repeated-soi.jpeg");
        let early_sos_bytes = jpeg_with_sos_before_sof();
        let early_sos = document.embed_image(&early_sos_bytes, "sos-before-sof.jpeg");
        let valid_gif_bytes = gif_1x1();
        let valid_gif = document.embed_image(&valid_gif_bytes, "valid.gif");
        let zero_code_size_bytes = gif_with_minimum_code_size(0);
        let zero_code_size = document.embed_image(&zero_code_size_bytes, "zero-code-size.gif");
        let empty_gif_bytes = gif_with_empty_image_data();
        let empty_gif = document.embed_image(&empty_gif_bytes, "empty-data.gif");
        let zero_width_bytes = gif_with_image_size(0, 1);
        let zero_width = document.embed_image(&zero_width_bytes, "zero-width.gif");
        let zero_height_bytes = gif_with_image_size(1, 0);
        let zero_height = document.embed_image(&zero_height_bytes, "zero-height.gif");
        let mut unsafe_run = CT_R::new("");
        unsafe_run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
            &unsafe_svg,
            1,
            1,
        )))];
        let mut lossy = CT_P::new();
        unsafe_run.properties = Some(CT_RPr {
            underline: Some(ST_Underline::Double),
            shading: Some(CT_Shd {
                val: "diagStripe".to_owned(),
                color: Some("FF0000".to_owned()),
                fill: Some("FFFF00".to_owned()),
            }),
            ..Default::default()
        });
        unsafe_run
            .content
            .push(RunContent::DeletedText(CT_Text::new(" deleted ")));
        lossy.runs.push(unsafe_run);
        let mut malformed_run = CT_R::new("");
        malformed_run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
            &malformed_png,
            1,
            1,
        )))];
        lossy.runs.push(malformed_run);
        let mut oversized_palette_run = CT_R::new("");
        oversized_palette_run.content = vec![RunContent::Drawing(CT_Drawing::inline(
            CT_Inline::new(&oversized_palette_png, 1, 1),
        ))];
        lossy.runs.push(oversized_palette_run);
        let mut invalid_chunk_run = CT_R::new("");
        invalid_chunk_run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
            &invalid_chunk_png,
            1,
            1,
        )))];
        lossy.runs.push(invalid_chunk_run);
        for relationship_id in [
            valid_jpeg,
            progressive_jpeg,
            repeated_soi,
            early_sos,
            valid_gif,
            zero_code_size,
            empty_gif,
            zero_width,
            zero_height,
        ] {
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Drawing(CT_Drawing::inline(CT_Inline::new(
                &relationship_id,
                1,
                1,
            )))];
            lossy.runs.push(run);
        }
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(lossy));
        let aliased = CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:root-foreign"><w:body><w:p xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:y="urn:foreign"><y:ins/><x:ins x:id="9" x:author="oracle"><x:r><x:t>aliased revision</x:t></x:r></x:ins></w:p></w:body></w:document>"#,
        )
        .unwrap();
        document
            .document
            .extra_namespaces
            .push(("xmlns:x".to_owned(), "urn:root-foreign".to_owned()));
        document.document.body.content.extend(aliased.body.content);
        let unsafe_userinfo =
            document.add_hyperlink_relationship("https://us[er]@example.com/path");
        document
            .add_paragraph("")
            .add_hyperlink("unsafe userinfo", &unsafe_userinfo);
        document.add_paragraph("Second").set_style("Heading1");
        document.add_paragraph("last");
        let result = document.to_epub_bytes().unwrap();
        let entries = archive_entries(&result.bytes);
        let opf = entry_text(&entries, "EPUB/package.opf");
        assert!(opf.contains("<itemref idref=\"front\"/>\n<itemref idref=\"chapter-001\"/>\n<itemref idref=\"chapter-002\"/>"));
        assert!(entry_text(&entries, "EPUB/front.xhtml").contains("Front matter"));
        let nav = entry_text(&entries, "EPUB/nav.xhtml");
        assert!(nav.contains("First</a><ol><li><a href=\"chapter-001.xhtml#heading-0002\">Nested"));
        assert!(nav.contains("chapter-002.xhtml#heading-0003\">Second"));
        let first = entry_text(&entries, "EPUB/chapter-001.xhtml");
        assert!(first.contains("<li><h1 id=\"heading-0001\">First</h1>"));
        assert_eq!(first.matches("<u>").count(), 1, "{first}");
        assert!(first.contains("<u> deleted </u>"), "{first}");
        let parent = first.find("parent").unwrap();
        assert!(first[parent..].find("<ul>").unwrap() < first[parent..].find("</li>").unwrap());
        assert!(first.contains("src=\"images/image-001.png\""));
        assert!(first.contains("alt=\"Fixture pixel\""));
        assert!(!first.contains("<script"), "{first}");
        assert!(first.contains("unsafe userinfo"), "{first}");
        assert!(!first.contains("us[er]@example.com"), "{first}");
        assert!(first.contains("<ol start=\"3\">"));
        assert!(first.contains("<ol start=\"4\">"));
        assert!(first.contains("<table>"));
        assert!(first.contains("table-cell list item"));
        assert!(
            first.contains("<p>style-derived deep heading</p>"),
            "{first}"
        );
        assert_eq!(
            entries
                .iter()
                .filter(|(name, _, _)| name.starts_with("EPUB/images/"))
                .count(),
            4
        );
        assert!(entries.iter().any(|(name, bytes, _)| {
            name.ends_with(".jpeg") && bytes.as_slice() == valid_jpeg_bytes.as_slice()
        }));
        assert!(entries.iter().any(|(name, bytes, _)| {
            name.ends_with(".jpeg") && bytes.as_slice() == progressive_jpeg_bytes.as_slice()
        }));
        assert!(entries.iter().any(|(name, bytes, _)| {
            name.ends_with(".gif") && bytes.as_slice() == valid_gif_bytes.as_slice()
        }));
        for expected in [
            "non-core EPUB image type",
            "flattened to a paragraph",
            "non-basic underline",
            "run shading pattern",
            "deleted-text revision",
            "preserved Word text spacing",
            "final section properties",
            "document background",
            "default paragraph or run formatting",
            "table-cell shading pattern",
            "paragraph revision wrapper",
            "unmodelled paragraph XML",
            "unsafe hyperlink target",
        ] {
            assert!(
                result
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.message.contains(expected)),
                "missing {expected}: {:?}",
                result.diagnostics
            );
        }
        std::fs::write(&path, &result.bytes).unwrap();
        let checked = std::process::Command::new("java")
            .arg("-jar")
            .arg(jar)
            .arg(&path)
            .output()
            .expect("validate EPUB");
        let report = format!(
            "{}{}",
            String::from_utf8_lossy(&checked.stdout),
            String::from_utf8_lossy(&checked.stderr)
        );
        assert!(checked.status.success(), "{report}");
        std::fs::remove_dir_all(root).unwrap();
    }
}
