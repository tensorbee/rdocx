//! Header and footer elements: `CT_HdrFtr`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::{local_namespace_overrides, namespace_bindings, word_prefixes_at};
use crate::properties::is_word_element;
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::table::CT_Tbl;
use crate::text::{
    CT_P, ROOT_R_BINDING, ROOT_WP_BINDING, declare_w14_on_part_root, root_binding_scope,
};

const VML_NS: &str = "urn:schemas-microsoft-com:vml";
const OFFICE_NS: &str = "urn:schemas-microsoft-com:office:office";
const RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const TEXT_WATERMARK_SHAPE_TYPE_ID: &str = "rdocx-watermark-type-text";
const IMAGE_WATERMARK_SHAPE_TYPE_ID: &str = "rdocx-watermark-type-image";
const WATERMARK_SHAPE_TYPE_PREFIX: &str = "rdocx-watermark-type-";

/// A conservative layout projection of one VML watermark shape.
#[derive(Debug, Clone, PartialEq)]
pub enum VmlWatermark {
    Text {
        text: String,
        width_pt: f64,
        height_pt: f64,
        rotation_degrees: f64,
        color: String,
        font_family: Option<String>,
        opacity: f64,
    },
    Image {
        relationship_id: String,
        width_pt: f64,
        height_pt: f64,
        rotation_degrees: f64,
        opacity: f64,
    },
}

impl VmlWatermark {
    /// Project a supported `w:pict` subtree without taking ownership of its XML.
    #[doc(hidden)]
    pub fn from_pict_xml(xml: &[u8]) -> Option<Self> {
        parse_vml_watermark(xml)
    }

    /// Write the canonical VML subtree used by the native authoring facade.
    #[doc(hidden)]
    pub fn to_pict_xml(&self) -> Vec<u8> {
        let mut writer = Writer::new(Vec::new());
        let mut pict = BytesStart::new("w:pict");
        pict.push_attribute(("xmlns:w", W_NS));
        pict.push_attribute(("xmlns:v", VML_NS));
        pict.push_attribute(("xmlns:o", OFFICE_NS));
        pict.push_attribute(("xmlns:r", RELATIONSHIPS_NS));
        writer
            .write_event(Event::Start(pict))
            .expect("writing VML to memory cannot fail");

        let (width, height, rotation, opacity) = match self {
            Self::Text {
                width_pt,
                height_pt,
                rotation_degrees,
                opacity,
                ..
            }
            | Self::Image {
                width_pt,
                height_pt,
                rotation_degrees,
                opacity,
                ..
            } => (*width_pt, *height_pt, *rotation_degrees, *opacity),
        };
        let style = format!(
            "position:absolute;width:{}pt;height:{}pt;rotation:{};z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
            compact_number(width),
            compact_number(height),
            compact_number(rotation)
        );
        let (shape_type_id, shape_type_number, text_path) = match self {
            Self::Text { .. } => (TEXT_WATERMARK_SHAPE_TYPE_ID, "136", true),
            Self::Image { .. } => (IMAGE_WATERMARK_SHAPE_TYPE_ID, "75", false),
        };
        let mut shape_type = BytesStart::new("v:shapetype");
        shape_type.push_attribute(("id", shape_type_id));
        shape_type.push_attribute(("coordsize", "21600,21600"));
        shape_type.push_attribute(("o:spt", shape_type_number));
        shape_type.push_attribute(("path", "m0,0l21600,0,21600,21600,0,21600xe"));
        if !text_path {
            shape_type.push_attribute(("filled", "f"));
            shape_type.push_attribute(("stroked", "f"));
        }
        writer
            .write_event(Event::Start(shape_type))
            .expect("writing VML to memory cannot fail");
        let mut shape_path = BytesStart::new("v:path");
        if text_path {
            shape_path.push_attribute(("textpathok", "t"));
        } else {
            shape_path.push_attribute(("o:extrusionok", "f"));
            shape_path.push_attribute(("o:connecttype", "rect"));
        }
        writer
            .write_event(Event::Empty(shape_path))
            .expect("writing VML to memory cannot fail");
        if text_path {
            let mut template_text_path = BytesStart::new("v:textpath");
            template_text_path.push_attribute(("on", "t"));
            template_text_path.push_attribute(("fitshape", "t"));
            writer
                .write_event(Event::Empty(template_text_path))
                .expect("writing VML to memory cannot fail");
        }
        writer
            .write_event(Event::End(BytesEnd::new("v:shapetype")))
            .expect("writing VML to memory cannot fail");

        let mut shape = BytesStart::new("v:shape");
        shape.push_attribute(("id", "rdocx-watermark"));
        shape.push_attribute(("o:spid", "_x0000_s1025"));
        shape.push_attribute((
            "type",
            match self {
                Self::Text { .. } => "#rdocx-watermark-type-text",
                Self::Image { .. } => "#rdocx-watermark-type-image",
            },
        ));
        shape.push_attribute(("style", style.as_str()));
        shape.push_attribute(("stroked", "f"));
        if let Self::Text { color, .. } = self {
            shape.push_attribute(("fillcolor", color.as_str()));
        }
        writer
            .write_event(Event::Start(shape))
            .expect("writing VML to memory cannot fail");

        let mut fill = BytesStart::new("v:fill");
        let opacity = compact_number(opacity);
        fill.push_attribute(("opacity", opacity.as_str()));
        writer
            .write_event(Event::Empty(fill))
            .expect("writing VML to memory cannot fail");

        match self {
            Self::Text {
                text, font_family, ..
            } => {
                let family = font_family.as_deref().unwrap_or("Calibri");
                let text_style = format!("font-family:\"{family}\";font-size:1pt");
                let mut textpath = BytesStart::new("v:textpath");
                textpath.push_attribute(("on", "t"));
                textpath.push_attribute(("fitshape", "t"));
                textpath.push_attribute(("style", text_style.as_str()));
                textpath.push_attribute(("string", text.as_str()));
                writer
                    .write_event(Event::Empty(textpath))
                    .expect("writing VML to memory cannot fail");
            }
            Self::Image {
                relationship_id, ..
            } => {
                let mut image = BytesStart::new("v:imagedata");
                image.push_attribute(("r:id", relationship_id.as_str()));
                image.push_attribute(("o:title", ""));
                writer
                    .write_event(Event::Empty(image))
                    .expect("writing VML to memory cannot fail");
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new("v:shape")))
            .expect("writing VML to memory cannot fail");
        writer
            .write_event(Event::End(BytesEnd::new("w:pict")))
            .expect("writing VML to memory cannot fail");
        writer.into_inner()
    }
}

/// `CT_HdrFtr` — Content of a header or footer part.
///
/// Contains paragraphs (and potentially tables, same as a document body).
#[derive(Debug, Clone, PartialEq)]
pub struct CT_HdrFtr {
    pub paragraphs: Vec<CT_P>,
    /// Supported VML watermark shapes projected from the original part bytes.
    watermarks: Vec<VmlWatermark>,
    /// Extra namespace declarations captured from the root element.
    pub extra_namespaces: Vec<(String, String)>,
    /// Non-namespace attributes of the root element, such as `mc:Ignorable`,
    /// in source order.
    root_attributes: Vec<(String, String)>,
    /// Unknown child elements captured as raw XML.
    pub extra_xml: Vec<Vec<u8>>,
    /// How many paragraphs precede each entry of `extra_xml`, so that a
    /// rewrite puts a table or a content control back where it was. An entry
    /// without a position is written after the last paragraph.
    extra_xml_positions: Vec<usize>,
    /// The namespace bindings of the root element, which a raw child is
    /// parsed with when replacement reaches into it.
    pub(crate) word_prefixes: Vec<String>,
}

#[allow(non_snake_case)]
impl CT_HdrFtr {
    pub fn new() -> Self {
        CT_HdrFtr {
            paragraphs: Vec::new(),
            watermarks: Vec::new(),
            extra_namespaces: Vec::new(),
            root_attributes: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            word_prefixes: vec!["w".to_owned()],
        }
    }

    /// Get the combined text of all paragraphs.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Return supported watermark projections in document order.
    pub fn watermarks(&self) -> &[VmlWatermark] {
        &self.watermarks
    }

    /// Parse from XML bytes (the content of header*.xml or footer*.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let watermarks = parse_vml_watermarks(xml);
        let mut reader = Reader::from_reader(xml);
        // A raw child is captured with the text events of this reader, so
        // trimming them would drop the edge spaces of its text, such as the
        // one of "Page " before a page number. The text between the children
        // of the root is skipped below.
        reader.config_mut().trim_text(false);

        let mut paragraphs = Vec::new();
        let mut extra_namespaces = Vec::new();
        let mut root_attributes = Vec::new();
        let mut extra_xml = Vec::new();
        let mut extra_xml_positions = Vec::new();
        let mut buf = Vec::new();
        let mut word_prefixes = Vec::new();

        let known_ns: &[&[u8]] = &[b"xmlns:w", b"xmlns:r", b"xmlns"];

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"p", &prefixes) {
                        paragraphs.push(CT_P::from_xml_with_prefixes_and_root(
                            &mut reader,
                            &prefixes,
                            Some(e),
                        )?);
                    } else if is_word_element(name.as_ref(), b"hdr", &prefixes)
                        || is_word_element(name.as_ref(), b"ftr", &prefixes)
                    {
                        // Capture extra namespace declarations and the other
                        // attributes from root element
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            let is_namespace = key.starts_with(b"xmlns:") || key == b"xmlns";
                            if is_namespace && !known_ns.contains(&key) {
                                let key_str = std::str::from_utf8(key).unwrap_or("").to_string();
                                let val_str = attr
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        e.decoder(),
                                    )?
                                    .into_owned();
                                extra_namespaces.push((key_str, val_str));
                            } else if !is_namespace {
                                root_attributes.push((
                                    std::str::from_utf8(key)?.to_owned(),
                                    attr.decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        e.decoder(),
                                    )?
                                    .into_owned(),
                                ));
                            }
                        }
                        word_prefixes = prefixes;
                    } else {
                        // Capture unknown elements as raw XML
                        extra_xml.push(capture_element(&mut reader, e)?);
                        extra_xml_positions.push(paragraphs.len());
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if is_word_element(name.as_ref(), b"p", &prefixes) {
                        paragraphs.push(CT_P::from_empty_root(e, &prefixes)?);
                    } else if !matches_local_name(name.as_ref(), b"hdr")
                        && !matches_local_name(name.as_ref(), b"ftr")
                    {
                        extra_xml.push(capture_empty_element(e)?);
                        extra_xml_positions.push(paragraphs.len());
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_HdrFtr {
            paragraphs,
            watermarks,
            extra_namespaces,
            root_attributes,
            extra_xml,
            extra_xml_positions,
            word_prefixes,
        })
    }

    /// The number of direct `w:tbl` children, which the model keeps as raw
    /// XML between its paragraphs.
    pub fn table_count(&self) -> usize {
        self.table_entries().len()
    }

    /// Read the `index`-th direct table, or `None` when there is none.
    pub fn read_table<R>(
        &self,
        index: usize,
        read: impl FnOnce(&CT_Tbl) -> R,
    ) -> Result<Option<R>> {
        let Some(&entry) = self.table_entries().get(index) else {
            return Ok(None);
        };
        let (table, _) = parse_raw_table(&self.extra_xml[entry], &self.word_prefixes)?;
        Ok(Some(read(&table)))
    }

    /// Edit the `index`-th direct table in place, or return `None` when
    /// there is none. The table is written back with the namespaces its
    /// source declared, and an edit whose result would leave a prefix
    /// unbound is refused.
    pub fn edit_table<R>(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut CT_Tbl) -> R,
    ) -> Result<Option<R>> {
        let Some(&entry) = self.table_entries().get(index) else {
            return Ok(None);
        };
        let raw = &self.extra_xml[entry];
        let (mut table, bindings) = parse_raw_table(raw, &self.word_prefixes)?;
        let result = edit(&mut table);
        let mut writer = Writer::new(Vec::new());
        table.to_xml(&mut writer)?;
        let rewritten = crate::placeholder::with_source_namespaces(
            raw,
            &writer.into_inner(),
            &bindings,
            &self.word_prefixes,
        )
        .ok_or_else(|| {
            OxmlError::InvalidValue(
                "the header or footer table uses namespaces a rewrite cannot keep".into(),
            )
        })?;
        self.extra_xml[entry] = rewritten;
        Ok(Some(result))
    }

    /// Append a table after the last child of the story.
    pub fn push_table(&mut self, table: &CT_Tbl) -> Result<()> {
        let mut writer = Writer::new(Vec::new());
        table.to_xml(&mut writer)?;
        self.extra_xml.push(writer.into_inner());
        self.extra_xml_positions.push(self.paragraphs.len());
        Ok(())
    }

    /// The `extra_xml` indices of the direct `w:tbl` children, in order.
    fn table_entries(&self) -> Vec<usize> {
        self.extra_xml
            .iter()
            .enumerate()
            .filter(|(_, raw)| raw_word_element_is(raw, b"tbl", &self.word_prefixes))
            .map(|(index, _)| index)
            .collect()
    }

    /// Serialize to XML bytes as a header.
    pub fn to_xml_header(&self) -> Result<Vec<u8>> {
        self.to_xml_root("w:hdr")
    }

    /// Serialize to XML bytes as a footer.
    pub fn to_xml_footer(&self) -> Result<Vec<u8>> {
        self.to_xml_root("w:ftr")
    }

    fn to_xml_root(&self, root_tag: &str) -> Result<Vec<u8>> {
        let wp_ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        let wp_root_is_canonical = self
            .extra_namespaces
            .iter()
            .find(|(name, _)| name == "xmlns:wp")
            .is_none_or(|(_, namespace)| namespace == wp_ns);
        let _binding_scope = root_binding_scope(
            ROOT_R_BINDING
                | if wp_root_is_canonical {
                    ROOT_WP_BINDING
                } else {
                    0
                },
        );
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

        // Always emit xmlns:wp for drawing elements
        let mut has_wp = false;
        for (key, _) in &self.extra_namespaces {
            if key == "xmlns:wp" {
                has_wp = true;
                break;
            }
        }
        if !has_wp {
            start.push_attribute(("xmlns:wp", wp_ns));
        }

        // Replay captured extra namespaces, then the root attributes that
        // may name their prefixes, such as `mc:Ignorable`.
        for (key, val) in self.extra_namespaces.iter().chain(&self.root_attributes) {
            start.push_attribute((key.as_str(), val.as_str()));
        }

        writer.write_event(Event::Start(start))?;

        // Write each captured unknown element before the paragraph it
        // preceded, and the rest after the last paragraph.
        let mut raw_children = self
            .extra_xml
            .iter()
            .enumerate()
            .map(|(index, raw)| {
                let position = self.extra_xml_positions.get(index).copied();
                (position.unwrap_or(usize::MAX), raw)
            })
            .peekable();
        for (index, p) in self.paragraphs.iter().enumerate() {
            while let Some((_, raw)) = raw_children.next_if(|(position, _)| *position <= index) {
                writer.get_mut().extend_from_slice(raw);
            }
            p.to_xml(&mut writer)?;
        }
        for (_, raw) in raw_children {
            writer.get_mut().extend_from_slice(raw);
        }

        writer.write_event(Event::End(BytesEnd::new(root_tag)))?;

        let mut xml = writer.into_inner();
        declare_w14_on_part_root(&mut xml)?;
        Ok(xml)
    }
}

/// Whether a raw child is the Word element `local`.
fn raw_word_element_is(raw: &[u8], local: &[u8], part_prefixes: &[String]) -> bool {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    match reader.read_event_into(&mut buffer) {
        Ok(Event::Start(start) | Event::Empty(start)) => word_prefixes_at(&start, part_prefixes)
            .is_ok_and(|prefixes| is_word_element(start.name().as_ref(), local, &prefixes)),
        _ => false,
    }
}

/// Parse a raw `w:tbl` child, with the namespace bindings its start tag
/// declares other than the part does.
fn parse_raw_table(
    raw: &[u8],
    part_prefixes: &[String],
) -> Result<(CT_Tbl, Vec<(String, String)>)> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let Event::Start(start) = reader.read_event_into(&mut buffer)? else {
        return Err(OxmlError::InvalidValue(
            "a header or footer table has no content".into(),
        ));
    };
    let start = start.into_owned();
    let prefixes = word_prefixes_at(&start, part_prefixes)?;
    let bindings = local_namespace_overrides(&start, part_prefixes)?;
    let table =
        CT_Tbl::from_xml_with_prefixes_and_owner_bindings(&mut reader, &prefixes, &bindings)?;
    Ok((table, bindings))
}

/// Replace the exact API-owned VML shape without reconstructing its header.
#[doc(hidden)]
pub fn replace_authored_watermark(xml: &[u8], watermark: &VmlWatermark) -> Result<Vec<u8>> {
    let mut replacement_pict = watermark.to_pict_xml();
    let replacement_shape_type_id = match watermark {
        VmlWatermark::Text { .. } => TEXT_WATERMARK_SHAPE_TYPE_ID,
        VmlWatermark::Image { .. } => IMAGE_WATERMARK_SHAPE_TYPE_ID,
    };
    if contains_vml_shapetype(xml, replacement_shape_type_id)?
        && let Some(range) = vml_shape_type_range(&replacement_pict, replacement_shape_type_id)?
    {
        replacement_pict.drain(range);
    }
    let normalized_xml =
        prepare_owned_shape_type(xml, &replacement_pict, replacement_shape_type_id)?;
    let xml = normalized_xml.as_deref().unwrap_or(xml);
    let replacement_range = api_owned_shape_ranges(&replacement_pict)?
        .into_iter()
        .next()
        .ok_or_else(|| OxmlError::MissingElement("generated watermark shape".to_owned()))?;
    let mut replacement_shape = replacement_pict[replacement_range].to_vec();

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut pict_depth = None;
    let mut owned_start = None;
    let mut owned_ranges = Vec::new();
    let mut header_end = None;
    let mut empty_header = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        let is_word = namespace_is(&namespace, W_NS);
        let is_vml = namespace_is(&namespace, VML_NS);
        let mut completed_owner = None;
        let mut empty_owner = false;
        let mut completed_header = false;
        let mut empty_header_name = None;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                let name = element.name();
                let local = local_name(name.as_ref());
                if is_word && local == b"pict" {
                    pict_depth = Some(depth);
                } else if pict_depth.is_some()
                    && is_vml
                    && local == b"shape"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some("rdocx-watermark")
                {
                    owned_start = Some((depth, event_start));
                }
            }
            Event::Empty(ref element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                if depth == 0 && is_word && local == b"hdr" {
                    empty_header_name = Some(name.as_ref().to_vec());
                } else if pict_depth.is_some()
                    && is_vml
                    && local == b"shape"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some("rdocx-watermark")
                {
                    empty_owner = true;
                }
            }
            Event::End(ref element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                if owned_start.is_some_and(|(owner_depth, _)| owner_depth == depth)
                    && is_vml
                    && local == b"shape"
                    && let Some((_, start)) = owned_start.take()
                {
                    completed_owner = Some(start);
                }
                if pict_depth == Some(depth) && is_word && local == b"pict" {
                    pict_depth = None;
                }
                if depth == 1 && is_word && local == b"hdr" {
                    completed_header = true;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
        drop(event);
        drop(namespace);
        let event_end = reader.buffer_position() as usize;
        if let Some(start) = completed_owner {
            owned_ranges.push(start..event_end);
        }
        if empty_owner {
            owned_ranges.push(event_start..event_end);
        }
        if completed_header {
            header_end = Some(event_start);
        }
        if let Some(root_name) = empty_header_name {
            empty_header = Some((event_start..event_end, root_name));
        }
        buffer.clear();
    }

    if !owned_ranges.is_empty() {
        if !contains_vml_shapetype(xml, replacement_shape_type_id)? {
            let attribute = format!(" type=\"#{replacement_shape_type_id}\"");
            if let Some(position) = replacement_shape
                .windows(attribute.len())
                .position(|window| window == attribute.as_bytes())
            {
                replacement_shape.drain(position..position + attribute.len());
            }
        }
        replacement_shape = ensure_fixed_prefix_bindings(
            replacement_shape,
            xml,
            owned_ranges[0].start,
            matches!(watermark, VmlWatermark::Image { .. }),
        )?;
        let mut output = Vec::with_capacity(xml.len() + replacement_shape.len());
        let mut copied = 0usize;
        for (index, range) in owned_ranges.into_iter().enumerate() {
            output.extend_from_slice(&xml[copied..range.start]);
            if index == 0 {
                output.extend_from_slice(&replacement_shape);
            }
            copied = range.end;
        }
        output.extend_from_slice(&xml[copied..]);
        return Ok(output);
    }

    let mut paragraph = format!("<w:p xmlns:w=\"{W_NS}\"><w:r>").into_bytes();
    paragraph.extend_from_slice(&replacement_pict);
    paragraph.extend_from_slice(b"</w:r></w:p>");
    if let Some(position) = header_end {
        let mut output = Vec::with_capacity(xml.len() + paragraph.len());
        output.extend_from_slice(&xml[..position]);
        output.extend_from_slice(&paragraph);
        output.extend_from_slice(&xml[position..]);
        return Ok(output);
    }
    if let Some((range, root_name)) = empty_header {
        let empty = &xml[range.clone()];
        let close = empty
            .windows(2)
            .rposition(|window| window == b"/>")
            .ok_or_else(|| OxmlError::InvalidValue("empty header root".to_owned()))?;
        let mut output = Vec::with_capacity(xml.len() + paragraph.len() + root_name.len() + 2);
        output.extend_from_slice(&xml[..range.start]);
        output.extend_from_slice(&empty[..close]);
        output.push(b'>');
        output.extend_from_slice(&paragraph);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&root_name);
        output.push(b'>');
        output.extend_from_slice(&xml[range.end..]);
        return Ok(output);
    }
    Err(OxmlError::MissingElement("header root".to_owned()))
}

fn api_owned_shape_ranges(xml: &[u8]) -> Result<Vec<std::ops::Range<usize>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut pict_depth = None;
    let mut owned_start = None;
    let mut ranges = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        let is_word = namespace_is(&namespace, W_NS);
        let is_vml = namespace_is(&namespace, VML_NS);
        let mut completed_owner = None;
        let mut empty_owner = false;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                let name = element.name();
                let local = local_name(name.as_ref());
                if is_word && local == b"pict" {
                    pict_depth = Some(depth);
                } else if pict_depth.is_some()
                    && is_vml
                    && local == b"shape"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some("rdocx-watermark")
                {
                    owned_start = Some((depth, event_start));
                }
            }
            Event::Empty(ref element)
                if pict_depth.is_some()
                    && is_vml
                    && local_name(element.name().as_ref()) == b"shape"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some("rdocx-watermark") =>
            {
                empty_owner = true;
            }
            Event::End(ref element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                if owned_start.is_some_and(|(owner_depth, _)| owner_depth == depth)
                    && is_vml
                    && local == b"shape"
                    && let Some((_, start)) = owned_start.take()
                {
                    completed_owner = Some(start);
                }
                if pict_depth == Some(depth) && is_word && local == b"pict" {
                    pict_depth = None;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
        drop(event);
        drop(namespace);
        let event_end = reader.buffer_position() as usize;
        if let Some(start) = completed_owner {
            ranges.push(start..event_end);
        }
        if empty_owner {
            ranges.push(event_start..event_end);
        }
        buffer.clear();
    }
    Ok(ranges)
}

fn contains_vml_shapetype(xml: &[u8], expected_id: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shapetype"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some(expected_id) =>
            {
                return Ok(true);
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn prepare_owned_shape_type(
    xml: &[u8],
    replacement_pict: &[u8],
    replacement_id: &str,
) -> Result<Option<Vec<u8>>> {
    let Some(owned_type) = api_owned_shape_type(xml)? else {
        return Ok(None);
    };
    if owned_type.as_deref() == Some(replacement_id) && contains_vml_shapetype(xml, replacement_id)?
    {
        return Ok(None);
    }
    let replacement = if contains_vml_shapetype(xml, replacement_id)? {
        None
    } else {
        let replacement_range = vml_shape_type_range(replacement_pict, replacement_id)?
            .ok_or_else(|| OxmlError::MissingElement("generated watermark shapetype".to_owned()))?;
        Some(replacement_pict[replacement_range].to_vec())
    };
    let old_definition_is_owned = owned_type
        .as_deref()
        .is_some_and(|owned_id| owned_id.starts_with(WATERMARK_SHAPE_TYPE_PREFIX));
    let old_definition_is_shared = match owned_type.as_deref() {
        Some(owned_id) => shape_type_is_used_by_producer_shape(xml, owned_id)?,
        None => false,
    };
    if !old_definition_is_owned || old_definition_is_shared {
        let Some(replacement) = replacement else {
            return Ok(None);
        };
        let insertion = api_owned_shape_ranges(xml)?
            .into_iter()
            .next()
            .ok_or_else(|| OxmlError::MissingElement("owned watermark shape".to_owned()))?
            .start;
        let replacement = ensure_fixed_prefix_bindings(replacement, xml, insertion, false)?;
        let mut output = Vec::with_capacity(xml.len() + replacement.len());
        output.extend_from_slice(&xml[..insertion]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&xml[insertion..]);
        return Ok(Some(output));
    }
    let Some(source_range) = vml_shape_type_range(
        xml,
        owned_type.as_deref().expect("owned definition id exists"),
    )?
    else {
        let Some(replacement) = replacement else {
            return Ok(None);
        };
        let insertion = api_owned_shape_ranges(xml)?
            .into_iter()
            .next()
            .ok_or_else(|| OxmlError::MissingElement("owned watermark shape".to_owned()))?
            .start;
        let replacement = ensure_fixed_prefix_bindings(replacement, xml, insertion, false)?;
        let mut output = Vec::with_capacity(xml.len() + replacement.len());
        output.extend_from_slice(&xml[..insertion]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&xml[insertion..]);
        return Ok(Some(output));
    };
    let replacement = replacement
        .map(|replacement| {
            ensure_fixed_prefix_bindings(replacement, xml, source_range.start, false)
        })
        .transpose()?
        .unwrap_or_default();
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..source_range.start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[source_range.end..]);
    Ok(Some(output))
}

fn ensure_fixed_prefix_bindings(
    mut fragment: Vec<u8>,
    source: &[u8],
    offset: usize,
    needs_relationships: bool,
) -> Result<Vec<u8>> {
    let bindings = ancestor_namespace_bindings(source, offset)?;
    let mut declarations = Vec::new();
    for (prefix, namespace) in [("v", VML_NS), ("o", OFFICE_NS)] {
        if !bindings.iter().any(|(bound_prefix, bound_namespace)| {
            bound_prefix == prefix && bound_namespace == namespace
        }) {
            declarations.push((prefix, namespace));
        }
    }
    if needs_relationships
        && !bindings
            .iter()
            .any(|(prefix, namespace)| prefix == "r" && namespace == RELATIONSHIPS_NS)
    {
        declarations.push(("r", RELATIONSHIPS_NS));
    }
    if declarations.is_empty() {
        return Ok(fragment);
    }
    let name_end = fragment
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'>' | b'/'))
        .ok_or_else(|| OxmlError::InvalidValue("watermark XML fragment".to_owned()))?;
    let mut attributes = Vec::new();
    for (prefix, namespace) in declarations {
        attributes.extend_from_slice(format!(" xmlns:{prefix}=\"{namespace}\"").as_bytes());
    }
    fragment.splice(name_end..name_end, attributes);
    Ok(fragment)
}

fn ancestor_namespace_bindings(source: &[u8], offset: usize) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes = vec![Vec::new()];
    loop {
        if reader.buffer_position() as usize == offset {
            return Ok(namespace_bindings(
                scopes.last().expect("namespace root scope exists"),
            ));
        }
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                let inherited = scopes.last().expect("namespace root scope exists");
                scopes.push(word_prefixes_at(&start, inherited)?);
            }
            Event::End(_) => {
                scopes.pop().ok_or_else(|| {
                    OxmlError::InvalidValue("unbalanced watermark XML namespace scope".to_owned())
                })?;
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "watermark XML splice point".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn api_owned_shape_type(xml: &[u8]) -> Result<Option<Option<String>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut pict_depth = None;
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                if namespace_is(&namespace, W_NS) && local_name(element.name().as_ref()) == b"pict"
                {
                    pict_depth = Some(depth);
                } else if pict_depth.is_some()
                    && namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shape"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some("rdocx-watermark")
                {
                    return Ok(Some(
                        raw_unqualified_attribute(element, b"type")
                            .and_then(|value| value.strip_prefix('#').map(str::to_owned)),
                    ));
                }
            }
            Event::Empty(ref element)
                if pict_depth.is_some()
                    && namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shape"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some("rdocx-watermark") =>
            {
                return Ok(Some(
                    raw_unqualified_attribute(element, b"type")
                        .and_then(|value| value.strip_prefix('#').map(str::to_owned)),
                ));
            }
            Event::End(ref element) => {
                if pict_depth == Some(depth)
                    && namespace_is(&namespace, W_NS)
                    && local_name(element.name().as_ref()) == b"pict"
                {
                    pict_depth = None;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn shape_type_is_used_by_producer_shape(xml: &[u8], expected_id: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut pict_depth = None;
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                if namespace_is(&namespace, W_NS) && local_name(element.name().as_ref()) == b"pict"
                {
                    pict_depth = Some(depth);
                } else if namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shape"
                    && raw_unqualified_attribute(element, b"type")
                        .as_deref()
                        .and_then(|value| value.strip_prefix('#'))
                        == Some(expected_id)
                    && !(pict_depth.is_some()
                        && raw_unqualified_attribute(element, b"id").as_deref()
                            == Some("rdocx-watermark"))
                {
                    return Ok(true);
                }
            }
            Event::Empty(ref element)
                if namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shape"
                    && raw_unqualified_attribute(element, b"type")
                        .as_deref()
                        .and_then(|value| value.strip_prefix('#'))
                        == Some(expected_id)
                    && !(pict_depth.is_some()
                        && raw_unqualified_attribute(element, b"id").as_deref()
                            == Some("rdocx-watermark")) =>
            {
                return Ok(true);
            }
            Event::End(ref element) => {
                if pict_depth == Some(depth)
                    && namespace_is(&namespace, W_NS)
                    && local_name(element.name().as_ref()) == b"pict"
                {
                    pict_depth = None;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn vml_shape_type_range(xml: &[u8], expected_id: &str) -> Result<Option<std::ops::Range<usize>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        let mut completed = None;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                if namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shapetype"
                    && raw_unqualified_attribute(element, b"id").as_deref() == Some(expected_id)
                {
                    start = Some((depth, event_start));
                }
            }
            Event::Empty(ref element)
                if namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shapetype"
                    && raw_unqualified_attribute(element, b"id").as_deref()
                        == Some(expected_id) =>
            {
                return Ok(Some(event_start..reader.buffer_position() as usize));
            }
            Event::End(ref element) => {
                if start.is_some_and(|(start_depth, _)| start_depth == depth)
                    && namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shapetype"
                    && let Some((_, range_start)) = start.take()
                {
                    completed = Some(range_start..reader.buffer_position() as usize);
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        if let Some(range) = completed {
            return Ok(Some(range));
        }
        buffer.clear();
    }
}

impl Default for CT_HdrFtr {
    fn default() -> Self {
        Self::new()
    }
}

fn parse_vml_watermark(xml: &[u8]) -> Option<VmlWatermark> {
    parse_vml_watermarks(xml).into_iter().next()
}

#[derive(Default)]
struct PendingWatermark {
    width: Option<f64>,
    height: Option<f64>,
    rotation: f64,
    color: Option<String>,
    opacity: f64,
    text: Option<String>,
    font_family: Option<String>,
    relationship_id: Option<String>,
}

impl PendingWatermark {
    fn finish(self) -> Option<VmlWatermark> {
        let width_pt = self.width?;
        let height_pt = self.height?;
        if !width_pt.is_finite()
            || !height_pt.is_finite()
            || !self.rotation.is_finite()
            || !self.opacity.is_finite()
            || width_pt <= 0.0
            || height_pt <= 0.0
        {
            return None;
        }
        let opacity = self.opacity.clamp(0.0, 1.0);
        match (self.text, self.relationship_id) {
            (Some(text), None) => Some(VmlWatermark::Text {
                text,
                width_pt,
                height_pt,
                rotation_degrees: self.rotation,
                color: self.color.unwrap_or_else(|| "D9D9D9".to_owned()),
                font_family: self.font_family,
                opacity,
            }),
            (None, Some(relationship_id)) => Some(VmlWatermark::Image {
                relationship_id,
                width_pt,
                height_pt,
                rotation_degrees: self.rotation,
                opacity,
            }),
            _ => None,
        }
    }
}

fn parse_vml_watermarks(xml: &[u8]) -> Vec<VmlWatermark> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut in_pict = false;
    let mut pending = None;
    let mut watermarks = Vec::new();

    loop {
        let Ok((namespace, event)) = reader.read_resolved_event_into(&mut buffer) else {
            return Vec::new();
        };
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = element.name();
                let local = local_name(name.as_ref());
                if namespace_is(&namespace, W_NS) && local == b"pict" {
                    in_pict = true;
                } else if in_pict && namespace_is(&namespace, VML_NS) && local == b"shape" {
                    let Some(style) = unqualified_attribute(&reader, element, b"style") else {
                        buffer.clear();
                        continue;
                    };
                    pending = Some(PendingWatermark {
                        width: style_number(&style, "width", "pt"),
                        height: style_number(&style, "height", "pt"),
                        rotation: style_number(&style, "rotation", "").unwrap_or(0.0),
                        color: unqualified_attribute(&reader, element, b"fillcolor")
                            .map(|value| value.trim_start_matches('#').to_owned()),
                        opacity: style_number(&style, "opacity", "").unwrap_or(1.0),
                        ..PendingWatermark::default()
                    });
                } else if let Some(shape) = pending.as_mut()
                    && namespace_is(&namespace, VML_NS)
                    && local == b"fill"
                {
                    shape.opacity = unqualified_attribute(&reader, element, b"opacity")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(shape.opacity);
                } else if let Some(shape) = pending.as_mut()
                    && namespace_is(&namespace, VML_NS)
                    && local == b"textpath"
                {
                    shape.text = unqualified_attribute(&reader, element, b"string");
                    shape.font_family = unqualified_attribute(&reader, element, b"style")
                        .and_then(|style| style_value(&style, "font-family"))
                        .map(|family| family.trim_matches(['\'', '"']).to_owned());
                } else if let Some(shape) = pending.as_mut()
                    && namespace_is(&namespace, VML_NS)
                    && local == b"imagedata"
                {
                    shape.relationship_id =
                        namespaced_attribute(&reader, element, RELATIONSHIPS_NS, b"id");
                }
            }
            Event::End(ref element)
                if namespace_is(&namespace, VML_NS)
                    && local_name(element.name().as_ref()) == b"shape" =>
            {
                if let Some(watermark) = pending.take().and_then(PendingWatermark::finish) {
                    watermarks.push(watermark);
                }
            }
            Event::End(ref element)
                if namespace_is(&namespace, W_NS)
                    && local_name(element.name().as_ref()) == b"pict" =>
            {
                in_pict = false;
                pending = None;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    watermarks
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn namespace_is(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected.as_bytes())
}

fn unqualified_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == expected {
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .ok()
                .map(|value| value.into_owned())
        } else {
            None
        }
    })
}

fn raw_unqualified_attribute(element: &BytesStart<'_>, expected: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.as_ref() == expected)
            .then(|| {
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .ok()
                    .map(|value| value.into_owned())
            })
            .flatten()
    })
}

fn namespaced_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &str,
    expected_local: &[u8],
) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_is(&namespace, expected_namespace) && local.as_ref() == expected_local {
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .ok()
                .map(|value| value.into_owned())
        } else {
            None
        }
    })
}

fn style_value(style: &str, expected: &str) -> Option<String> {
    style.split(';').find_map(|declaration| {
        let (name, value) = declaration.split_once(':')?;
        name.trim()
            .eq_ignore_ascii_case(expected)
            .then(|| value.trim().to_owned())
    })
}

fn style_number(style: &str, name: &str, suffix: &str) -> Option<f64> {
    style_value(style, name)?
        .strip_suffix(suffix)?
        .trim()
        .parse()
        .ok()
}

fn compact_number(value: f64) -> String {
    let mut value = format!("{value:.6}");
    while value.ends_with('0') {
        value.pop();
    }
    if value.ends_with('.') {
        value.pop();
    }
    value
}

/// Header/footer reference type in section properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HdrFtrType {
    /// Default header/footer
    Default,
    /// First page header/footer
    First,
    /// Even page header/footer
    Even,
}

impl HdrFtrType {
    pub fn from_str(s: &str) -> Self {
        match s {
            "first" => Self::First,
            "even" => Self::Even,
            _ => Self::Default,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::First => "first",
            Self::Even => "even",
        }
    }
}

/// A header or footer reference (stored in section properties).
#[derive(Debug, Clone, PartialEq)]
pub struct HdrFtrRef {
    /// The type (default, first, even)
    pub hdr_ftr_type: HdrFtrType,
    /// Relationship ID
    pub rel_id: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_paragraph_identity_stays_bound_when_only_the_paragraph_declares_w14() {
        let w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}"><w:p xmlns:w14="{w14}" w14:paraId="1A2B3C4D"><w:r><w:t>header</w:t></w:r></w:p></w:hdr>"#
        );
        let header = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        let output = String::from_utf8(header.to_xml_header().unwrap()).unwrap();
        let root = &output[output.find("<w:hdr").unwrap()..];
        let root = &root[..root.find('>').unwrap()];
        assert!(root.contains(&format!(r#"xmlns:w14="{w14}""#)), "{output}");
        assert!(output.contains(r#"w14:paraId="1A2B3C4D""#), "{output}");
    }

    #[test]
    fn a_rewritten_header_is_compact_and_declares_w_once() {
        // Word writes a header without indentation and puts `w:rsid*` on
        // nearly every paragraph and run, bound by the root alone.
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}"><w:p w:rsidR="00A1B2C3"><w:r w:rsidRPr="00A1B2C4"><w:t xml:space="preserve">header </w:t></w:r></w:p></w:hdr>"#
        );
        let header = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        let output = String::from_utf8(header.to_xml_header().unwrap()).unwrap();
        assert_eq!(output.matches("xmlns:w=").count(), 1, "{output}");
        assert!(!output.contains('\n'), "{output}");
        assert!(
            output.contains(r#"<w:p w:rsidR="00A1B2C3"><w:r w:rsidRPr="00A1B2C4"><w:t xml:space="preserve">header </w:t></w:r></w:p>"#),
            "{output}"
        );
        assert_eq!(
            CT_HdrFtr::from_xml(output.as_bytes()).unwrap().paragraphs,
            header.paragraphs
        );
    }

    #[test]
    fn round_trip_header() {
        let mut hdr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        p.add_run("Page Header");
        hdr.paragraphs.push(p);

        let xml = hdr.to_xml_header().unwrap();
        let parsed = CT_HdrFtr::from_xml(&xml).unwrap();
        assert_eq!(parsed.paragraphs.len(), 1);
        assert_eq!(parsed.text(), "Page Header");
    }

    #[test]
    fn round_trip_footer() {
        let mut ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        p.add_run("Page Footer");
        ftr.paragraphs.push(p);

        let xml = ftr.to_xml_footer().unwrap();
        let parsed = CT_HdrFtr::from_xml(&xml).unwrap();
        assert_eq!(parsed.text(), "Page Footer");
    }

    #[test]
    fn empty_header() {
        let hdr = CT_HdrFtr::new();
        let xml = hdr.to_xml_header().unwrap();
        let parsed = CT_HdrFtr::from_xml(&xml).unwrap();
        assert_eq!(parsed.paragraphs.len(), 0);
    }

    /// A rewrite wrote every table and content control after the last
    /// paragraph. A raw child a caller adds still goes after it.
    #[test]
    fn raw_children_keep_their_place_between_paragraphs() {
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}"><w:tbl><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:t>one</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt><w:p><w:r><w:t>two</w:t></w:r></w:p><w:bookmarkStart w:id="0" w:name="end"/></w:hdr>"#
        );
        let mut parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        parsed
            .extra_xml
            .push(br#"<w:bookmarkEnd w:id="0"/>"#.to_vec());

        let written = String::from_utf8(parsed.to_xml_header().unwrap()).unwrap();

        let positions = [
            "<w:tbl>",
            ">one<",
            "<w:sdt>",
            ">two<",
            "<w:bookmarkStart",
            "<w:bookmarkEnd",
        ]
        .map(|marker| written.find(marker).unwrap());
        assert!(positions.is_sorted(), "{written}");
    }

    /// #304: a table is edited where it stands, among the raw children, with
    /// the prefix its source used and the children around it in place.
    fn first_cell_text(table: &CT_Tbl) -> String {
        table.rows[0].cells[0].text()
    }

    #[test]
    fn a_direct_table_is_read_edited_and_appended_in_place() {
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>one</w:t></w:r></w:p><x:tbl xmlns:x="{W_NS}"><x:tblGrid><x:gridCol x:w="100"/></x:tblGrid><x:tr><x:tc><x:p><x:r><x:t>cell</x:t></x:r></x:p></x:tc></x:tr></x:tbl><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt><w:p><w:r><w:t>two</w:t></w:r></w:p></w:hdr>"#
        );
        let mut header = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(header.table_count(), 1);
        assert_eq!(
            header.read_table(0, first_cell_text).unwrap().as_deref(),
            Some("cell")
        );
        assert!(header.read_table(1, first_cell_text).unwrap().is_none());
        header
            .edit_table(0, |table| {
                table.rows[0].cells[0].paragraphs_mut()[0].add_run(" edited");
            })
            .unwrap()
            .unwrap();
        let mut appended = CT_Tbl::new();
        appended.rows.push(crate::table::CT_Row::new());
        header.push_table(&appended).unwrap();
        assert_eq!(header.table_count(), 2);

        let written = String::from_utf8(header.to_xml_header().unwrap()).unwrap();
        let positions = [
            ">one<",
            "cell",
            " edited",
            "<w:sdt>",
            ">two<",
            "<w:tbl><w:tr>",
        ]
        .map(|marker| written.find(marker).unwrap());
        assert!(positions.is_sorted(), "{written}");
        let reopened = CT_HdrFtr::from_xml(written.as_bytes()).unwrap();
        assert_eq!(
            reopened.read_table(0, first_cell_text).unwrap().as_deref(),
            Some("cell edited")
        );
    }

    /// A raw child was captured with trimmed text events, so the page-number
    /// control of a footer came back as "Page" once the part was rewritten.
    #[test]
    fn a_raw_child_keeps_the_edge_spaces_of_its_text() {
        let control = r#"<w:sdt><w:sdtContent><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "/></w:p></w:sdtContent></w:sdt>"#;
        let xml = format!("<w:ftr xmlns:w=\"{W_NS}\">\n  {control}\n  <w:p/>\n</w:ftr>");

        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();

        assert_eq!(parsed.extra_xml, [control.as_bytes()]);
        assert_eq!(parsed.paragraphs.len(), 1);
    }

    #[test]
    fn aliased_header_paragraph_properties_keep_root_scope() {
        let xml = format!(
            r#"<q:hdr xmlns:q="{W_NS}" xmlns:ext="urn:producer"><ext:p><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:p><q:p><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr><q:r><q:t>Header</q:t></q:r></q:p></q:hdr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(parsed.paragraphs.len(), 1);
        assert_eq!(parsed.text(), "Header");
        assert_eq!(
            parsed.paragraphs[0].properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn default_namespace_header_properties_keep_root_scope() {
        let xml = format!(
            r#"<hdr xmlns="{W_NS}" xmlns:w="{W_NS}" xmlns:ext="urn:producer"><ext:p><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:p><p><pPr><ext:jc ext:val="right"/><jc w:val="center"/></pPr><r><t>Header</t></r></p></hdr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(parsed.paragraphs.len(), 1);
        assert_eq!(parsed.text(), "Header");
        assert_eq!(
            parsed.paragraphs[0].properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn word_vml_watermarks_parse_and_preserve_source_bytes() {
        let text_pict = r##"<q:pict><x:shape style="width:468pt;height:117pt;rotation:315" fillcolor="#D9D9D9"><x:textpath string="DRAFT" style="font-family:&quot;Calibri&quot;"/></x:shape></q:pict>"##;
        let image_pict = r#"<q:pict><x:shape style="width:72pt;height:36pt;rotation:0"><x:fill opacity=".25"/><x:imagedata rel:id="rId7"/></x:shape></q:pict>"#;
        let ordinary_pict = r#"<q:pict><x:shape id="ordinary"><x:path/></x:shape></q:pict>"#;
        let xml = format!(
            r#"<q:hdr xmlns:q="{W_NS}" xmlns:x="{VML_NS}" xmlns:rel="{RELATIONSHIPS_NS}"><q:p><q:r>{text_pict}{image_pict}{ordinary_pict}</q:r></q:p></q:hdr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert!(matches!(
            &parsed.watermarks()[0],
            VmlWatermark::Text {
                text,
                width_pt: 468.0,
                height_pt: 117.0,
                rotation_degrees: 315.0,
                color,
                font_family: Some(font_family),
                opacity: 1.0,
            } if text == "DRAFT" && color == "D9D9D9" && font_family == "Calibri"
        ));
        assert!(matches!(
            &parsed.watermarks()[1],
            VmlWatermark::Image {
                relationship_id,
                width_pt: 72.0,
                height_pt: 36.0,
                rotation_degrees: 0.0,
                opacity: 0.25,
            } if relationship_id == "rId7"
        ));
        let serialized = parsed.to_xml_header().unwrap();
        for source in [text_pict, image_pict, ordinary_pict] {
            assert!(
                serialized
                    .windows(source.len())
                    .any(|window| window == source.as_bytes()),
                "{}",
                String::from_utf8_lossy(&serialized)
            );
        }
    }

    #[test]
    fn generated_watermarks_write_fixed_prefixes_and_vml_child_order() {
        let watermark = VmlWatermark::Text {
            text: "DRAFT".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };
        let xml = watermark.to_pict_xml();
        let fill = xml
            .windows(b"<v:fill".len())
            .position(|w| w == b"<v:fill")
            .unwrap();
        let textpath = xml
            .windows(b"<v:textpath".len())
            .rposition(|w| w == b"<v:textpath")
            .unwrap();
        assert!(fill < textpath);
        let shape_type = xml
            .windows(b"<v:shapetype".len())
            .position(|window| window == b"<v:shapetype")
            .unwrap();
        let shape = xml
            .windows(b"<v:shape ".len())
            .position(|window| window == b"<v:shape ")
            .unwrap();
        assert!(shape_type < shape);
        assert!(
            xml.windows(b"id=\"rdocx-watermark-type-text\"".len())
                .any(|window| { window == b"id=\"rdocx-watermark-type-text\"" })
        );
        assert!(
            xml.windows(b"type=\"#rdocx-watermark-type-text\"".len())
                .any(|window| { window == b"type=\"#rdocx-watermark-type-text\"" })
        );
        assert!(xml.starts_with(
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#
        ));
        assert_eq!(VmlWatermark::from_pict_xml(&xml), Some(watermark.clone()));

        let image = VmlWatermark::Image {
            relationship_id: "rId3".to_owned(),
            width_pt: 72.0,
            height_pt: 36.0,
            rotation_degrees: 0.0,
            opacity: 0.5,
        };
        let image_xml = image.to_pict_xml();
        let fill = image_xml
            .windows(b"<v:fill".len())
            .position(|w| w == b"<v:fill")
            .unwrap();
        let image_data = image_xml
            .windows(b"<v:imagedata".len())
            .position(|w| w == b"<v:imagedata")
            .unwrap();
        assert!(fill < image_data);
        assert!(
            image_xml
                .windows(b"id=\"rdocx-watermark-type-image\"".len())
                .any(|window| window == b"id=\"rdocx-watermark-type-image\"")
        );
        assert!(
            image_xml
                .windows(b"type=\"#rdocx-watermark-type-image\"".len())
                .any(|window| window == b"type=\"#rdocx-watermark-type-image\"")
        );
        assert_eq!(VmlWatermark::from_pict_xml(&image_xml), Some(image));

        let legacy = format!(
            r##"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}"><w:p><w:r><w:pict><v:shape id="rdocx-watermark" type="#_x0000_t136" style="width:468pt;height:117pt"><v:textpath string="OLD"/></v:shape></w:pict></w:r></w:p></w:hdr>"##
        );
        let replaced = replace_authored_watermark(legacy.as_bytes(), &watermark).unwrap();
        assert!(
            !replaced
                .windows(b"type=\"#_x0000_t136\"".len())
                .any(|window| window == b"type=\"#_x0000_t136\"")
        );
        assert_eq!(
            CT_HdrFtr::from_xml(&replaced).unwrap().watermarks(),
            &[watermark]
        );
    }

    #[test]
    fn replacing_owned_watermark_keeps_a_producer_shared_shape_type() {
        let producer = r##"<v:shape id="producer" type="#_x0000_t75"><v:path/></v:shape>"##;
        let source = format!(
            r##"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}" xmlns:o="{OFFICE_NS}"><w:p><w:r><w:pict><v:shapetype id="_x0000_t75"><v:path/></v:shapetype>{producer}<v:shape id="rdocx-watermark" type="#_x0000_t75"><v:imagedata/></v:shape></w:pict></w:r></w:p></w:hdr>"##
        );
        let text = VmlWatermark::Text {
            text: "DRAFT".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };

        let replaced = replace_authored_watermark(source.as_bytes(), &text).unwrap();
        let replaced = String::from_utf8(replaced).unwrap();
        assert!(replaced.contains(producer));
        assert!(replaced.contains(r#"<v:shapetype id="_x0000_t75""#));
        assert!(replaced.contains(r#"<v:shapetype id="rdocx-watermark-type-text""#));
        assert!(replaced.contains(r##"type="#rdocx-watermark-type-text""##));
    }

    #[test]
    fn replacing_owned_watermark_preserves_an_unused_producer_shape_type() {
        let producer_type = r#"<v:shapetype id="_x0000_t75"><v:path/></v:shapetype>"#;
        let source = format!(
            r##"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}"><w:p><w:r><w:pict>{producer_type}<v:shape id="rdocx-watermark" type="#_x0000_t75"><v:imagedata/></v:shape></w:pict></w:r></w:p></w:hdr>"##
        );
        let text = VmlWatermark::Text {
            text: "DRAFT".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };

        let replaced =
            String::from_utf8(replace_authored_watermark(source.as_bytes(), &text).unwrap())
                .unwrap();
        assert!(replaced.contains(producer_type));
        assert_eq!(replaced.matches(r#"id="_x0000_t75""#).count(), 1);
        assert_eq!(
            replaced
                .matches(r#"id="rdocx-watermark-type-text""#)
                .count(),
            1
        );
    }

    #[test]
    fn first_watermark_insertion_reuses_an_existing_target_shape_type() {
        let producer_standard_type = r#"<v:shapetype id="_x0000_t136"><v:path/></v:shapetype>"#;
        let producer_target_type =
            r#"<v:shapetype id="rdocx-watermark-type-text"><v:path/></v:shapetype>"#;
        let source = format!(
            r#"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}"><w:p><w:r><w:pict>{producer_standard_type}{producer_target_type}</w:pict></w:r></w:p></w:hdr>"#
        );
        let text = VmlWatermark::Text {
            text: "DRAFT".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };

        let updated =
            String::from_utf8(replace_authored_watermark(source.as_bytes(), &text).unwrap())
                .unwrap();
        assert!(updated.contains(producer_standard_type));
        assert!(updated.contains(producer_target_type));
        assert_eq!(updated.matches(r#"id="_x0000_t136""#).count(), 1);
        assert_eq!(
            updated.matches(r#"id="rdocx-watermark-type-text""#).count(),
            1
        );
        assert!(updated.contains(r##"type="#rdocx-watermark-type-text""##));
    }

    #[test]
    fn watermark_replacement_binds_fixed_prefixes_at_an_alias_prefixed_splice() {
        let source = format!(
            r##"<w:hdr xmlns:w="{W_NS}" xmlns:x="{VML_NS}"><w:p><w:r><w:pict><x:shapetype id="producer"><x:path/></x:shapetype><x:shape id="rdocx-watermark" type="#producer"><x:imagedata/></x:shape></w:pict></w:r></w:p></w:hdr>"##
        );
        let image = VmlWatermark::Image {
            relationship_id: "rId3".to_owned(),
            width_pt: 72.0,
            height_pt: 36.0,
            rotation_degrees: 0.0,
            opacity: 0.5,
        };

        let updated = replace_authored_watermark(source.as_bytes(), &image).unwrap();
        let updated_text = String::from_utf8(updated.clone()).unwrap();
        assert!(updated_text.contains(&format!(r#"xmlns:v="{VML_NS}""#)));
        assert!(updated_text.contains(&format!(r#"xmlns:o="{OFFICE_NS}""#)));
        assert!(updated_text.contains(&format!(r#"xmlns:r="{RELATIONSHIPS_NS}""#)));
        assert_eq!(
            CT_HdrFtr::from_xml(&updated).unwrap().watermarks(),
            &[image],
            "{updated_text}"
        );
    }

    #[test]
    fn watermark_replacement_does_not_inherit_bindings_from_removed_elements() {
        let source = format!(
            r##"<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:pict><v:shapetype xmlns:v="{VML_NS}" xmlns:o="{OFFICE_NS}" id="{TEXT_WATERMARK_SHAPE_TYPE_ID}" o:spt="136"><v:path/></v:shapetype><v:shape xmlns:v="{VML_NS}" xmlns:o="{OFFICE_NS}" xmlns:r="{RELATIONSHIPS_NS}" id="rdocx-watermark" o:spid="_x0000_s1025" type="#{TEXT_WATERMARK_SHAPE_TYPE_ID}" style="width:468pt;height:117pt"><v:textpath string="OLD"/></v:shape></w:pict></w:r></w:p></w:hdr>"##
        );
        let image = VmlWatermark::Image {
            relationship_id: "rId3".to_owned(),
            width_pt: 72.0,
            height_pt: 36.0,
            rotation_degrees: 0.0,
            opacity: 0.5,
        };

        let updated = replace_authored_watermark(source.as_bytes(), &image).unwrap();
        let updated_text = String::from_utf8(updated.clone()).unwrap();
        assert!(updated_text.contains(&format!(r#"xmlns:v="{VML_NS}""#)));
        assert!(updated_text.contains(&format!(r#"xmlns:o="{OFFICE_NS}""#)));
        assert!(updated_text.contains(&format!(r#"xmlns:r="{RELATIONSHIPS_NS}""#)));
        assert_eq!(
            CT_HdrFtr::from_xml(&updated).unwrap().watermarks(),
            &[image]
        );
    }

    /// #160: a typed rewrite dropped `mc:Ignorable` from the part root.
    #[test]
    fn root_attributes_survive_a_rewrite_after_the_namespace_declarations() {
        let xml = format!(
            r#"<w:ftr xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="{W_NS}" mc:Ignorable="w14" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p><w:r><w:t>Page</w:t></w:r></w:p></w:ftr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        let written = String::from_utf8(parsed.to_xml_footer().unwrap()).unwrap();
        assert!(
            written.contains(
                r#" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">"#
            ),
            "{written}"
        );
        let reparsed = CT_HdrFtr::from_xml(written.as_bytes()).unwrap();
        assert_eq!(reparsed.root_attributes, parsed.root_attributes);
        assert_eq!(reparsed.to_xml_footer().unwrap(), written.as_bytes());
    }

    /// A declaration value is read unescaped, as `CT_Document` reads its own,
    /// so a rewrite escapes it once and still binds the same namespace.
    #[test]
    fn root_namespace_values_are_unescaped_once_through_a_rewrite() {
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}" xmlns:x="urn:a&amp;b?c=&quot;1&quot;" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert!(
            parsed
                .extra_namespaces
                .contains(&("xmlns:x".to_owned(), r#"urn:a&b?c="1""#.to_owned())),
            "{:?}",
            parsed.extra_namespaces
        );
        let written = String::from_utf8(parsed.to_xml_header().unwrap()).unwrap();
        assert!(
            written.contains(r#" xmlns:x="urn:a&amp;b?c=&quot;1&quot;""#),
            "{written}"
        );
        let reparsed = CT_HdrFtr::from_xml(written.as_bytes()).unwrap();
        assert!(
            reparsed
                .extra_namespaces
                .contains(&("xmlns:x".to_owned(), r#"urn:a&b?c="1""#.to_owned()))
        );
        assert_eq!(reparsed.to_xml_header().unwrap(), written.as_bytes());
    }

    #[test]
    fn non_watermark_w_pict_remains_opaque() {
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape id="ordinary"><v:path/></v:shape></w:pict></w:r></w:p></w:hdr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert!(parsed.watermarks().is_empty());
        let serialized = parsed.to_xml_header().unwrap();
        assert!(
            serialized
                .windows(b"id=\"ordinary\"".len())
                .any(|w| w == b"id=\"ordinary\"")
        );
    }

    #[test]
    fn authored_watermark_patch_preserves_complete_header_bytes() {
        let closing = b"</q:hdr>";
        let source = format!(
            r#"<?xml version="1.0"?><q:hdr xmlns:q="{W_NS}" xmlns:mc="urn:mc" xmlns:x="urn:producer" mc:Ignorable="x" x:root="kept"><q:tbl><q:tr><q:tc><q:p/></q:tc></q:tr></q:tbl><q:sdt><q:sdtPr/><q:sdtContent><q:p/></q:sdtContent></q:sdt><q:p><q:r xmlns:v="{VML_NS}"><q:pict><v:shape id="producer"><v:path/></v:shape></q:pict></q:r></q:p>{}</q:hdr>"#,
            "tail"
        );
        let watermark = VmlWatermark::Text {
            text: "DRAFT".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };
        let updated = replace_authored_watermark(source.as_bytes(), &watermark).unwrap();
        let prefix = &source.as_bytes()[..source.len() - closing.len()];
        assert_eq!(&updated[..prefix.len()], prefix);
        assert_eq!(&updated[updated.len() - closing.len()..], closing);
        assert_eq!(
            CT_HdrFtr::from_xml(&updated).unwrap().watermarks(),
            &[watermark]
        );
    }

    #[test]
    fn authored_watermark_ownership_requires_the_vml_shape_id_attribute() {
        let unrelated = br#"<x:data>id=&quot;rdocx-watermark&quot;</x:data>"#;
        let source = format!(
            r#"<w:hdr xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:p><w:r>{}</w:r></w:p></w:hdr>"#,
            String::from_utf8_lossy(unrelated)
        );
        let watermark = VmlWatermark::Text {
            text: "FINAL".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };
        let updated = replace_authored_watermark(source.as_bytes(), &watermark).unwrap();
        assert!(
            updated
                .windows(unrelated.len())
                .any(|window| window == unrelated)
        );
        let replaced = replace_authored_watermark(&updated, &watermark).unwrap();
        assert!(
            replaced
                .windows(unrelated.len())
                .any(|window| window == unrelated)
        );
        assert_eq!(
            CT_HdrFtr::from_xml(&replaced).unwrap().watermarks(),
            &[watermark]
        );
    }

    #[test]
    fn converting_a_watermark_keeps_an_outgoing_type_shared_by_a_word_object() {
        let carrier = format!(
            r##"<w:object><v:shape id="rdocx-watermark" type="#{TEXT_WATERMARK_SHAPE_TYPE_ID}"><v:textpath string="PRODUCER"/></v:shape></w:object>"##
        );
        let source = format!(
            r##"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}" xmlns:o="{OFFICE_NS}"><w:p><w:r><w:pict><v:shapetype id="{TEXT_WATERMARK_SHAPE_TYPE_ID}" o:spt="136"><v:path/></v:shapetype><v:shape id="rdocx-watermark" type="#{TEXT_WATERMARK_SHAPE_TYPE_ID}" style="width:468pt;height:117pt"><v:textpath string="API"/></v:shape></w:pict>{carrier}</w:r></w:p></w:hdr>"##
        );
        let image = VmlWatermark::Image {
            relationship_id: "rId7".to_owned(),
            width_pt: 72.0,
            height_pt: 36.0,
            rotation_degrees: 0.0,
            opacity: 0.5,
        };
        let replaced = replace_authored_watermark(source.as_bytes(), &image).unwrap();
        assert!(
            replaced
                .windows(carrier.len())
                .any(|window| window == carrier.as_bytes())
        );
        let replaced = String::from_utf8(replaced).unwrap();
        assert_eq!(replaced.matches(r#"id="rdocx-watermark""#).count(), 2);
        assert_eq!(
            replaced
                .matches(&format!(r#"id="{TEXT_WATERMARK_SHAPE_TYPE_ID}""#))
                .count(),
            1
        );
        assert_eq!(
            replaced
                .matches(&format!(r#"id="{IMAGE_WATERMARK_SHAPE_TYPE_ID}""#))
                .count(),
            1
        );
        assert!(replaced.contains(&format!(r##"type="#{TEXT_WATERMARK_SHAPE_TYPE_ID}""##)));
        assert!(replaced.contains(&format!(r##"type="#{IMAGE_WATERMARK_SHAPE_TYPE_ID}""##)));
    }

    #[test]
    fn owned_watermarks_with_missing_or_malformed_types_gain_the_generated_type() {
        let watermark = VmlWatermark::Text {
            text: "FINAL".to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        };
        for type_attribute in ["", r#" type="malformed""#] {
            let source = format!(
                r#"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}"><w:p><w:r><w:pict><v:shape id="rdocx-watermark"{type_attribute} style="width:1pt;height:1pt"><v:textpath string="OLD"/></v:shape></w:pict></w:r></w:p></w:hdr>"#
            );
            let replaced = replace_authored_watermark(source.as_bytes(), &watermark).unwrap();
            let xml = String::from_utf8(replaced.clone()).unwrap();
            assert_eq!(
                xml.matches(&format!(r#"id="{TEXT_WATERMARK_SHAPE_TYPE_ID}""#))
                    .count(),
                1,
                "{xml}"
            );
            assert!(xml.contains(&format!(r##"type="#{TEXT_WATERMARK_SHAPE_TYPE_ID}""##)));
            assert_eq!(
                CT_HdrFtr::from_xml(&replaced).unwrap().watermarks(),
                std::slice::from_ref(&watermark)
            );
            assert_eq!(
                replace_authored_watermark(&replaced, &watermark).unwrap(),
                replaced
            );
        }
    }

    #[test]
    fn foreign_same_local_end_tags_do_not_terminate_vml_projection() {
        let xml = format!(
            r#"<w:hdr xmlns:w="{W_NS}" xmlns:v="{VML_NS}" xmlns:x="urn:producer"><w:p><w:r><w:pict><v:shape style="width:468pt;height:117pt"><x:shape><x:pict/></x:shape><x:pict></x:pict><v:textpath string="DRAFT"/></v:shape></w:pict></w:r></w:p></w:hdr>"#
        );
        let parsed = CT_HdrFtr::from_xml(xml.as_bytes()).unwrap();
        assert!(matches!(
            parsed.watermarks(),
            [VmlWatermark::Text { text, .. }] if text == "DRAFT"
        ));
    }
}
